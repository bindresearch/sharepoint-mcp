"""Small asynchronous Microsoft Graph client with retry handling."""

from __future__ import annotations

import asyncio
import email.utils
import json
import random
from collections.abc import AsyncIterator, Mapping
from contextlib import asynccontextmanager
from datetime import datetime, timezone
from typing import Any
from urllib.parse import quote, urljoin, urlsplit

import httpx
from azure.core.credentials import AccessToken
from azure.identity.aio import ClientSecretCredential

from .config import Settings

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"
_RETRYABLE_STATUS_CODES = {429, 500, 502, 503, 504}
_REDIRECT_STATUS_CODES = {301, 302, 303, 307, 308}


class GraphError(RuntimeError):
    """A sanitized error returned by Microsoft Graph or a download host."""

    def __init__(
        self,
        message: str,
        *,
        status_code: int | None = None,
        code: str | None = None,
        request_id: str | None = None,
    ) -> None:
        super().__init__(message)
        self.status_code = status_code
        self.code = code
        self.request_id = request_id


class GraphClient:
    """Microsoft Graph client that uses an Entra application identity."""

    def __init__(
        self,
        settings: Settings,
        *,
        credential: ClientSecretCredential | None = None,
        http_client: httpx.AsyncClient | None = None,
        download_client: httpx.AsyncClient | None = None,
    ) -> None:
        self._settings = settings
        self._credential = credential or ClientSecretCredential(
            tenant_id=settings.tenant_id,
            client_id=settings.client_id,
            client_secret=settings.client_secret.get_secret_value(),
        )
        timeout = httpx.Timeout(settings.graph_request_timeout_seconds)
        self._http = http_client or httpx.AsyncClient(timeout=timeout)
        self._downloads = download_client or httpx.AsyncClient(
            timeout=timeout, follow_redirects=True
        )
        self._owns_credential = credential is None
        self._owns_http = http_client is None
        self._owns_download_client = download_client is None

    async def __aenter__(self) -> GraphClient:
        return self

    async def __aexit__(self, *_: object) -> None:
        await self.close()

    async def close(self) -> None:
        if self._owns_http:
            await self._http.aclose()
        if self._owns_download_client:
            await self._downloads.aclose()
        if self._owns_credential:
            await self._credential.close()

    async def get_json(
        self,
        path_or_url: str,
        *,
        params: Mapping[str, str | int] | None = None,
    ) -> dict[str, Any]:
        return await self.request_json("GET", path_or_url, params=params)

    async def request_json(
        self,
        method: str,
        path_or_url: str,
        *,
        params: Mapping[str, str | int] | None = None,
        json_body: Mapping[str, Any] | None = None,
    ) -> dict[str, Any]:
        url = self._graph_url(path_or_url)
        last_error: GraphError | None = None

        for attempt in range(self._settings.graph_max_retries + 1):
            token = await self._credential.get_token(GRAPH_SCOPE)
            try:
                response = await self._http.request(
                    method,
                    url,
                    params=params,
                    json=json_body,
                    headers=self._headers(token),
                )
            except httpx.RequestError as exc:
                last_error = GraphError(f"Microsoft Graph request failed: {exc}")
                if attempt >= self._settings.graph_max_retries:
                    raise last_error from exc
                await self._sleep_before_retry(attempt, None)
                continue

            if response.status_code in _RETRYABLE_STATUS_CODES:
                last_error = await self._error_from_response(response)
                if attempt < self._settings.graph_max_retries:
                    await self._sleep_before_retry(attempt, response)
                    continue

            if response.is_error:
                raise await self._error_from_response(response)

            try:
                payload = response.json()
            except (json.JSONDecodeError, UnicodeDecodeError) as exc:
                raise GraphError(
                    "Microsoft Graph returned an invalid JSON response",
                    status_code=response.status_code,
                    request_id=response.headers.get("request-id"),
                ) from exc
            if not isinstance(payload, dict):
                raise GraphError("Microsoft Graph returned an unexpected response")
            return payload

        assert last_error is not None  # The loop always sets this before exhaustion.
        raise last_error

    @asynccontextmanager
    async def stream_drive_item_content(
        self, drive_id: str, item_id: str
    ) -> AsyncIterator[httpx.Response]:
        """Stream an item while keeping the Graph bearer token off download hosts."""

        graph_url = self._graph_url(
            f"/drives/{quote(drive_id, safe='')}/items/"
            f"{quote(item_id, safe='')}/content"
        )
        last_error: GraphError | None = None

        for attempt in range(self._settings.graph_max_retries + 1):
            token = await self._credential.get_token(GRAPH_SCOPE)
            graph_response: httpx.Response | None = None
            content_response: httpx.Response | None = None
            try:
                request = self._http.build_request(
                    "GET", graph_url, headers=self._headers(token)
                )
                graph_response = await self._http.send(
                    request, stream=True, follow_redirects=False
                )

                if graph_response.status_code in _REDIRECT_STATUS_CODES:
                    location = graph_response.headers.get("location")
                    await graph_response.aclose()
                    graph_response = None
                    if not location:
                        raise GraphError(
                            "Microsoft Graph returned a download redirect without a location"
                        )
                    download_url = urljoin(graph_url, location)
                    parsed_download = urlsplit(download_url)
                    if (
                        parsed_download.scheme.lower() != "https"
                        or not parsed_download.hostname
                    ):
                        raise GraphError(
                            "Microsoft Graph returned an unsafe document download URL"
                        )
                    download_request = self._downloads.build_request(
                        "GET", download_url
                    )
                    content_response = await self._downloads.send(
                        download_request, stream=True, follow_redirects=True
                    )
                else:
                    content_response = graph_response
                    graph_response = None

                if content_response.status_code in _RETRYABLE_STATUS_CODES:
                    await content_response.aread()
                    last_error = await self._error_from_response(content_response)
                    if attempt < self._settings.graph_max_retries:
                        retry_response = content_response
                        await content_response.aclose()
                        content_response = None
                        await self._sleep_before_retry(attempt, retry_response)
                        continue

                if content_response.is_error:
                    await content_response.aread()
                    raise await self._error_from_response(content_response)

                try:
                    yield content_response
                finally:
                    await content_response.aclose()
                return
            except httpx.RequestError as exc:
                last_error = GraphError(f"Document download failed: {exc}")
                if attempt >= self._settings.graph_max_retries:
                    raise last_error from exc
                await self._sleep_before_retry(attempt, None)
            finally:
                if graph_response is not None:
                    await graph_response.aclose()
                if content_response is not None and not content_response.is_closed:
                    await content_response.aclose()

        assert last_error is not None
        raise last_error

    @staticmethod
    def _headers(token: AccessToken) -> dict[str, str]:
        return {
            "Authorization": f"Bearer {token.token}",
            "Accept": "application/json",
        }

    @staticmethod
    def _graph_url(path_or_url: str) -> str:
        if path_or_url.startswith("/"):
            return f"{GRAPH_ROOT}{path_or_url}"

        parsed = urlsplit(path_or_url)
        if (
            parsed.scheme.lower() != "https"
            or parsed.hostname != "graph.microsoft.com"
            or not parsed.path.startswith("/v1.0/")
        ):
            raise GraphError("Refusing to request an untrusted pagination URL")
        return path_or_url

    async def _sleep_before_retry(
        self, attempt: int, response: httpx.Response | None
    ) -> None:
        retry_after = self._retry_after_seconds(response)
        delay = retry_after if retry_after is not None else min(2**attempt, 30)
        await asyncio.sleep(delay + random.uniform(0, 0.25))

    @staticmethod
    def _retry_after_seconds(response: httpx.Response | None) -> float | None:
        if response is None:
            return None
        value = response.headers.get("retry-after")
        if not value:
            return None
        try:
            return max(0.0, min(float(value), 120.0))
        except ValueError:
            try:
                retry_at = email.utils.parsedate_to_datetime(value)
                if retry_at.tzinfo is None:
                    retry_at = retry_at.replace(tzinfo=timezone.utc)
                return max(
                    0.0,
                    min((retry_at - datetime.now(timezone.utc)).total_seconds(), 120.0),
                )
            except (TypeError, ValueError, OverflowError):
                return None

    @staticmethod
    async def _error_from_response(response: httpx.Response) -> GraphError:
        code: str | None = None
        message = f"Microsoft Graph returned HTTP {response.status_code}"
        try:
            body = response.json()
            error = body.get("error", {}) if isinstance(body, dict) else {}
            if isinstance(error, dict):
                if isinstance(error.get("code"), str):
                    code = error["code"]
                if isinstance(error.get("message"), str):
                    message = error["message"]
        except (json.JSONDecodeError, UnicodeDecodeError):
            pass

        request_id = response.headers.get("request-id") or response.headers.get(
            "client-request-id"
        )
        suffix = f" (request ID: {request_id})" if request_id else ""
        return GraphError(
            f"{message}{suffix}",
            status_code=response.status_code,
            code=code,
            request_id=request_id,
        )
