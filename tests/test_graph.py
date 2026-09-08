from __future__ import annotations

import httpx
import pytest
from azure.core.credentials import AccessToken

from mcp_sharepoint.config import Settings
from mcp_sharepoint.graph import GraphClient, GraphError


def settings() -> Settings:
    return Settings(
        tenant_id="tenant",
        client_id="client",
        client_secret="secret",
        site_urls=("https://example.sharepoint.com/sites/Team",),
        search_region="EUR",
        graph_max_retries=0,
    )


class FakeCredential:
    async def get_token(self, *scopes, **kwargs):
        return AccessToken("test-token", 4_102_444_800)

    async def close(self):
        return None


@pytest.mark.asyncio
async def test_json_request_adds_application_token() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        assert request.headers["authorization"] == "Bearer test-token"
        return httpx.Response(200, json={"value": []})

    http = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    downloads = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    try:
        graph = GraphClient(
            settings(),
            credential=FakeCredential(),  # type: ignore[arg-type]
            http_client=http,
            download_client=downloads,
        )
        assert await graph.get_json("/sites") == {"value": []}
    finally:
        await http.aclose()
        await downloads.aclose()


@pytest.mark.asyncio
async def test_download_redirect_does_not_forward_graph_token() -> None:
    def graph_handler(request: httpx.Request) -> httpx.Response:
        assert request.headers["authorization"] == "Bearer test-token"
        return httpx.Response(
            302,
            headers={"location": "https://download.example.test/document"},
        )

    def download_handler(request: httpx.Request) -> httpx.Response:
        assert "authorization" not in request.headers
        return httpx.Response(200, content=b"document bytes")

    http = httpx.AsyncClient(transport=httpx.MockTransport(graph_handler))
    downloads = httpx.AsyncClient(transport=httpx.MockTransport(download_handler))
    try:
        graph = GraphClient(
            settings(),
            credential=FakeCredential(),  # type: ignore[arg-type]
            http_client=http,
            download_client=downloads,
        )
        async with graph.stream_drive_item_content("drive", "item") as response:
            assert await response.aread() == b"document bytes"
    finally:
        await http.aclose()
        await downloads.aclose()


def test_untrusted_pagination_url_is_rejected() -> None:
    with pytest.raises(GraphError, match="untrusted pagination URL"):
        GraphClient._graph_url("https://attacker.example/v1.0/sites")
