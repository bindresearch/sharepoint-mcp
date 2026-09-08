"""Resolve configured SharePoint sites and their document libraries."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any
from urllib.parse import quote, unquote, urlsplit

from .graph import GraphClient, GraphError


@dataclass(frozen=True, slots=True)
class DriveSource:
    id: str
    name: str
    web_url: str
    site_id: str
    site_name: str
    site_web_url: str


@dataclass(frozen=True, slots=True)
class SiteSource:
    id: str
    name: str
    web_url: str
    drives: tuple[DriveSource, ...]


class SourceRegistry:
    """An in-memory registry built from SHAREPOINT_SITE_URLS at startup."""

    def __init__(self, sites: tuple[SiteSource, ...]) -> None:
        self.sites = sites
        self.drives = {drive.id: drive for site in sites for drive in site.drives}

    @classmethod
    async def load(
        cls, graph: GraphClient, configured_site_urls: tuple[str, ...]
    ) -> SourceRegistry:
        sites: list[SiteSource] = []
        seen_site_ids: set[str] = set()

        for configured_url in configured_site_urls:
            site_data = await graph.get_json(cls._site_lookup_path(configured_url))
            site_id = cls._required_string(site_data, "id", "site")
            if site_id in seen_site_ids:
                continue
            seen_site_ids.add(site_id)

            site_name = cls._display_name(site_data, configured_url)
            site_web_url = cls._required_string(site_data, "webUrl", "site")
            drives = await cls._load_drives(
                graph,
                site_id=site_id,
                site_name=site_name,
                site_web_url=site_web_url,
            )
            if not drives:
                raise GraphError(
                    f"Configured SharePoint site {site_web_url!r} has no document libraries"
                )
            sites.append(
                SiteSource(
                    id=site_id,
                    name=site_name,
                    web_url=site_web_url.rstrip("/"),
                    drives=drives,
                )
            )

        return cls(tuple(sites))

    def require_drive(self, drive_id: str) -> DriveSource:
        """Return a configured drive or reject a forged document reference."""

        try:
            return self.drives[drive_id]
        except KeyError as exc:
            raise ValueError(
                "The document is not in a configured SharePoint site"
            ) from exc

    def drive_for_web_url(self, web_url: str) -> DriveSource | None:
        """Find the most specific configured drive containing a web URL."""

        matching = [
            drive
            for drive in self.drives.values()
            if self._url_is_within(web_url, drive.web_url)
        ]
        if not matching:
            return None
        return max(matching, key=lambda drive: len(urlsplit(drive.web_url).path))

    def as_dicts(self) -> list[dict[str, Any]]:
        return [
            {
                "name": site.name,
                "web_url": site.web_url,
                "libraries": [
                    {"name": drive.name, "web_url": drive.web_url}
                    for drive in site.drives
                ],
            }
            for site in self.sites
        ]

    @staticmethod
    def _site_lookup_path(site_url: str) -> str:
        parsed = urlsplit(site_url)
        hostname = parsed.hostname
        if hostname is None:  # Settings validation prevents this.
            raise ValueError("SharePoint site URL has no hostname")
        relative_path = unquote(parsed.path.rstrip("/") or "/")
        return f"/sites/{quote(hostname, safe='')}:{quote(relative_path, safe='/')}"

    @classmethod
    async def _load_drives(
        cls,
        graph: GraphClient,
        *,
        site_id: str,
        site_name: str,
        site_web_url: str,
    ) -> tuple[DriveSource, ...]:
        encoded_site_id = quote(site_id, safe="")
        next_url: str | None = f"/sites/{encoded_site_id}/drives"
        params: dict[str, str] | None = {"$select": "id,name,webUrl,driveType"}
        drives: list[DriveSource] = []

        while next_url:
            payload = await graph.get_json(next_url, params=params)
            params = None
            values = payload.get("value", [])
            if not isinstance(values, list):
                raise GraphError("Microsoft Graph returned an invalid drives response")

            for value in values:
                if not isinstance(value, dict):
                    continue
                drive_id = value.get("id")
                drive_name = value.get("name")
                drive_web_url = value.get("webUrl")
                if not isinstance(drive_id, str) or not drive_id:
                    continue
                if not isinstance(drive_name, str) or not drive_name:
                    continue
                if not isinstance(drive_web_url, str) or not drive_web_url:
                    continue
                drives.append(
                    DriveSource(
                        id=drive_id,
                        name=drive_name,
                        web_url=drive_web_url.rstrip("/"),
                        site_id=site_id,
                        site_name=site_name,
                        site_web_url=site_web_url.rstrip("/"),
                    )
                )

            raw_next_url = payload.get("@odata.nextLink")
            next_url = raw_next_url if isinstance(raw_next_url, str) else None

        return tuple(drives)

    @staticmethod
    def _required_string(data: dict[str, Any], key: str, resource: str) -> str:
        value = data.get(key)
        if not isinstance(value, str) or not value:
            raise GraphError(
                f"Microsoft Graph returned a {resource} without a valid {key}"
            )
        return value

    @staticmethod
    def _display_name(site_data: dict[str, Any], fallback: str) -> str:
        for key in ("displayName", "name"):
            value = site_data.get(key)
            if isinstance(value, str) and value:
                return value
        return fallback

    @staticmethod
    def _url_is_within(candidate: str, root: str) -> bool:
        candidate_url = urlsplit(candidate)
        root_url = urlsplit(root)
        if (
            candidate_url.scheme.lower() != root_url.scheme.lower()
            or candidate_url.netloc.lower() != root_url.netloc.lower()
        ):
            return False
        candidate_path = unquote(candidate_url.path).rstrip("/").casefold()
        root_path = unquote(root_url.path).rstrip("/").casefold()
        return candidate_path == root_path or candidate_path.startswith(f"{root_path}/")
