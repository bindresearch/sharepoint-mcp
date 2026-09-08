"""SharePoint search, metadata, and text retrieval services."""

from __future__ import annotations

import asyncio
import html
import re
import tempfile
from datetime import date
from typing import Any, BinaryIO, cast
from urllib.parse import quote

from .config import Settings
from .extractors import (
    SUPPORTED_EXTENSIONS,
    ExtractionError,
    extension_for,
    extract_document,
)
from .graph import GraphClient, GraphError
from .references import DocumentReference
from .sources import DriveSource, SourceRegistry

_HIGHLIGHT_TAG = re.compile(r"</?c\d+>", flags=re.IGNORECASE)
_OMITTED_TEXT_TAG = re.compile(r"<ddd\s*/>", flags=re.IGNORECASE)
_SEARCH_TERM = re.compile(r"[\w.@+\-]+", flags=re.UNICODE)
_MEBIBYTE = 1024 * 1024


class SharePointService:
    """Stateless operations used by the MCP tools."""

    def __init__(
        self,
        graph: GraphClient,
        sources: SourceRegistry,
        settings: Settings,
    ) -> None:
        self.graph = graph
        self.sources = sources
        self.settings = settings

    async def search_documents(
        self,
        query: str,
        *,
        file_types: list[str] | None = None,
        created_after: str | None = None,
        modified_after: str | None = None,
        limit: int = 20,
        offset: int = 0,
    ) -> dict[str, Any]:
        if (
            not isinstance(limit, int)
            or not 1 <= limit <= self.settings.max_search_results
        ):
            raise ValueError(
                f"limit must be between 1 and {self.settings.max_search_results}"
            )
        if not isinstance(offset, int) or offset < 0:
            raise ValueError("offset must be zero or greater")

        extensions = self._normalize_file_types(file_types)
        query_string = self._build_kql(
            query,
            extensions=extensions,
            created_after=created_after,
            modified_after=modified_after,
        )
        payload = await self.graph.request_json(
            "POST",
            "/search/query",
            json_body={
                "requests": [
                    {
                        "entityTypes": ["driveItem"],
                        "query": {"queryString": query_string},
                        "from": offset,
                        "size": limit,
                        "region": self.settings.search_region,
                    }
                ]
            },
        )

        hits, total, more_results = self._search_hits(payload)
        results: list[dict[str, Any]] = []
        for hit in hits:
            parsed = self._parse_search_hit(hit, extensions)
            if parsed is not None:
                results.append(parsed)

        return {
            "query": query,
            "offset": offset,
            "limit": limit,
            "total_estimate": total,
            "more_results_available": more_results,
            "next_offset": offset + limit if more_results else None,
            "results": results,
        }

    async def get_document_metadata(self, document_ref: str) -> dict[str, Any]:
        reference, drive = self._validated_reference(document_ref)
        item = await self._get_drive_item(reference)
        return self._metadata(item, drive, document_ref=document_ref)

    async def read_document(
        self,
        document_ref: str,
        *,
        start_section: int = 0,
        section_count: int = 10,
        max_characters: int | None = None,
    ) -> dict[str, Any]:
        if not isinstance(start_section, int) or start_section < 0:
            raise ValueError("start_section must be zero or greater")
        if not isinstance(section_count, int) or not 1 <= section_count <= 50:
            raise ValueError("section_count must be between 1 and 50")

        response_character_limit = self.settings.max_response_characters
        if max_characters is not None:
            if (
                not isinstance(max_characters, int)
                or max_characters < self.settings.max_section_characters
            ):
                raise ValueError(
                    "max_characters must be at least MAX_SECTION_CHARACTERS "
                    f"({self.settings.max_section_characters})"
                )
            response_character_limit = min(
                max_characters, self.settings.max_response_characters
            )

        reference, drive = self._validated_reference(document_ref)
        item = await self._get_drive_item(reference)
        filename = self._required_string(item, "name", "drive item")
        extension = extension_for(filename)
        if extension not in SUPPORTED_EXTENSIONS:
            raise ValueError(f"Unsupported file type: .{extension or '(none)'}")
        if not isinstance(item.get("file"), dict):
            raise ValueError("The document reference identifies a folder, not a file")

        max_download_bytes = self.settings.max_download_mb * _MEBIBYTE
        size = item.get("size")
        if isinstance(size, int) and size > max_download_bytes:
            raise ValueError(
                f"Document is too large to read ({size} bytes; "
                f"limit is {max_download_bytes} bytes)"
            )

        downloaded_bytes = 0
        with tempfile.SpooledTemporaryFile(
            max_size=self.settings.spool_threshold_mb * _MEBIBYTE, mode="w+b"
        ) as spool:
            async with self.graph.stream_drive_item_content(
                reference.drive_id, reference.item_id
            ) as response:
                content_length = self._content_length(response)
                if content_length is not None and content_length > max_download_bytes:
                    raise ValueError(
                        f"Document is too large to read ({content_length} bytes; "
                        f"limit is {max_download_bytes} bytes)"
                    )
                async for chunk in response.aiter_bytes():
                    downloaded_bytes += len(chunk)
                    if downloaded_bytes > max_download_bytes:
                        raise ValueError(
                            "Document exceeded MAX_DOWNLOAD_MB while it was downloaded"
                        )
                    spool.write(chunk)

            spool.seek(0)
            try:
                extraction = await asyncio.wait_for(
                    asyncio.to_thread(
                        extract_document,
                        cast(BinaryIO, spool),
                        filename,
                        section_character_limit=self.settings.max_section_characters,
                        document_character_limit=self.settings.max_extracted_characters,
                    ),
                    timeout=self.settings.extraction_timeout_seconds,
                )
            except TimeoutError as exc:
                raise ExtractionError(
                    f"Text extraction exceeded {self.settings.extraction_timeout_seconds} seconds"
                ) from exc

        sections = extraction.sections
        if start_section > len(sections):
            raise ValueError(
                f"start_section is beyond the document's {len(sections)} sections"
            )

        selected: list[dict[str, int | str]] = []
        characters = 0
        stop = min(start_section + section_count, len(sections))
        next_section = start_section
        for section in sections[start_section:stop]:
            added_characters = len(section.text)
            if selected and characters + added_characters > response_character_limit:
                break
            if not selected and added_characters > response_character_limit:
                # Sections normally respect this limit. This is a final guard if
                # configuration changes between extraction and response assembly.
                selected.append(
                    {
                        "index": section.index,
                        "label": section.label,
                        "text": section.text[:response_character_limit],
                    }
                )
                characters = response_character_limit
                next_section = section.index + 1
                break
            selected.append(section.as_dict())
            characters += added_characters
            next_section = section.index + 1

        has_more = next_section < len(sections)
        return {
            "document": self._metadata(item, drive, document_ref=document_ref),
            "sections": selected,
            "start_section": start_section,
            "next_section": next_section if has_more else None,
            "total_sections": len(sections),
            "characters_returned": characters,
            "content_truncated_during_extraction": extraction.truncated,
            "note": (
                "The server is stateless; a continuation request downloads and parses "
                "the document again."
                if has_more
                else None
            ),
        }

    def list_sources(self) -> dict[str, Any]:
        return {"sites": self.sources.as_dicts()}

    def _validated_reference(
        self, document_ref: str
    ) -> tuple[DocumentReference, DriveSource]:
        reference = DocumentReference.decode(document_ref)
        drive = self.sources.require_drive(reference.drive_id)
        return reference, drive

    async def _get_drive_item(self, reference: DocumentReference) -> dict[str, Any]:
        drive_id = quote(reference.drive_id, safe="")
        item_id = quote(reference.item_id, safe="")
        item = await self.graph.get_json(
            f"/drives/{drive_id}/items/{item_id}",
            params={
                "$select": (
                    "id,name,webUrl,size,createdBy,createdDateTime,lastModifiedBy,"
                    "lastModifiedDateTime,parentReference,file,eTag"
                )
            },
        )
        parent = item.get("parentReference")
        returned_drive_id = parent.get("driveId") if isinstance(parent, dict) else None
        if returned_drive_id is not None and returned_drive_id != reference.drive_id:
            raise GraphError(
                "Microsoft Graph returned an item from an unexpected drive"
            )
        return item

    def _build_kql(
        self,
        query: str,
        *,
        extensions: tuple[str, ...],
        created_after: str | None,
        modified_after: str | None,
    ) -> str:
        if not isinstance(query, str) or not query.strip():
            raise ValueError("query must not be empty")
        if len(query) > 500:
            raise ValueError("query must contain no more than 500 characters")

        terms = [
            term.strip(".")[:100]
            for term in _SEARCH_TERM.findall(query)
            if term.strip(".")
        ]
        if not terms:
            raise ValueError("query does not contain searchable characters")
        if len(terms) > 32:
            raise ValueError("query must contain no more than 32 terms")

        clauses = ["(" + " AND ".join(f'"{term}"' for term in terms) + ")"]
        path_clauses = " OR ".join(
            f'Path:"{site.web_url}"' for site in self.sources.sites
        )
        clauses.append(f"({path_clauses})")
        extension_clauses = " OR ".join(
            f"filetype:{extension}" for extension in extensions
        )
        clauses.append(f"({extension_clauses})")
        clauses.append("isDocument=true")

        if created_after is not None:
            clauses.append(f"Created>={self._iso_date(created_after, 'created_after')}")
        if modified_after is not None:
            clauses.append(
                f"LastModifiedTime>={self._iso_date(modified_after, 'modified_after')}"
            )
        return " AND ".join(clauses)

    @staticmethod
    def _iso_date(value: str, field_name: str) -> str:
        try:
            return date.fromisoformat(value).isoformat()
        except (TypeError, ValueError) as exc:
            raise ValueError(f"{field_name} must use YYYY-MM-DD format") from exc

    @staticmethod
    def _normalize_file_types(file_types: list[str] | None) -> tuple[str, ...]:
        if not file_types:
            return tuple(sorted(SUPPORTED_EXTENSIONS))
        normalized = tuple(
            dict.fromkeys(
                value.strip().lower().lstrip(".")
                for value in file_types
                if isinstance(value, str) and value.strip()
            )
        )
        if not normalized:
            raise ValueError("file_types must contain at least one file type")
        unsupported = sorted(set(normalized) - SUPPORTED_EXTENSIONS)
        if unsupported:
            raise ValueError(
                "Unsupported file types: "
                + ", ".join(f".{item}" for item in unsupported)
            )
        return normalized

    @staticmethod
    def _search_hits(
        payload: dict[str, Any],
    ) -> tuple[list[dict[str, Any]], int | None, bool]:
        values = payload.get("value")
        if not isinstance(values, list) or not values:
            return [], 0, False
        first_response = values[0]
        if not isinstance(first_response, dict):
            raise GraphError("Microsoft Search returned an invalid response")
        containers = first_response.get("hitsContainers")
        if not isinstance(containers, list) or not containers:
            return [], 0, False
        container = containers[0]
        if not isinstance(container, dict):
            raise GraphError("Microsoft Search returned an invalid hits container")
        raw_hits = container.get("hits", [])
        if not isinstance(raw_hits, list):
            raise GraphError("Microsoft Search returned an invalid hits list")
        hits = [hit for hit in raw_hits if isinstance(hit, dict)]
        total = container.get("total")
        parsed_total = total if isinstance(total, int) else None
        more = container.get("moreResultsAvailable")
        return hits, parsed_total, bool(more)

    def _parse_search_hit(
        self, hit: dict[str, Any], extensions: tuple[str, ...]
    ) -> dict[str, Any] | None:
        resource = hit.get("resource")
        if not isinstance(resource, dict):
            return None
        item_id = resource.get("id") or hit.get("hitId")
        name = resource.get("name")
        web_url = resource.get("webUrl")
        if not isinstance(item_id, str) or not item_id:
            return None
        if not isinstance(name, str) or not name:
            return None
        if not isinstance(web_url, str) or not web_url:
            return None
        if extension_for(name) not in extensions:
            return None

        parent = resource.get("parentReference")
        drive_id = parent.get("driveId") if isinstance(parent, dict) else None
        drive = self.sources.drives.get(drive_id) if isinstance(drive_id, str) else None
        if drive is None:
            drive = self.sources.drive_for_web_url(web_url)
        if drive is None:
            return None

        document_ref = DocumentReference(drive.id, item_id).encode()
        summary = hit.get("summary")
        file_data = resource.get("file")
        clean_summary = (
            " ".join(
                html.unescape(
                    _HIGHLIGHT_TAG.sub("", _OMITTED_TEXT_TAG.sub(" … ", summary))
                ).split()
            )
            if isinstance(summary, str)
            else None
        )
        return {
            "document_ref": document_ref,
            "name": name,
            "web_url": web_url,
            "site": drive.site_name,
            "library": drive.name,
            "size": resource.get("size"),
            "mime_type": (
                file_data.get("mimeType") if isinstance(file_data, dict) else None
            ),
            "created_by": self._identity(resource.get("createdBy")),
            "created_at": resource.get("createdDateTime"),
            "modified_by": self._identity(resource.get("lastModifiedBy")),
            "modified_at": resource.get("lastModifiedDateTime"),
            "summary": clean_summary,
            "rank": hit.get("rank"),
        }

    def _metadata(
        self,
        item: dict[str, Any],
        drive: DriveSource,
        *,
        document_ref: str,
    ) -> dict[str, Any]:
        file_data = item.get("file")
        return {
            "document_ref": document_ref,
            "name": item.get("name"),
            "web_url": item.get("webUrl"),
            "site": drive.site_name,
            "site_web_url": drive.site_web_url,
            "library": drive.name,
            "library_web_url": drive.web_url,
            "size": item.get("size"),
            "mime_type": file_data.get("mimeType")
            if isinstance(file_data, dict)
            else None,
            "created_by": self._identity(item.get("createdBy")),
            "created_at": item.get("createdDateTime"),
            "modified_by": self._identity(item.get("lastModifiedBy")),
            "modified_at": item.get("lastModifiedDateTime"),
            "etag": item.get("eTag"),
        }

    @staticmethod
    def _identity(value: object) -> dict[str, Any] | None:
        if not isinstance(value, dict):
            return None
        for identity_type in ("user", "application", "device"):
            identity = value.get(identity_type)
            if isinstance(identity, dict):
                return {
                    "type": identity_type,
                    "id": identity.get("id"),
                    "display_name": identity.get("displayName"),
                    "email": identity.get("email"),
                }
        return None

    @staticmethod
    def _required_string(data: dict[str, Any], key: str, resource: str) -> str:
        value = data.get(key)
        if not isinstance(value, str) or not value:
            raise GraphError(
                f"Microsoft Graph returned a {resource} without a valid {key}"
            )
        return value

    @staticmethod
    def _content_length(response: Any) -> int | None:
        value = response.headers.get("content-length")
        if value is None:
            return None
        try:
            return int(value)
        except (TypeError, ValueError):
            return None
