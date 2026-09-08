from __future__ import annotations

from contextlib import asynccontextmanager

import httpx
import pytest

from mcp_sharepoint.config import Settings
from mcp_sharepoint.references import DocumentReference
from mcp_sharepoint.service import SharePointService
from mcp_sharepoint.sources import DriveSource, SiteSource, SourceRegistry


def settings() -> Settings:
    return Settings(
        tenant_id="tenant",
        client_id="client",
        client_secret="secret",
        site_urls=("https://example.sharepoint.com/sites/Team",),
        search_region="EUR",
        max_download_mb=1,
        max_section_characters=1_000,
        max_response_characters=5_000,
    )


def sources() -> SourceRegistry:
    drive = DriveSource(
        id="drive-1",
        name="Documents",
        web_url="https://example.sharepoint.com/sites/Team/Shared Documents",
        site_id="site-1",
        site_name="Team",
        site_web_url="https://example.sharepoint.com/sites/Team",
    )
    site = SiteSource(
        id="site-1",
        name="Team",
        web_url="https://example.sharepoint.com/sites/Team",
        drives=(drive,),
    )
    return SourceRegistry((site,))


class FakeGraph:
    def __init__(self) -> None:
        self.request_body = None

    async def request_json(self, method, path, *, json_body=None, **kwargs):
        self.request_body = json_body
        # Microsoft Search uses hitId as the driveItem ID and does not always
        # include the id or file facet in the default resource projection.
        search_resource = self.item()
        search_resource.pop("id")
        search_resource.pop("file")
        return {
            "value": [
                {
                    "hitsContainers": [
                        {
                            "total": 1,
                            "moreResultsAvailable": False,
                            "hits": [
                                {
                                    "hitId": "item-1",
                                    "rank": 1,
                                    "summary": "A <c0>matching</c0> report <ddd/>",
                                    "resource": search_resource,
                                }
                            ],
                        }
                    ]
                }
            ]
        }

    async def get_json(self, path, *, params=None):
        return self.item()

    @asynccontextmanager
    async def stream_drive_item_content(self, drive_id, item_id):
        assert drive_id == "drive-1"
        assert item_id == "item-1"
        yield httpx.Response(200, content=b"First line\nSecond line")

    @staticmethod
    def item():
        return {
            "id": "item-1",
            "name": "report.txt",
            "webUrl": "https://example.sharepoint.com/sites/Team/Shared Documents/report.txt",
            "size": 22,
            "file": {"mimeType": "text/plain"},
            "parentReference": {"driveId": "drive-1"},
            "createdBy": {
                "user": {
                    "id": "person-1",
                    "displayName": "Ada Lovelace",
                    "email": "ada@example.com",
                }
            },
            "createdDateTime": "2025-01-01T10:00:00Z",
            "lastModifiedDateTime": "2025-01-02T10:00:00Z",
            "eTag": '"etag-1"',
        }


@pytest.mark.asyncio
async def test_search_scopes_query_and_returns_reference() -> None:
    graph = FakeGraph()
    service = SharePointService(graph, sources(), settings())

    result = await service.search_documents("annual report", file_types=["txt"])

    request = graph.request_body["requests"][0]
    assert (
        'Path:"https://example.sharepoint.com/sites/Team"'
        in request["query"]["queryString"]
    )
    assert "filetype:txt" in request["query"]["queryString"]
    assert "isDocument=true" in request["query"]["queryString"]
    assert result["results"][0]["summary"] == "A matching report …"
    reference = DocumentReference.decode(result["results"][0]["document_ref"])
    assert reference == DocumentReference("drive-1", "item-1")


@pytest.mark.asyncio
async def test_read_document_downloads_and_extracts_text() -> None:
    graph = FakeGraph()
    service = SharePointService(graph, sources(), settings())
    document_ref = DocumentReference("drive-1", "item-1").encode()

    result = await service.read_document(document_ref)

    assert result["sections"][0]["text"] == "First line\nSecond line"
    assert result["document"]["created_by"]["display_name"] == "Ada Lovelace"
    assert result["next_section"] is None


@pytest.mark.asyncio
async def test_document_reference_cannot_escape_configured_drives() -> None:
    service = SharePointService(FakeGraph(), sources(), settings())
    forged_ref = DocumentReference("other-drive", "item-1").encode()

    with pytest.raises(ValueError, match="configured SharePoint site"):
        await service.get_document_metadata(forged_ref)
