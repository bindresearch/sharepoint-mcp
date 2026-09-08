from __future__ import annotations

import pytest

from mcp_sharepoint.sources import DriveSource, SiteSource, SourceRegistry


class FakeGraph:
    def __init__(self) -> None:
        self.paths: list[str] = []

    async def get_json(self, path: str, *, params=None):
        self.paths.append(path)
        if "/drives" in path:
            return {
                "value": [
                    {
                        "id": "drive-1",
                        "name": "Documents",
                        "webUrl": "https://example.sharepoint.com/sites/Team/Shared Documents",
                        "driveType": "documentLibrary",
                    }
                ]
            }
        return {
            "id": "example.sharepoint.com,site-id,web-id",
            "displayName": "Team",
            "webUrl": "https://example.sharepoint.com/sites/Team",
        }


@pytest.mark.asyncio
async def test_source_registry_resolves_all_site_drives() -> None:
    graph = FakeGraph()

    registry = await SourceRegistry.load(
        graph, ("https://example.sharepoint.com/sites/Team",)
    )

    assert list(registry.drives) == ["drive-1"]
    assert registry.require_drive("drive-1").site_name == "Team"
    assert graph.paths[0] == "/sites/example.sharepoint.com:/sites/Team"


def test_drive_for_web_url_observes_path_boundary() -> None:
    drive = DriveSource(
        id="drive-1",
        name="Documents",
        web_url="https://example.sharepoint.com/sites/Team",
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
    registry = SourceRegistry((site,))

    assert (
        registry.drive_for_web_url(
            "https://example.sharepoint.com/sites/Team/report.docx"
        )
        == drive
    )
    assert (
        registry.drive_for_web_url(
            "https://example.sharepoint.com/sites/TeamOther/file.docx"
        )
        is None
    )
