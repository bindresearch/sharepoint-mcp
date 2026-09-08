from __future__ import annotations

import pytest

from mcp_sharepoint.config import Settings
from mcp_sharepoint.server import create_server


@pytest.mark.asyncio
async def test_server_registers_public_tool_arguments_only() -> None:
    settings = Settings(
        tenant_id="tenant",
        client_id="client",
        client_secret="secret",
        site_urls=("https://example.sharepoint.com/sites/Team",),
        search_region="EUR",
    )

    tools = {tool.name: tool for tool in await create_server(settings).list_tools()}

    assert set(tools) == {
        "search_documents",
        "get_document_metadata",
        "read_document",
        "list_sources",
    }
    assert "context" not in tools["search_documents"].inputSchema["properties"]
    assert tools["search_documents"].outputSchema is not None
    assert tools["search_documents"].annotations is not None
    assert tools["search_documents"].annotations.readOnlyHint is True
