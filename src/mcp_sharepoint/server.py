"""MCP entry point and tool definitions."""

from __future__ import annotations

import argparse
import logging
import sys
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from dataclasses import dataclass
from typing import Any, Literal, cast

from mcp.server.fastmcp import Context, FastMCP
from mcp.types import ToolAnnotations
from pydantic import ValidationError

from .config import Settings
from .graph import GraphClient
from .service import SharePointService
from .sources import SourceRegistry

LOGGER = logging.getLogger("mcp_sharepoint")
Transport = Literal["stdio", "streamable-http"]
READ_ONLY_TOOL = ToolAnnotations(
    readOnlyHint=True,
    destructiveHint=False,
    idempotentHint=True,
    openWorldHint=True,
)


@dataclass(slots=True)
class ApplicationContext:
    service: SharePointService


def create_server(settings: Settings) -> FastMCP[ApplicationContext]:
    """Create an MCP server configured for the supplied environment."""

    @asynccontextmanager
    async def lifespan(_: FastMCP[Any]) -> AsyncIterator[ApplicationContext]:
        LOGGER.info(
            "Resolving %d configured SharePoint site(s)", len(settings.site_urls)
        )
        async with GraphClient(settings) as graph:
            sources = await SourceRegistry.load(
                graph, tuple(str(url).rstrip("/") for url in settings.site_urls)
            )
            LOGGER.info(
                "Loaded %d SharePoint site(s) and %d document library drive(s)",
                len(sources.sites),
                len(sources.drives),
            )
            yield ApplicationContext(
                service=SharePointService(graph, sources, settings)
            )

    mcp: FastMCP[ApplicationContext] = FastMCP(
        "SharePoint",
        instructions=(
            "Search configured employee SharePoint sites and retrieve current document "
            "metadata or extracted text. The server is read-only and stateless."
        ),
        host=settings.mcp_host,
        port=settings.mcp_port,
        log_level=cast(Any, settings.log_level),
        stateless_http=True,
        json_response=True,
        lifespan=lifespan,
    )

    def service(
        context: Context[Any, ApplicationContext, Any],
    ) -> SharePointService:
        return context.request_context.lifespan_context.service

    @mcp.tool(
        title="Search SharePoint documents",
        description=(
            "Full-text keyword search across all configured SharePoint sites and "
            "document libraries. Use document_ref from a result with the metadata "
            "and read tools. Dates use YYYY-MM-DD."
        ),
        annotations=READ_ONLY_TOOL,
        structured_output=True,
    )
    async def search_documents(
        query: str,
        context: Context[Any, ApplicationContext, Any],
        file_types: list[str] | None = None,
        created_after: str | None = None,
        modified_after: str | None = None,
        limit: int = 20,
        offset: int = 0,
    ) -> dict[str, Any]:
        return await service(context).search_documents(
            query,
            file_types=file_types,
            created_after=created_after,
            modified_after=modified_after,
            limit=limit,
            offset=offset,
        )

    @mcp.tool(
        title="Get SharePoint document metadata",
        description=(
            "Get current metadata for a SharePoint search result. The document_ref "
            "must come from search_documents."
        ),
        annotations=READ_ONLY_TOOL,
        structured_output=True,
    )
    async def get_document_metadata(
        document_ref: str,
        context: Context[Any, ApplicationContext, Any],
    ) -> dict[str, Any]:
        return await service(context).get_document_metadata(document_ref)

    @mcp.tool(
        title="Read SharePoint document text",
        description=(
            "Download a SharePoint document and extract text from PDF, DOCX, PPTX, "
            "XLSX, or a plain-text format. Results are sectioned and bounded. Use "
            "next_section as start_section to continue. A continuation downloads the "
            "file again because the server is stateless."
        ),
        annotations=READ_ONLY_TOOL,
        structured_output=True,
    )
    async def read_document(
        document_ref: str,
        context: Context[Any, ApplicationContext, Any],
        start_section: int = 0,
        section_count: int = 10,
        max_characters: int | None = None,
    ) -> dict[str, Any]:
        return await service(context).read_document(
            document_ref,
            start_section=start_section,
            section_count=section_count,
            max_characters=max_characters,
        )

    @mcp.tool(
        title="List configured SharePoint sources",
        description=(
            "List the configured SharePoint sites and the document libraries "
            "discovered in each site."
        ),
        annotations=READ_ONLY_TOOL,
        structured_output=True,
    )
    async def list_sources(
        context: Context[Any, ApplicationContext, Any],
    ) -> dict[str, Any]:
        return service(context).list_sources()

    return mcp


def _arguments(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="mcp-sharepoint",
        description="Run the stateless SharePoint Microsoft Graph MCP server.",
    )
    parser.add_argument(
        "--transport",
        choices=("stdio", "streamable-http"),
        help="Override MCP_TRANSPORT for this process.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = _arguments(argv)
    try:
        settings = Settings()  # type: ignore[call-arg]
        if args.transport:
            settings = settings.model_copy(update={"mcp_transport": args.transport})
    except ValidationError as exc:
        print(f"Configuration error: {exc}", file=sys.stderr)
        raise SystemExit(2) from exc

    logging.basicConfig(
        level=getattr(logging, settings.log_level),
        stream=sys.stderr,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    LOGGER.info(
        "Starting SharePoint MCP server with transport %s",
        settings.mcp_transport,
    )
    server = create_server(settings)
    server.run(transport=cast(Transport, settings.mcp_transport))


if __name__ == "__main__":
    main()
