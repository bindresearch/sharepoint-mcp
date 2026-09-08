> [!WARNING]
> This is entirely vibe-coded (but human-reviewed)

# SharePoint MCP server

A read-only, stateless MCP server for Microsoft SharePoint Online. It uses Microsoft Graph to search configured SharePoint sites, read document metadata, and return extracted document text.

The server does not use a local search index or database. Microsoft Search supplies keyword results. A selected document is downloaded and parsed for each `read_document` call.

## Features

- Full-text keyword search through Microsoft Graph Search
- Search across every document library in configured SharePoint sites
- PDF, DOCX, PPTX, XLSX, TXT, CSV, Markdown, and log-file text extraction
- Creator and modification metadata
- STDIO and stateless Streamable HTTP transports
- Application authentication with an Entra client secret
- No Graph write permissions or MCP write tools

## Requirements

- Python 3.13 or later
- An on-premises host with outbound HTTPS access to Microsoft identity, Graph, and SharePoint download endpoints
- An Entra ID application with Microsoft Graph application permission `Sites.Read.All`
- Tenant administrator consent for that permission

`Sites.Read.All` lets the application credential read all SharePoint sites in the tenant. `SHAREPOINT_SITE_URLS` limits what this server searches and returns, but it is not an OAuth permission boundary. Protect the client secret accordingly.

## Entra application setup

1. In the Azure portal, open **Microsoft Entra ID > App registrations**.
2. Create an application registration.
3. Add the Microsoft Graph **Application** permission `Sites.Read.All`.
4. Grant tenant administrator consent.
5. Create a client secret.
6. Record the tenant ID, application/client ID, and secret value.

Do not add a Graph write permission.

## Installation

Using [uv](https://docs.astral.sh/uv/):

```bash
uv sync
```

Or install the project into a virtual environment:

```bash
python -m pip install .
```

## Configuration

Copy `.env.example` to `.env` and edit it for local use. The server loads `.env` from its current working directory; process environment variables take precedence over values in the file.

Required variables:

```bash
export AZURE_TENANT_ID="00000000-0000-0000-0000-000000000000"
export AZURE_CLIENT_ID="00000000-0000-0000-0000-000000000000"
export AZURE_CLIENT_SECRET="replace-me"
export GRAPH_SEARCH_REGION="EUR"
export SHAREPOINT_SITE_URLS="https://example.sharepoint.com/sites/Engineering;https://example.sharepoint.com/sites/Research"
```

`SHAREPOINT_SITE_URLS` is a semicolon-separated list. Configure SharePoint **site URLs**, not document-library URLs. At startup, the server resolves each site and discovers all drives (document libraries) in that site.

Microsoft Search requires `GRAPH_SEARCH_REGION` for application-permission requests. Use the Microsoft 365 data location for the tenant. See [Get the region value](https://learn.microsoft.com/graph/search-concept-searchall).

Optional variables:

| Variable | Default | Purpose |
| --- | ---: | --- |
| `MCP_TRANSPORT` | `stdio` | `stdio` or `streamable-http` |
| `MCP_HOST` | `127.0.0.1` | HTTP bind address |
| `MCP_PORT` | `8000` | HTTP port |
| `MAX_SEARCH_RESULTS` | `50` | Largest accepted search page |
| `MAX_DOWNLOAD_MB` | `100` | Largest document that can be read |
| `MAX_RESPONSE_CHARACTERS` | `50000` | Largest extracted-text response |
| `MAX_SECTION_CHARACTERS` | `8000` | Maximum extracted section size |
| `MAX_EXTRACTED_CHARACTERS` | `20000000` | Extraction safety limit per document |
| `EXTRACTION_TIMEOUT_SECONDS` | `120` | Text-extraction timeout |
| `GRAPH_REQUEST_TIMEOUT_SECONDS` | `60` | Graph HTTP timeout |
| `GRAPH_MAX_RETRIES` | `4` | Retries for throttling and transient failures |
| `SPOOL_THRESHOLD_MB` | `8` | Memory threshold before a download uses a temporary file |
| `LOG_LEVEL` | `INFO` | Server log level |

## Run with STDIO

```bash
uv run mcp-sharepoint --transport stdio
```

Example MCP client configuration:

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "uv",
      "args": [
        "--directory",
        "/absolute/path/to/mcp-sharepoint",
        "run",
        "mcp-sharepoint",
        "--transport",
        "stdio"
      ],
      "env": {
        "AZURE_TENANT_ID": "...",
        "AZURE_CLIENT_ID": "...",
        "AZURE_CLIENT_SECRET": "...",
        "GRAPH_SEARCH_REGION": "EUR",
        "SHAREPOINT_SITE_URLS": "https://example.sharepoint.com/sites/Engineering"
      }
    }
  }
}
```

STDIO logs are written to stderr. Stdout is reserved for MCP messages.

Do not put the client secret in MCP configurations on employee workstations. STDIO mode is intended for a client that runs on the secured server or for development.

## Run with Streamable HTTP

```bash
export MCP_TRANSPORT="streamable-http"
export MCP_HOST="127.0.0.1"
export MCP_PORT="8000"
uv run mcp-sharepoint
```

The MCP endpoint is `http://127.0.0.1:8000/mcp` by default. This server does not authenticate MCP clients. Use a private interface, firewall rules, and a TLS reverse proxy when exposing it to the internal network.

## MCP tools

### `search_documents`

Searches file content and metadata in configured sites.

Arguments:

- `query`: required keyword text
- `file_types`: optional list such as `["pdf", "docx"]`
- `created_after`: optional `YYYY-MM-DD` date
- `modified_after`: optional `YYYY-MM-DD` date
- `limit`: 1 to `MAX_SEARCH_RESULTS`
- `offset`: zero-based result offset

The tool builds KQL itself. It does not accept raw KQL. Results include a `document_ref` for the other tools.

### `get_document_metadata`

Returns current Graph metadata for a `document_ref`, including creator, creation time, modifier, modification time, size, MIME type, and SharePoint URL.

### `read_document`

Downloads a document and returns extracted sections. Use `next_section` as `start_section` to continue. Because the server is stateless, each continuation downloads and parses the document again.

### `list_sources`

Lists configured SharePoint sites and the document libraries discovered at startup.

## Supported file types

- PDF files with embedded text; OCR is not performed
- DOCX
- PPTX
- XLSX
- TXT, TEXT, CSV, Markdown, and LOG

Legacy DOC, PPT, and XLS formats are not supported.

## Operational notes

- Microsoft Search is eventually consistent. A new or changed file might not appear immediately.
- Search snippets and metadata come from the Microsoft Search index. `get_document_metadata` gets current metadata directly from Graph.
- Large files and repeated reads can be slow because there is no persistent cache.
- The server retries Graph throttling responses and honors `Retry-After`.
- Downloads use temporary spooled files and are removed after each request.

## Development

```bash
uv sync --all-groups
uv run pytest
uvx ruff check .
```
