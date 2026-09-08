"""Environment-based configuration for the SharePoint MCP server."""

from __future__ import annotations

from typing import Annotated, Literal, Self

from pydantic import (
    AnyUrl,
    Field,
    SecretStr,
    StringConstraints,
    UrlConstraints,
    field_validator,
    model_validator,
)
from pydantic_settings import BaseSettings, NoDecode, SettingsConfigDict

NonEmptyString = Annotated[str, StringConstraints(strip_whitespace=True, min_length=1)]
SearchRegion = Annotated[
    str,
    StringConstraints(strip_whitespace=True, to_upper=True, min_length=2, max_length=3),
]
SharePointSiteUrl = Annotated[
    AnyUrl,
    UrlConstraints(allowed_schemes=["https"], host_required=True),
]
Transport = Literal["stdio", "streamable-http"]
LogLevel = Literal["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]


class Settings(BaseSettings):
    """Validated settings loaded directly from environment variables."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        env_ignore_empty=True,
        extra="ignore",
        frozen=True,
        populate_by_name=True,
    )

    tenant_id: NonEmptyString = Field(validation_alias="AZURE_TENANT_ID")
    client_id: NonEmptyString = Field(validation_alias="AZURE_CLIENT_ID")
    client_secret: SecretStr = Field(
        validation_alias="AZURE_CLIENT_SECRET", min_length=1, repr=False
    )
    site_urls: Annotated[tuple[SharePointSiteUrl, ...], NoDecode] = Field(
        validation_alias="SHAREPOINT_SITE_URLS", min_length=1
    )
    search_region: SearchRegion = Field(validation_alias="GRAPH_SEARCH_REGION")

    graph_request_timeout_seconds: int = Field(default=60, ge=1, le=600)
    graph_max_retries: int = Field(default=4, ge=0, le=10)
    max_download_mb: int = Field(default=100, ge=1, le=10_000)
    spool_threshold_mb: int = Field(default=8, ge=1, le=1_000)
    extraction_timeout_seconds: int = Field(default=120, ge=1, le=3_600)
    max_extracted_characters: int = Field(default=20_000_000, ge=10_000, le=200_000_000)
    max_section_characters: int = Field(default=8_000, ge=500, le=1_000_000)
    max_response_characters: int = Field(default=50_000, ge=1_000, le=1_000_000)
    max_search_results: int = Field(default=50, ge=1, le=100)
    mcp_host: NonEmptyString = "127.0.0.1"
    mcp_port: int = Field(default=8_000, ge=1, le=65_535)
    mcp_transport: Transport = "stdio"
    log_level: LogLevel = "INFO"

    @field_validator("site_urls", mode="before")
    @classmethod
    def split_site_urls(cls, value: object) -> object:
        if isinstance(value, str):
            return tuple(part.strip() for part in value.split(";") if part.strip())
        return value

    @model_validator(mode="after")
    def validate_section_size(self) -> Self:
        if self.max_section_characters > self.max_response_characters:
            raise ValueError(
                "MAX_SECTION_CHARACTERS must not exceed MAX_RESPONSE_CHARACTERS"
            )
        return self
