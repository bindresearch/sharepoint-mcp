from __future__ import annotations

from pathlib import Path

import pytest
from pydantic import ValidationError

from mcp_sharepoint.config import Settings


BASE_ENV = {
    "AZURE_TENANT_ID": "tenant",
    "AZURE_CLIENT_ID": "client",
    "AZURE_CLIENT_SECRET": "secret",
    "GRAPH_SEARCH_REGION": "eur",
    "SHAREPOINT_SITE_URLS": (
        "https://example.sharepoint.com/sites/Engineering/;"
        "https://example.sharepoint.com/sites/Research"
    ),
}


@pytest.fixture(autouse=True)
def isolate_settings_sources(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.chdir(tmp_path)
    for name in BASE_ENV:
        monkeypatch.delenv(name, raising=False)


def test_settings_load_and_normalize_values() -> None:
    settings = Settings.model_validate(BASE_ENV)

    assert settings.search_region == "EUR"
    assert tuple(str(url).rstrip("/") for url in settings.site_urls) == (
        "https://example.sharepoint.com/sites/Engineering",
        "https://example.sharepoint.com/sites/Research",
    )
    assert settings.mcp_transport == "stdio"
    assert "secret" not in repr(settings)


def test_settings_load_from_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    for name, value in BASE_ENV.items():
        monkeypatch.setenv(name, value)

    settings = Settings()  # type: ignore[call-arg]

    assert settings.tenant_id == "tenant"
    assert len(settings.site_urls) == 2


def test_settings_load_dotenv_file(tmp_path: Path) -> None:
    (tmp_path / ".env").write_text(
        "\n".join(f"{name}={value}" for name, value in BASE_ENV.items()),
        encoding="utf-8",
    )

    settings = Settings()  # type: ignore[call-arg]

    assert settings.tenant_id == "tenant"
    assert settings.client_secret.get_secret_value() == "secret"
    assert len(settings.site_urls) == 2


def test_settings_reject_non_https_site_url() -> None:
    values = {**BASE_ENV, "SHAREPOINT_SITE_URLS": "http://example.test/site"}

    with pytest.raises(ValidationError, match="https"):
        Settings.model_validate(values)


def test_settings_require_graph_credentials() -> None:
    values = {
        key: value for key, value in BASE_ENV.items() if key != "AZURE_CLIENT_SECRET"
    }

    with pytest.raises(ValidationError, match="AZURE_CLIENT_SECRET"):
        Settings.model_validate(values)


def test_section_size_must_fit_in_response() -> None:
    values = {
        **BASE_ENV,
        "MAX_SECTION_CHARACTERS": 2_000,
        "MAX_RESPONSE_CHARACTERS": 1_000,
    }

    with pytest.raises(ValidationError, match="must not exceed"):
        Settings.model_validate(values)
