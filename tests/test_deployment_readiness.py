from __future__ import annotations

from io import StringIO
from pathlib import Path

import pandas as pd
import pytest

from app import (
    APP_TITLE,
    _authentication_required,
    _email_is_allowed,
    _oidc_is_configured,
    _workspace_id_for_identity,
    prepare_uploaded_data,
    status_for_value,
)
from storage import get_engine, load_workspace, save_workspace


def _valid_upload() -> pd.DataFrame:
    return pd.read_csv(
        StringIO(
            "Week Start,Revenue,Active Users,Signups,Conversions,Support Tickets Open,CSAT,Headcount\n"
            "2026-09-28,1200,3000,100,7,20,4.5,12\n"
        )
    )


@pytest.fixture
def isolated_database(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    get_engine.cache_clear()
    database_path = tmp_path / "scorecard.sqlite"
    monkeypatch.setenv("DATABASE_URL", f"sqlite:///{database_path.as_posix()}")
    yield
    get_engine().dispose()
    get_engine.cache_clear()


def test_upload_derives_week_fields_and_conversion_rate() -> None:
    data = prepare_uploaded_data(_valid_upload())

    assert data.loc[0, "Week #"] == 40
    assert data.loc[0, "Week End"] == pd.Timestamp("2026-10-04")
    assert data.loc[0, "Conversion Rate"] == pytest.approx(0.07)
    assert data.loc[0, "Notes"] == ""


def test_upload_rejects_invalid_week_start() -> None:
    data = _valid_upload()
    data.loc[0, "Week Start"] = "not-a-date"

    with pytest.raises(ValueError, match="Week Start must contain a valid date"):
        prepare_uploaded_data(data)


def test_kpi_status_guardrails() -> None:
    assert [status_for_value(value, 100, True) for value in (100, 90, 89)] == [
        "On Track",
        "At Risk",
        "Off Track",
    ]
    assert [status_for_value(value, 100, False) for value in (100, 115, 116)] == [
        "On Track",
        "At Risk",
        "Off Track",
    ]


def test_product_name_is_kpiling() -> None:
    assert APP_TITLE == "KPIling | KPI Scorecard"


def test_oidc_configuration_and_email_allowlist() -> None:
    settings = {
        "redirect_uri": "https://scorecard.example.com/oauth2callback",
        "cookie_secret": "test-only",
        "client_id": "test-client",
        "client_secret": "test-secret",
        "server_metadata_url": "https://issuer.example.com/.well-known/openid-configuration",
    }

    assert _oidc_is_configured(settings)
    assert not _oidc_is_configured({"client_id": "test-client"})
    assert _email_is_allowed(" Analyst@Example.com ", {"analyst@example.com"})
    assert not _email_is_allowed("other@example.com", {"analyst@example.com"})
    assert _email_is_allowed("anyone@example.com", set())


def test_workspace_id_is_stable_and_issuer_scoped() -> None:
    first = _workspace_id_for_identity("https://issuer-a.example", "subject-1")

    assert first == _workspace_id_for_identity("https://issuer-a.example", "subject-1")
    assert first != _workspace_id_for_identity("https://issuer-b.example", "subject-1")
    assert len(first) == 64


def test_authentication_policy_fails_closed_for_production_inputs() -> None:
    assert _authentication_required(True, "sqlite:///demo.db", set())
    assert _authentication_required(False, "postgresql+psycopg://host/db", set())
    assert _authentication_required(False, "sqlite:///demo.db", {"analyst@example.com"})
    assert not _authentication_required(False, "sqlite:///demo.db", set())


def test_database_persists_data_and_isolates_users(isolated_database: None) -> None:
    raw_data = prepare_uploaded_data(_valid_upload())
    targets = pd.DataFrame({"Metric": ["Revenue"], "Weekly Target": [5000.0]})

    save_workspace("user-a", raw_data, targets)
    saved = load_workspace("user-a")

    assert saved is not None
    assert saved["raw_data"].loc[0, "Revenue"] == 1200
    assert saved["targets"].loc[0, "Weekly Target"] == 5000
    assert load_workspace("user-b") is None
