"""AuditLogsClient tests (fake Graph clients, captured filter strings)."""

from __future__ import annotations

from datetime import datetime, timezone
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

from m365_graph_toolkit.clients.audit_logs import AuditLogsClient
from m365_graph_toolkit.errors import ValidationError
from tests.fakes import page


def _sign_in_obj(**overrides):
    status = SimpleNamespace(error_code=0, failure_reason=None)
    data = {
        "id": "log-1",
        "created_date_time": "2026-10-20T09:00:00+00:00",
        "user_display_name": "Ada Example",
        "user_principal_name": "ada@example.com",
        "user_id": "u1",
        "app_display_name": "App",
        "app_id": "a1",
        "ip_address": "203.0.113.10",
        "client_app_used": "Browser",
        "is_interactive": True,
        "risk_state": "none",
        "risk_level_aggregated": "low",
        "risk_level_during_sign_in": "low",
        "status": status,
    }
    data.update(overrides)
    return SimpleNamespace(**data)


async def test_get_audit_logs_applies_time_filter_to_sign_ins(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict = {}

    def fake_config(self, **kwargs):  # noqa: ANN001, ANN202
        captured.update(kwargs)
        return "CONFIG"

    monkeypatch.setattr(AuditLogsClient, "_logs_configuration", fake_config)
    client = MagicMock()
    client.audit_logs.sign_ins.get = AsyncMock(return_value=page([_sign_in_obj()]))
    start = datetime(2026, 10, 13, tzinfo=timezone.utc)
    end = datetime(2026, 10, 20, tzinfo=timezone.utc)
    logs = await AuditLogsClient(client).get_audit_logs(
        log_type="signIn", start_time=start, end_time=end
    )
    assert captured["filter_value"] == (
        "createdDateTime ge 2026-10-13T00:00:00Z and createdDateTime le 2026-10-20T00:00:00Z"
    )
    assert captured["top"] == 1000
    assert logs[0]["status"] == {"error_code": 0, "failure_reason": None}


async def test_get_directory_audit_logs_honours_days(monkeypatch: pytest.MonkeyPatch) -> None:
    captured: dict = {}

    def fake_config(self, **kwargs):  # noqa: ANN001, ANN202
        captured.update(kwargs)
        return "CONFIG"

    monkeypatch.setattr(AuditLogsClient, "_logs_configuration", fake_config)
    client = MagicMock()
    audit = SimpleNamespace(
        id="a1",
        activity_date_time="2026-10-19T09:00:00+00:00",
        activity_display_name="Update user",
        category="UserManagement",
        result="success",
        logged_by_service="Core Directory",
        initiated_by=None,
        target_resources=None,
    )
    client.audit_logs.directory_audits.get = AsyncMock(return_value=page([audit]))
    logs = await AuditLogsClient(client).get_directory_audit_logs(days=3)
    # v1 hardcoded 30 days here regardless of the argument. The filter
    # must use activityDateTime (not createdDateTime) for directory audits.
    assert "activityDateTime ge" in captured["filter_value"]
    assert "createdDateTime" not in captured["filter_value"]
    assert len(logs) == 1


async def test_get_audit_logs_rejects_naive_datetimes() -> None:
    client = MagicMock()
    with pytest.raises(ValidationError):
        await AuditLogsClient(client).get_audit_logs(
            start_time=datetime(2026, 10, 13), end_time=datetime(2026, 10, 20)
        )


async def test_get_audit_logs_rejects_bad_log_type() -> None:
    with pytest.raises(ValidationError):
        await AuditLogsClient(MagicMock()).get_audit_logs(log_type="everything")


async def test_get_all_sign_in_logs_groups_by_user(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(AuditLogsClient, "_logs_configuration", lambda self, **kwargs: "CONFIG")
    client = MagicMock()
    client.audit_logs.sign_ins.get = AsyncMock(
        return_value=page([_sign_in_obj(user_id="u1"), _sign_in_obj(id="log-2", user_id="u2")])
    )
    grouped = await AuditLogsClient(client).get_all_sign_in_logs(days=7)
    assert set(grouped) == {"u1", "u2"}
    assert len(grouped["u1"]) == 1


async def test_get_user_sign_in_logs_success_only(monkeypatch: pytest.MonkeyPatch) -> None:
    captured: dict = {}

    def fake_config(self, **kwargs):  # noqa: ANN001, ANN202
        captured.update(kwargs)
        return "CONFIG"

    monkeypatch.setattr(AuditLogsClient, "_logs_configuration", fake_config)
    client = MagicMock()
    client.audit_logs.sign_ins.get = AsyncMock(return_value=page([]))
    await AuditLogsClient(client).get_user_sign_in_logs("u1", days=30, success_only=True)
    assert "status/errorCode eq 0" in captured["filter_value"]
    assert "userId eq 'u1'" in captured["filter_value"]
