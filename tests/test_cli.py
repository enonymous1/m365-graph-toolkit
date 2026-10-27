"""CLI tests: demo mode offline, live commands against fake clients."""

from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

from m365_graph_toolkit import cli
from m365_graph_toolkit.demo_data import demo_users


def test_demo_writes_fictional_exports(tmp_path: Path) -> None:
    assert cli.main(["demo", "--out", str(tmp_path)]) == 0
    names = {path.name for path in tmp_path.iterdir()}
    assert "DEMO_README.txt" in names
    assert any(name.startswith("users_demo_") and name.endswith(".csv") for name in names)
    assert any(name.startswith("sign_in_analysis_demo_") for name in names)
    readme = (tmp_path / "DEMO_README.txt").read_text(encoding="utf-8")
    assert "fictional" in readme
    analysis_files = list(tmp_path.glob("sign_in_analysis_demo_*.json"))
    analysis = json.loads(analysis_files[0].read_text(encoding="utf-8"))
    assert analysis["total_sign_ins"] == 4
    assert analysis["failed_sign_ins"] == 1


def test_export_users_with_fake_client(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    fake_client = MagicMock()
    fake_client.get_users = AsyncMock(return_value=demo_users())
    monkeypatch.setattr(cli, "_make_users_client", lambda: fake_client)
    assert cli.main(["export", "users", "--out", str(tmp_path)]) == 0
    csv_files = list(tmp_path.glob("users_*.csv"))
    assert csv_files, "expected a users CSV export"
    assert "Ada Example" in csv_files[0].read_text(encoding="utf-8")


def test_logs_sign_ins_with_fake_client(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    from m365_graph_toolkit.demo_data import demo_sign_in_logs

    fake_client = MagicMock()
    fake_client.get_audit_logs = AsyncMock(return_value=demo_sign_in_logs())
    monkeypatch.setattr(cli, "_make_audit_client", lambda: fake_client)
    assert cli.main(["logs", "sign-ins", "--out", str(tmp_path), "--days", "7"]) == 0
    assert list(tmp_path.glob("sign_in_logs_*.csv"))


def test_check_auth_prints_expiry_only(monkeypatch: pytest.MonkeyPatch, capsys) -> None:
    fake_auth = MagicMock()
    fake_auth.get_access_token.return_value = SimpleNamespace(expires_on=1234567890)
    monkeypatch.setattr(cli, "AuthManager", lambda: fake_auth)
    assert cli.main(["check-auth"]) == 0
    captured = capsys.readouterr()
    assert "1234567890" in captured.out
    assert "token value" in captured.out.lower()


def test_missing_credentials_returns_1(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    # Live path with no env vars: AuthManager raises AuthError, CLI
    # returns 1 instead of a traceback.
    monkeypatch.delenv("MS_TENANT_ID", raising=False)
    monkeypatch.delenv("MS_CLIENT_ID", raising=False)
    monkeypatch.delenv("MS_CLIENT_SECRET", raising=False)
    assert cli.main(["export", "users", "--out", str(tmp_path)]) == 1
