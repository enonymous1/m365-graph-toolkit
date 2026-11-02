"""Google Sheets exporter + ``sync`` CLI tests, all offline with fakes.

The fakes mirror the real ``gspread`` 6.2 surface verified against the
installed package: ``Spreadsheet.worksheet(title)`` raises
``WorksheetNotFound``, ``Spreadsheet.add_worksheet(title, rows, cols)``
creates one, and ``Worksheet.update(values, range_name=...)`` writes
values starting at that range. No test here talks to Google.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from unittest.mock import AsyncMock, MagicMock

import pytest

from m365_graph_toolkit import cli
from m365_graph_toolkit.demo_data import demo_mail_messages, demo_users
from m365_graph_toolkit.errors import ExportError, ValidationError
from m365_graph_toolkit.sheets import (
    SheetsExporter,
    _load_service_account_info,
    sanitize_worksheet_title,
)


class WorksheetNotFound(Exception):
    """Fake matching ``gspread.exceptions.WorksheetNotFound`` by name."""


class FakeAPIError(Exception):
    """Fake Sheets API error carrying an HTTP status, like ``APIError``."""

    def __init__(self, status: int) -> None:
        super().__init__(f"API error {status}")
        self.response_status_code = status


class FakeWorksheet:
    def __init__(self, title: str, rows: int, cols: int) -> None:
        self.title = title
        self.rows = rows
        self.cols = cols
        self.data: list[list[Any]] = []
        self.clears = 0
        self.updates: list[tuple[list[list[Any]], str | None]] = []
        self.failures: list[BaseException] = []
        self.resize_calls: list[tuple[int, int]] = []

    @property
    def row_count(self) -> int:
        return self.rows

    @property
    def col_count(self) -> int:
        return self.cols

    def resize(self, rows: int | None = None, cols: int | None = None) -> dict[str, Any]:
        if rows is not None:
            self.rows = rows
        if cols is not None:
            self.cols = cols
        self.resize_calls.append((self.rows, self.cols))
        return {}

    def clear(self) -> None:
        self.data = []
        self.clears += 1

    def update(self, values: list[list[Any]], range_name: str | None = None) -> dict[str, Any]:
        if self.failures:
            raise self.failures.pop(0)
        self.updates.append((values, range_name))
        self.data = list(values)
        return {}


class FakeSpreadsheet:
    def __init__(self) -> None:
        self.worksheets: dict[str, FakeWorksheet] = {}

    def worksheet(self, title: str) -> FakeWorksheet:
        try:
            return self.worksheets[title]
        except KeyError as exc:
            raise WorksheetNotFound(title) from exc

    def add_worksheet(self, title: str, rows: int, cols: int) -> FakeWorksheet:
        worksheet = FakeWorksheet(title, rows, cols)
        self.worksheets[title] = worksheet
        return worksheet


class FakeClient:
    def __init__(self) -> None:
        self.spreadsheet = FakeSpreadsheet()
        self.opened_keys: list[str] = []
        self.open_failures: list[BaseException] = []

    def open_by_key(self, key: str) -> FakeSpreadsheet:
        if self.open_failures:
            raise self.open_failures.pop(0)
        self.opened_keys.append(key)
        return self.spreadsheet


def make_exporter(client: FakeClient, **kwargs: Any) -> SheetsExporter:
    return SheetsExporter(client=client, spreadsheet_id="sheet-1", base_delay=0, **kwargs)


def test_sanitize_worksheet_title() -> None:
    assert sanitize_worksheet_title("users") == "users"
    assert sanitize_worksheet_title("mail-a/b[c]") == "mail-a_b_c_"
    assert sanitize_worksheet_title("  spaced title  ") == "spaced title"
    assert len(sanitize_worksheet_title("x" * 200)) == 100
    with pytest.raises(ValidationError):
        sanitize_worksheet_title("   ")
    # A title of only forbidden characters becomes underscores, not empty.
    assert sanitize_worksheet_title("[]*?/\\") == "______"


def test_write_rows_creates_worksheet_and_writes_values() -> None:
    client = FakeClient()
    exporter = make_exporter(client)
    stamp = datetime(2026, 10, 20, 9, 0, tzinfo=timezone.utc)
    rows = [
        {"id": "1", "name": "Ada", "when": stamp, "note": None, "tags": ["a", "b"]},
        {"id": "2", "name": "Ben", "when": None, "note": "hi", "tags": []},
    ]
    assert exporter.write_rows(rows, "users") == 2
    # Construction opened nothing; the write did, exactly once.
    assert client.opened_keys == ["sheet-1"]
    worksheet = client.spreadsheet.worksheets["users"]
    assert worksheet.rows >= 100
    assert worksheet.cols >= 5
    assert worksheet.clears == 1
    assert len(worksheet.updates) == 1
    values, range_name = worksheet.updates[0]
    assert range_name == "A1"
    assert values[0] == ["id", "name", "when", "note", "tags"]
    assert values[1] == ["1", "Ada", "2026-10-20T09:00:00+00:00", "", '["a", "b"]']
    assert values[2][3] == "hi"
    # Input rows are not mutated.
    assert rows[0]["note"] is None


def test_write_rows_reuses_existing_worksheet() -> None:
    client = FakeClient()
    client.spreadsheet.add_worksheet("users", rows=100, cols=10)
    exporter = make_exporter(client)
    assert exporter.write_rows([{"id": "1"}], "users") == 1
    assert len(client.spreadsheet.worksheets) == 1


def test_write_rows_resizes_too_small_existing_worksheet() -> None:
    # A sync into a worksheet created for fewer rows/columns used to
    # fail on the Sheets update; the exporter must grow the grid.
    client = FakeClient()
    worksheet = client.spreadsheet.add_worksheet("users", rows=2, cols=1)
    exporter = make_exporter(client)
    rows = [{"id": str(i), "name": f"User {i}"} for i in range(3)]
    assert exporter.write_rows(rows, "users") == 3
    assert worksheet.resize_calls == [(4, 2)]
    assert worksheet.rows == 4
    assert worksheet.cols == 2


def test_write_rows_empty_without_fieldnames_raises() -> None:
    with pytest.raises(ValidationError):
        make_exporter(FakeClient()).write_rows([], "users")


def test_write_rows_empty_with_fieldnames_writes_header_only() -> None:
    client = FakeClient()
    exporter = make_exporter(client)
    assert exporter.write_rows([], "users", fieldnames=["id", "name"]) == 0
    worksheet = client.spreadsheet.worksheets["users"]
    assert worksheet.updates[0][0] == [["id", "name"]]


def test_write_rows_rejects_unknown_fields() -> None:
    with pytest.raises(ValidationError):
        make_exporter(FakeClient()).write_rows(
            [{"id": "1", "surprise": 2}], "users", fieldnames=["id"]
        )


def test_write_rows_chunks_and_ranges_do_not_overlap() -> None:
    client = FakeClient()
    exporter = make_exporter(client)
    rows = [{"id": str(i)} for i in range(5)]
    assert exporter.write_rows(rows, "users", chunk_size=2) == 5
    worksheet = client.spreadsheet.worksheets["users"]
    ranges = [range_name for _, range_name in worksheet.updates]
    assert ranges == ["A1", "A4", "A6"]
    chunks = [values for values, _ in worksheet.updates]
    assert chunks[0][0] == ["id"]
    assert len(chunks[0]) == 3  # header + 2 rows
    assert len(chunks[1]) == 2
    assert len(chunks[2]) == 1
    # Reassembled, the chunks are the full sheet with no row lost or doubled.
    reassembled = [chunks[0][0], *chunks[0][1:], *chunks[1], *chunks[2]]
    assert [row[0] for row in reassembled] == ["id", "0", "1", "2", "3", "4"]


def test_load_service_account_info_forms() -> None:
    assert _load_service_account_info({"type": "service_account"}) == {"type": "service_account"}
    info = {"type": "service_account", "project_id": "p"}
    assert _load_service_account_info(json.dumps(info)) == info
    with pytest.raises(ValidationError):
        _load_service_account_info("")
    with pytest.raises(ValidationError):
        # A mistyped path is not valid JSON either: say so plainly.
        _load_service_account_info("/definitely/not/a/key.json")
    with pytest.raises(ValidationError):
        _load_service_account_info("[1, 2]")


def test_load_service_account_info_from_path(tmp_path: Path) -> None:
    key_file = tmp_path / "key.json"
    key_file.write_text(json.dumps({"type": "service_account"}), encoding="utf-8")
    assert _load_service_account_info(str(key_file)) == {"type": "service_account"}
    bad_file = tmp_path / "bad.json"
    bad_file.write_text("{not json", encoding="utf-8")
    with pytest.raises(ExportError):
        _load_service_account_info(str(bad_file))


def test_write_rows_retries_retryable_status_only() -> None:
    client = FakeClient()
    worksheet = client.spreadsheet.add_worksheet("users", rows=100, cols=10)
    worksheet.failures = [FakeAPIError(429), FakeAPIError(429)]
    sleeps: list[float] = []
    exporter = SheetsExporter(
        client=client, spreadsheet_id="sheet-1", base_delay=0, sleep=sleeps.append
    )
    assert exporter.write_rows([{"id": "1"}], "users") == 1
    assert len(worksheet.updates) == 1
    assert len(sleeps) == 2

    client2 = FakeClient()
    worksheet2 = client2.spreadsheet.add_worksheet("users", rows=100, cols=10)
    worksheet2.failures = [FakeAPIError(403)]
    with pytest.raises(ExportError):
        make_exporter(client2).write_rows([{"id": "1"}], "users")
    assert len(worksheet2.updates) == 0


def test_open_failure_is_wrapped_as_export_error() -> None:
    client = FakeClient()
    client.open_failures = [RuntimeError("boom")]
    with pytest.raises(ExportError, match="boom"):
        make_exporter(client).write_rows([{"id": "1"}], "users")


def test_real_gspread_shape() -> None:
    """Fail loudly if the gspread API this exporter targets drifts."""
    gspread = pytest.importorskip("gspread")
    import inspect

    params = list(inspect.signature(gspread.Worksheet.update).parameters)
    assert params[1] == "values"  # params[0] is self
    assert "range_name" in params
    assert "title" in inspect.signature(gspread.Spreadsheet.add_worksheet).parameters
    assert gspread.WorksheetNotFound is not None


def test_sync_users_dry_run_needs_no_sheets_config(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    fake_client = MagicMock()
    fake_client.get_users = AsyncMock(return_value=demo_users())
    monkeypatch.setattr(cli, "_make_users_client", lambda: fake_client)
    monkeypatch.setattr(
        cli,
        "_make_sheets_exporter",
        lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("must not build")),
    )
    monkeypatch.delenv("GCS_SERVICE_ACCOUNT_KEY_JSON", raising=False)
    monkeypatch.delenv("TARGET_GOOGLE_SHEET_ID", raising=False)
    assert cli.main(["sync", "users", "--dry-run"]) == 0
    assert "Dry-run" in capsys.readouterr().out


def test_sync_users_writes_via_exporter(monkeypatch: pytest.MonkeyPatch) -> None:
    fake_client = MagicMock()
    fake_client.get_users = AsyncMock(return_value=demo_users())
    monkeypatch.setattr(cli, "_make_users_client", lambda: fake_client)
    recorded: dict[str, Any] = {}

    class FakeExporter:
        def write_rows(self, rows: list[dict[str, Any]], title: str, **kwargs: Any) -> int:
            recorded["rows"] = rows
            recorded["title"] = title
            recorded["fieldnames"] = kwargs.get("fieldnames")
            return len(rows)

    monkeypatch.setattr(cli, "_make_sheets_exporter", lambda *args: FakeExporter())
    monkeypatch.setenv("GCS_SERVICE_ACCOUNT_KEY_JSON", "{}")
    monkeypatch.setenv("TARGET_GOOGLE_SHEET_ID", "sheet-1")
    assert cli.main(["sync", "users"]) == 0
    assert recorded["title"] == "users"
    assert recorded["rows"] == demo_users()
    from m365_graph_toolkit.clients.users import USER_FIELDS

    assert recorded["fieldnames"] == list(USER_FIELDS)


def test_sync_missing_sheets_config_returns_1(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    fake_client = MagicMock()
    fake_client.get_users = AsyncMock(return_value=demo_users())
    monkeypatch.setattr(cli, "_make_users_client", lambda: fake_client)
    monkeypatch.delenv("GCS_SERVICE_ACCOUNT_KEY_JSON", raising=False)
    monkeypatch.delenv("TARGET_GOOGLE_SHEET_ID", raising=False)
    assert cli.main(["sync", "users"]) == 1


def test_sync_users_without_config_raises_instead_of_asserting(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    # The old code used ``assert`` for this guard, which vanishes
    # under ``python -O``. A direct call with no config must raise a
    # validation error, not crash later inside the exporter.
    import asyncio

    fake_client = MagicMock()
    fake_client.get_users = AsyncMock(return_value=demo_users())
    monkeypatch.setattr(cli, "_make_users_client", lambda: fake_client)
    with pytest.raises(ValidationError):
        asyncio.run(
            cli._sync_users(
                service_account=None,
                spreadsheet_id=None,
                worksheet="users",
                dry_run=False,
                top=None,
            )
        )


def test_main_returns_1_on_output_error(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setattr(
        cli,
        "_run_demo",
        lambda out_dir: (_ for _ in ()).throw(OSError("disk full")),
    )
    assert cli.main(["demo", "--out", str(tmp_path)]) == 1


def test_sync_mail_default_worksheet_title_is_sanitised(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    fake_client = MagicMock()
    fake_client.get_user_messages = AsyncMock(return_value=demo_mail_messages())
    monkeypatch.setattr(cli, "_make_mail_client", lambda: fake_client)
    recorded: dict[str, Any] = {}

    class FakeExporter:
        def write_rows(self, rows: list[dict[str, Any]], title: str, **kwargs: Any) -> int:
            recorded["title"] = title
            return len(rows)

    monkeypatch.setattr(cli, "_make_sheets_exporter", lambda *args: FakeExporter())
    monkeypatch.setenv("GCS_SERVICE_ACCOUNT_KEY_JSON", "{}")
    monkeypatch.setenv("TARGET_GOOGLE_SHEET_ID", "sheet-1")
    assert cli.main(["sync", "mail", "--user", "ada/b"]) == 0
    assert recorded["title"] == "mail-ada_b"
