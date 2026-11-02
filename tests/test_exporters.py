from __future__ import annotations

import csv
import json
from datetime import datetime, timezone
from pathlib import Path

import pytest

from m365_graph_toolkit.errors import ExportError, ValidationError
from m365_graph_toolkit.exporters import render_cell_value, write_csv, write_json


def test_render_cell_value_is_the_shared_sink_conversion() -> None:
    # CSV and Sheets must render a value the same way; this is the
    # one function both sinks use, so a drift here fails loudly.
    stamp = datetime(2026, 10, 20, 9, 0, tzinfo=timezone.utc)
    assert render_cell_value(None) == ""
    assert render_cell_value(stamp) == "2026-10-20T09:00:00+00:00"
    assert render_cell_value({"a": 1}) == '{"a": 1}'
    assert render_cell_value(42) == 42
    assert render_cell_value(True) is True


def test_write_csv_does_not_mutate_input(tmp_path: Path) -> None:
    rows = [{"id": "1", "storage": {"used_gb": 1.5}, "items": {"files": 3}}]
    path = write_csv(rows, tmp_path, "usage")
    # v1 pop()ed nested keys out of the caller's dicts.
    assert rows == [{"id": "1", "storage": {"used_gb": 1.5}, "items": {"files": 3}}]
    with path.open(newline="", encoding="utf-8") as handle:
        saved = list(csv.DictReader(handle))
    assert saved[0]["storage"] == '{"used_gb": 1.5}'
    assert saved[0]["items"] == '{"files": 3}'


def test_write_csv_empty_without_fieldnames_raises(tmp_path: Path) -> None:
    with pytest.raises(ValidationError):
        write_csv([], tmp_path, "empty")


def test_write_csv_empty_with_fieldnames_writes_header_only(tmp_path: Path) -> None:
    path = write_csv([], tmp_path, "empty", fieldnames=["id", "name"])
    with path.open(newline="", encoding="utf-8") as handle:
        rows = list(csv.reader(handle))
    assert rows == [["id", "name"]]


def test_write_csv_rejects_unknown_fields(tmp_path: Path) -> None:
    with pytest.raises(ValidationError):
        write_csv([{"id": "1", "surprise": 2}], tmp_path, "x", fieldnames=["id"])


def test_write_csv_renders_none_and_datetimes(tmp_path: Path) -> None:
    stamp = datetime(2026, 10, 20, 9, 0, tzinfo=timezone.utc)
    path = write_csv([{"id": "1", "when": stamp, "note": None}], tmp_path, "x")
    with path.open(newline="", encoding="utf-8") as handle:
        saved = list(csv.DictReader(handle))
    assert saved[0]["when"] == "2026-10-20T09:00:00+00:00"
    assert saved[0]["note"] == ""


def test_write_json_round_trips_and_keeps_types(tmp_path: Path) -> None:
    stamp = datetime(2026, 10, 20, 9, 0, tzinfo=timezone.utc)
    path = write_json([{"id": "1", "when": stamp, "note": None}], tmp_path, "x")
    data = json.loads(path.read_text(encoding="utf-8"))
    assert data[0]["when"] == "2026-10-20T09:00:00+00:00"
    assert data[0]["note"] is None


def test_write_json_non_serializable_creates_no_file(tmp_path: Path) -> None:
    with pytest.raises(ExportError):
        write_json({"bad": object()}, tmp_path, "x")
    assert list(tmp_path.iterdir()) == []
