"""STIG planning tests (pure module, no Graph dependency)."""

from __future__ import annotations

from pathlib import Path

import pytest

from m365_graph_toolkit.errors import ValidationError
from m365_graph_toolkit.stig import (
    discover_csv_files,
    load_stig_csv,
    map_row_to_fields,
    plan_list_from_csv,
    plan_lists_from_directory,
    resolve_column_map,
)

SAMPLE_CSV = Path(__file__).resolve().parents[1] / "examples" / "Fictional Example STIG.csv"


def _write_csv(path: Path, text: str) -> Path:
    path.write_text(text, encoding="utf-8")
    return path


def test_sample_csv_plans_as_expected() -> None:
    plan = plan_list_from_csv(SAMPLE_CSV)
    assert plan["list_name"] == "Fictional Example STIG"
    assert plan["row_count"] == 2
    assert plan["title_header"] == "Vuln ID"
    assert plan["skipped_headers"] == ["_UIVersionString"]
    assert [c["header"] for c in plan["columns"]] == [
        "STIG",
        "Benchmark ID",
        "Release Info",
        "Version",
        "Vuln ID",
        "Severity",
        "Rule ID",
        "STIG ID",
        "Classification",
        "Rule Title",
        "Fix Text",
        "Discussion",
    ]
    assert [c for c in plan["columns"] if c["is_title"]][0]["header"] == "Vuln ID"


def test_plan_maps_vuln_id_case_insensitively(tmp_path: Path) -> None:
    csv_file = _write_csv(
        tmp_path / "Case.csv",
        "vuln id,Severity\nV-1,high\n",
    )
    plan = plan_list_from_csv(csv_file)
    assert plan["title_header"] == "vuln id"


def test_plan_rejects_missing_vuln_id(tmp_path: Path) -> None:
    csv_file = _write_csv(tmp_path / "NoTitle.csv", "Severity,Rule Title\nhigh,Do it\n")
    with pytest.raises(ValidationError, match="Vuln ID"):
        plan_list_from_csv(csv_file)


def test_plan_rejects_duplicate_headers(tmp_path: Path) -> None:
    csv_file = _write_csv(tmp_path / "Dup.csv", "Vuln ID,Severity, severity\nV-1,high,low\n")
    with pytest.raises(ValidationError, match="duplicate header"):
        plan_list_from_csv(csv_file)


def test_plan_rejects_header_only_csv(tmp_path: Path) -> None:
    csv_file = _write_csv(tmp_path / "Empty.csv", "Vuln ID,Severity\n")
    with pytest.raises(ValidationError, match="no data rows"):
        plan_list_from_csv(csv_file)


def test_plan_rejects_empty_and_missing_files(tmp_path: Path) -> None:
    empty = _write_csv(tmp_path / "blank.csv", "")
    with pytest.raises(ValidationError, match="empty"):
        plan_list_from_csv(empty)
    with pytest.raises(ValidationError, match="not found"):
        plan_list_from_csv(tmp_path / "nope.csv")


def test_load_stig_csv_returns_plan_and_rows() -> None:
    plan, rows = load_stig_csv(SAMPLE_CSV)
    assert plan["row_count"] == len(rows) == 2
    assert rows[0]["Vuln ID"] == "V-999001"
    assert rows[0]["Severity"] == "medium"


def test_discover_and_plan_directory(tmp_path: Path) -> None:
    _write_csv(tmp_path / "B.csv", "Vuln ID,Severity\nV-2,low\n")
    _write_csv(tmp_path / "A.csv", "Vuln ID,Severity\nV-1,high\n")
    (tmp_path / "notes.txt").write_text("not a csv", encoding="utf-8")
    assert [p.name for p in discover_csv_files(tmp_path)] == ["A.csv", "B.csv"]
    plans = plan_lists_from_directory(tmp_path)
    assert [p["list_name"] for p in plans] == ["A", "B"]
    empty_dir = tmp_path / "empty"
    empty_dir.mkdir()
    with pytest.raises(ValidationError, match="No CSV files"):
        plan_lists_from_directory(empty_dir)
    with pytest.raises(ValidationError, match="not found"):
        plan_lists_from_directory(tmp_path / "missing")


def test_resolve_column_map() -> None:
    plan = plan_list_from_csv(SAMPLE_CSV)
    sp_columns = [{"name": "Title", "display_name": "Title", "is_text": False}]
    for column in plan["columns"]:
        if column["is_title"]:
            continue
        header = str(column["header"])
        sp_columns.append(
            {
                "name": header.replace(" ", "_x0020_"),
                "display_name": header,
                "is_text": True,
            }
        )
    column_map = resolve_column_map(plan, sp_columns)
    assert column_map["Vuln ID"] == "Title"
    assert column_map["Rule Title"] == "Rule_x0020_Title"
    assert column_map["Fix Text"] == "Fix_x0020_Text"


def test_resolve_column_map_rejects_missing_or_non_text() -> None:
    plan = plan_list_from_csv(SAMPLE_CSV)
    with pytest.raises(ValidationError, match="Title"):
        resolve_column_map(plan, [])
    columns = [
        {"name": "Title", "display_name": "Title", "is_text": False},
        {"name": "Severity", "display_name": "Severity", "is_text": False},
    ]
    with pytest.raises(ValidationError, match="not provisioned"):
        resolve_column_map(plan, columns)


def test_map_row_to_fields_drops_unmapped() -> None:
    plan = plan_list_from_csv(SAMPLE_CSV)
    column_map = {"Vuln ID": "Title", "Severity": "Severity"}
    _, rows = load_stig_csv(SAMPLE_CSV)
    fields = map_row_to_fields(plan, rows[0], column_map)
    assert fields == {"Title": "V-999001", "Severity": "medium"}
    assert map_row_to_fields(plan, rows[0], {}) == {}
