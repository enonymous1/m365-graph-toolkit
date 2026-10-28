"""STIG CSV planning: turn DISA STIG exports into SharePoint-list plans.

This module is pure — no Graph calls, no network, no file writes
outside the caller's hands. It covers the planning half of the old
``mp-m365-app`` STIG importer (``graph/sharepoint_lists.py``); the
write half lives in :class:`SharePointClient
<m365_graph_toolkit.clients.sharepoint.SharePointClient>` as
``import_list_from_csv`` / ``import_lists_from_directory`` (library
only — there is no CLI path that writes to a tenant).

Rules carried over from the old importer, made explicit and
fail-loud:

* The CSV header ``Vuln ID`` (case-insensitive, surrounding
  whitespace ignored) maps to SharePoint's ``Title`` column. Every
  other non-empty header becomes a multi-line text column.
* Read-only system fields are never written: ``_UIVersionString``
  (case-insensitive) is skipped both when defining columns and when
  mapping rows.
* Duplicate headers (after stripping and lower-casing) are an
  error. ``csv.DictReader`` silently collapses them, which would
  drop a column's data without a trace.
* A CSV with no headers, or with headers but no data rows, is an
  error for import planning — the old code created an empty list
  and called it success.

STIG CSVs are **not** bundled with this repository. They are
published by DISA, updated quarterly, and 46 MB in bulk in the old
tree. Download the current STIG library export from
``https://public.cyber.mil/stigs/`` (or your organisation's
approved source), unzip it into a directory, and point the planner
at that directory. The sample file in ``examples/Fictional Example STIG.csv``
is fictional data in STIG column shape, for offline preview only.
"""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Any

from .errors import ValidationError

TITLE_HEADER = "vuln id"
READ_ONLY_HEADERS = frozenset({"_uiversionstring"})
STIG_LIBRARY_URL = "https://public.cyber.mil/stigs/"


def _read_csv(path: str | Path) -> tuple[list[str], list[dict[str, Any]]]:
    """Read a CSV as ``(headers, rows)`` with duplicate-header checks.

    Headers are stripped; cell values are kept as-is. Short rows are
    padded with ``""`` and long rows are truncated, so every row has
    exactly one value per header.
    """
    csv_path = Path(path)
    if not csv_path.is_file():
        raise ValidationError(f"CSV file not found: {csv_path}")
    try:
        with csv_path.open(newline="", encoding="utf-8-sig") as handle:
            reader = csv.reader(handle)
            try:
                raw_headers = next(reader)
            except StopIteration:
                raise ValidationError(f"CSV file is empty: {csv_path}") from None
            headers = [header.strip() for header in raw_headers]
            if not any(headers):
                raise ValidationError(f"CSV file has no headers: {csv_path}")
            seen: dict[str, str] = {}
            for header in headers:
                key = header.lower()
                if header and key in seen:
                    raise ValidationError(f"CSV file {csv_path} has duplicate header {header!r}")
                if header:
                    seen[key] = header
            rows: list[dict[str, Any]] = []
            for values in reader:
                if not values and not rows:
                    continue
                padded = list(values) + [""] * (len(headers) - len(values))
                rows.append(dict(zip(headers, padded[: len(headers)], strict=True)))
    except UnicodeDecodeError as exc:
        raise ValidationError(f"CSV file {csv_path} is not valid UTF-8: {exc}") from exc
    except csv.Error as exc:
        raise ValidationError(f"CSV file {csv_path} could not be parsed: {exc}") from exc
    return headers, rows


def _plan_from_headers(csv_path: Path, headers: list[str], row_count: int) -> dict[str, Any]:
    list_name = csv_path.stem.strip()
    if not list_name:
        raise ValidationError(f"CSV file {csv_path} has no usable file-stem list name")
    if len(list_name) > 255:
        raise ValidationError(f"List name derived from {csv_path} exceeds 255 characters")
    if row_count == 0:
        raise ValidationError(
            f"CSV file {csv_path} has no data rows; refusing to plan an empty list"
        )
    title_header: str | None = None
    columns: list[dict[str, Any]] = []
    skipped: list[str] = []
    for header in headers:
        if not header:
            skipped.append(header)
            continue
        lowered = header.lower()
        if lowered in READ_ONLY_HEADERS:
            skipped.append(header)
            continue
        if lowered == TITLE_HEADER:
            title_header = header
            columns.append({"header": header, "is_title": True})
            continue
        columns.append({"header": header, "is_title": False})
    if title_header is None:
        raise ValidationError(
            f"CSV file {csv_path} has no 'Vuln ID' header; a Title mapping is required"
        )
    return {
        "list_name": list_name,
        "source_file": csv_path.name,
        "row_count": row_count,
        "title_header": title_header,
        "columns": columns,
        "skipped_headers": skipped,
    }


def plan_list_from_csv(path: str | Path) -> dict[str, Any]:
    """Build a list plan from one STIG CSV, without reading its rows twice.

    Returns a plan dict: ``list_name`` (the file stem), ``source_file``
    (file name), ``row_count``, ``title_header`` (the ``Vuln ID``
    header as written, or ``None``), ``columns`` (one entry per
    header that becomes a column), and ``skipped_headers`` (empty
    and read-only headers that were dropped).
    """
    csv_path = Path(path)
    headers, rows = _read_csv(csv_path)
    return _plan_from_headers(csv_path, headers, len(rows))


def load_stig_csv(path: str | Path) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    """Return ``(plan, rows)`` for one STIG CSV (single read)."""
    csv_path = Path(path)
    headers, rows = _read_csv(csv_path)
    return _plan_from_headers(csv_path, headers, len(rows)), rows


def discover_csv_files(directory: str | Path) -> list[Path]:
    """Return the ``.csv`` files in ``directory``, sorted by name."""
    directory_path = Path(directory)
    if not directory_path.is_dir():
        raise ValidationError(f"CSV directory not found: {directory_path}")
    files = sorted(
        (entry for entry in directory_path.iterdir() if entry.is_file()),
        key=lambda entry: entry.name.lower(),
    )
    csv_files = [entry for entry in files if entry.suffix.lower() == ".csv"]
    if not csv_files:
        raise ValidationError(f"No CSV files found in {directory_path}")
    return csv_files


def plan_lists_from_directory(directory: str | Path) -> list[dict[str, Any]]:
    """Build a list plan for every CSV in ``directory``.

    A file that fails planning aborts the whole call with
    :class:`ValidationError` naming the file — a directory of STIG
    exports is expected to be uniform, and silently skipping a
    broken one would hide a failed import.
    """
    return [plan_list_from_csv(csv_file) for csv_file in discover_csv_files(directory)]


def resolve_column_map(plan: dict[str, Any], sp_columns: list[dict[str, Any]]) -> dict[str, str]:
    """Map plan headers to SharePoint internal column names.

    ``sp_columns`` are rows from ``SharePointClient.get_list_columns``
    (``name``, ``display_name``, ``is_text``). The ``Vuln ID`` header
    maps to internal name ``Title``; every other header maps to the
    column whose display name matches it (case-insensitive). A
    missing column, or a column that is not a text column, raises
    :class:`ValidationError` — the list was not provisioned as
    planned and items must not be written against guessed names.
    """
    by_display = {
        str(column.get("display_name", "")).strip().lower(): column
        for column in sp_columns
        if column.get("display_name")
    }
    title_column = next((column for column in sp_columns if column.get("name") == "Title"), None)
    if title_column is None:
        raise ValidationError("SharePoint list has no 'Title' column")
    column_map: dict[str, str] = {}
    missing: list[str] = []
    for column in plan["columns"]:
        header = str(column["header"])
        if column["is_title"]:
            column_map[header] = "Title"
            continue
        match = by_display.get(header.strip().lower())
        if match is None or not match.get("is_text"):
            missing.append(header)
            continue
        internal = match.get("name")
        if not internal:
            missing.append(header)
            continue
        column_map[header] = str(internal)
    if missing:
        raise ValidationError(
            "SharePoint columns not provisioned as text columns: " + ", ".join(missing)
        )
    return column_map


def map_row_to_fields(
    plan: dict[str, Any], row: dict[str, Any], column_map: dict[str, str]
) -> dict[str, Any]:
    """Map one CSV row to SharePoint field values via ``column_map``.

    Headers with no entry in ``column_map`` (skipped or
    unprovisioned) contribute nothing. A row that maps to no fields
    at all returns ``{}``; callers skip it and count it as failed
    rather than creating an empty item.
    """
    fields: dict[str, Any] = {}
    for header, internal_name in column_map.items():
        if header in row:
            fields[internal_name] = row[header]
    return fields


__all__ = [
    "READ_ONLY_HEADERS",
    "STIG_LIBRARY_URL",
    "TITLE_HEADER",
    "discover_csv_files",
    "load_stig_csv",
    "map_row_to_fields",
    "plan_list_from_csv",
    "plan_lists_from_directory",
    "resolve_column_map",
]
