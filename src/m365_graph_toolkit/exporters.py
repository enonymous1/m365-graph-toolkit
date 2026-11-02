"""CSV and JSON exporters: explicit, non-destructive, no hidden cache.

Fixes over v1's ``CsvFileManager`` / ``JsonFileManager`` / ``DataExporter``:

* The output directory is supplied by the caller. v1 wrote into a
  directory derived from ``__file__`` (inside the installed package)
  regardless of where the caller asked for the file.
* Input rows are never mutated. v1's CSV saver ``pop()``-ed ``storage``
  and ``items`` out of the caller's dicts.
* No file rotation. v1 silently deleted every file past the newest
  three in the output directory — data loss presented as housekeeping.
  Nothing here deletes a file it did not just write.
* No per-instance cache. v1's ``save_*`` helpers created a fresh
  manager on every call, so the cache branch could never hit; the
  "caching" was dead code that only made behaviour harder to predict.
* Empty input is an error unless the caller supplies field names, in
  which case a header-only CSV is written. v1's ``save_data([])``
  raised in one saver, wrote nothing in another, and crashed on
  ``lists_data[0]`` in a third.
* CSV values: datetimes become ISO strings, lists/dicts become JSON
  strings, ``None`` becomes ``""``. The same rows written to JSON keep
  their native types.
"""

from __future__ import annotations

import csv
import json
from collections.abc import Iterable
from datetime import datetime
from pathlib import Path
from typing import Any

from .errors import ExportError, ValidationError
from .normalize import sanitize_filename


def timestamped_path(out_dir: str | Path, stem: str, extension: str) -> Path:
    """Return ``out_dir/<sanitised stem>_YYYYMMDD_HHMMSS.<extension>``."""
    directory = Path(out_dir)
    safe_stem = sanitize_filename(stem)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return directory / f"{safe_stem}_{timestamp}.{extension.lstrip('.')}"


def render_cell_value(value: Any) -> Any:
    """Render one row value for a flat export (CSV cell or Sheets cell).

    Shared by :func:`write_csv` and the Sheets exporter so the two
    sinks can never drift apart on what a value becomes: ``None``
    becomes ``""``, datetimes become ISO strings, lists/dicts become
    JSON strings, and everything else keeps its type.
    """
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, default=str)
    return value


def field_names_for(rows: list[dict[str, Any]]) -> list[str]:
    """Field names in first-seen order across all rows."""
    names: list[str] = []
    seen: set[str] = set()
    for row in rows:
        for key in row:
            if key not in seen:
                seen.add(key)
                names.append(key)
    return names


def write_csv(
    rows: Iterable[dict[str, Any]],
    out_dir: str | Path,
    stem: str,
    *,
    fieldnames: list[str] | None = None,
) -> Path:
    """Write ``rows`` to a timestamped CSV file; return its path.

    Rows are copied, never mutated. Extra keys in a row that are not
    in ``fieldnames`` raise :class:`ValueError` via ``csv``'s default
    behaviour only when ``extrasaction`` would hide them — here we
    validate explicitly so a typo'd column is caught, not dropped.
    """
    row_list = [dict(row) for row in rows]
    if fieldnames is None:
        if not row_list:
            raise ValidationError("Cannot write an empty CSV without fieldnames")
        fieldnames = field_names_for(row_list)
    unknown = sorted({key for row in row_list for key in row} - set(fieldnames))
    if unknown:
        raise ValidationError(f"Rows contain fields not in fieldnames: {', '.join(unknown)}")
    path = timestamped_path(out_dir, stem, "csv")
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="raise")
            writer.writeheader()
            for row in row_list:
                writer.writerow({key: render_cell_value(row.get(key)) for key in fieldnames})
    except (OSError, csv.Error) as exc:
        raise ExportError(f"Failed to write CSV to {path}: {exc}") from exc
    return path


def _json_default(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.isoformat()
    raise TypeError(f"Object of type {type(value).__name__} is not JSON serializable")


def write_json(data: Any, out_dir: str | Path, stem: str) -> Path:
    """Write ``data`` to a timestamped JSON file; return its path.

    Non-serializable data raises :class:`TypeError` before any file is
    created, so a failed export never leaves a partial file behind.
    ``datetime`` values are rendered as ISO strings.
    """
    path = timestamped_path(out_dir, stem, "json")
    try:
        text = json.dumps(data, indent=2, ensure_ascii=False, default=_json_default)
    except TypeError as exc:
        raise ExportError(f"Data is not JSON serializable: {exc}") from exc
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text + "\n", encoding="utf-8")
    except OSError as exc:
        raise ExportError(f"Failed to write JSON to {path}: {exc}") from exc
    return path


__all__ = ["field_names_for", "render_cell_value", "timestamped_path", "write_csv", "write_json"]
