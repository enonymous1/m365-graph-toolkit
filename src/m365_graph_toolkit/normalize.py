"""Normalising Graph SDK objects into plain dicts.

The v1 codebase mixed SDK model objects and dicts freely: some savers
expected ``.get()``, others expected attributes, and several call sites
crashed on whichever shape they were handed. In v2 every client returns
plain ``dict`` rows built by :func:`model_to_dict`, so exporters,
analysers, and the CLI all see one shape.

Rules:

* ``None`` values are kept as ``None`` (callers decide how to render
  them); they are never silently turned into ``""`` or ``"null"``.
* ``datetime`` values become ISO-8601 strings.
* Lists and nested dicts are kept as-is for JSON output; the CSV writer
  renders them as JSON strings (see :mod:`m365_graph_toolkit.exporters`).
* Unknown / unmodelled properties live on ``additional_data`` in Kiota
  models; :func:`model_to_dict` only reads the fields it is asked for,
  so callers must name the fields they need explicitly.
"""

from __future__ import annotations

import re
from collections.abc import Iterable, Mapping
from datetime import datetime
from typing import Any


def scalar_value(value: Any) -> Any:
    """Convert one SDK value into a JSON-friendly scalar/container."""
    if isinstance(value, datetime):
        return value.isoformat()
    return value


def model_to_dict(obj: Any, fields: Iterable[str]) -> dict[str, Any]:
    """Pick ``fields`` off an SDK model object or a plain dict.

    Missing attributes/keys become ``None`` instead of raising
    :class:`AttributeError`, because Graph routinely omits properties
    the caller did not ``$select``. A missing field is data, not a bug.
    """
    result: dict[str, Any] = {}
    if isinstance(obj, Mapping):
        for field in fields:
            result[field] = scalar_value(obj.get(field))
        return result
    for field in fields:
        result[field] = scalar_value(getattr(obj, field, None))
    return result


def flatten_dict(
    data: Mapping[str, Any], *, parent_key: str = "", sep: str = "_"
) -> dict[str, Any]:
    """Flatten nested dicts one level at a time.

    Unlike v1, lists are left untouched: flattening a list of dicts
    into ``key_0``/``key_1`` columns silently drops structure and
    makes column sets depend on row order. Callers that need list
    columns should expand them explicitly.
    """
    items: dict[str, Any] = {}
    for key, value in data.items():
        new_key = f"{parent_key}{sep}{key}" if parent_key else key
        if isinstance(value, Mapping):
            items.update(flatten_dict(value, parent_key=new_key, sep=sep))
        else:
            items[new_key] = value
    return items


_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*]')


def sanitize_filename(name: str) -> str:
    """Make ``name`` safe to use as (part of) a file name."""
    cleaned = _INVALID_FILENAME_CHARS.sub("_", name)
    cleaned = re.sub(r"\s+", "_", cleaned.strip())
    if not cleaned or cleaned in {".", ".."}:
        raise ValueError("filename must contain at least one valid character")
    return cleaned


def require_non_empty(value: str, *, name: str) -> str:
    """Return ``value`` or raise :class:`ValidationError` if it is blank."""
    from .errors import ValidationError

    if not isinstance(value, str) or not value.strip():
        raise ValidationError(f"{name} must be a non-empty string")
    return value


__all__ = [
    "flatten_dict",
    "model_to_dict",
    "require_non_empty",
    "sanitize_filename",
    "scalar_value",
]
