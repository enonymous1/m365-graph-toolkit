from __future__ import annotations

from datetime import datetime, timezone
from types import SimpleNamespace

import pytest

from m365_graph_toolkit.errors import ValidationError
from m365_graph_toolkit.normalize import (
    flatten_dict,
    model_to_dict,
    require_non_empty,
    sanitize_filename,
)


def test_model_to_dict_from_object_and_dict() -> None:
    obj = SimpleNamespace(id="1", display_name="Ada", missing=None)
    assert model_to_dict(obj, ["id", "display_name", "absent"]) == {
        "id": "1",
        "display_name": "Ada",
        "absent": None,
    }
    assert model_to_dict({"id": "1"}, ["id", "absent"]) == {"id": "1", "absent": None}


def test_model_to_dict_converts_datetimes() -> None:
    stamp = datetime(2026, 10, 20, 9, 0, tzinfo=timezone.utc)
    assert model_to_dict({"created": stamp}, ["created"]) == {
        "created": "2026-10-20T09:00:00+00:00"
    }


def test_flatten_dict_leaves_lists_alone() -> None:
    assert flatten_dict({"a": {"b": 1}, "tags": ["x", "y"]}) == {"a_b": 1, "tags": ["x", "y"]}


def test_sanitize_filename() -> None:
    assert sanitize_filename("My Site: Q1/Q2") == "My_Site__Q1_Q2"
    with pytest.raises(ValueError):
        sanitize_filename("..")
    with pytest.raises(ValueError):
        sanitize_filename("   ")


def test_require_non_empty() -> None:
    assert require_non_empty(" x ", name="thing") == " x "
    with pytest.raises(ValidationError):
        require_non_empty("  ", name="thing")
