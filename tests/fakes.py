"""Shared fake Graph objects for the offline test suite."""

from __future__ import annotations

from types import SimpleNamespace
from typing import Any


def page(items: list[Any] | None, *, next_link: str | None = None) -> SimpleNamespace:
    """A fake Graph collection page: ``value`` + ``odata_next_link``."""
    return SimpleNamespace(value=items, odata_next_link=next_link)


def user_obj(**overrides: Any) -> SimpleNamespace:
    """A fake SDK User object with the fields the clients read."""
    data = {
        "id": "11111111-1111-1111-1111-111111111111",
        "display_name": "Ada Example",
        "given_name": "Ada",
        "surname": "Example",
        "user_principal_name": "ada.example@example.onmicrosoft.com",
        "mail": "ada.example@example.com",
        "job_title": "Analyst",
        "department": "Operations",
        "office_location": "Building A",
        "account_enabled": True,
        "user_type": "Member",
        "mobile_phone": None,
        "business_phones": [],
        "created_date_time": None,
    }
    data.update(overrides)
    return SimpleNamespace(**data)


def group_obj(**overrides: Any) -> SimpleNamespace:
    """A fake SDK Group / directory-object."""
    data = {
        "id": "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
        "display_name": "Operations",
        "mail": "operations@example.com",
        "group_types": ["Unified"],
        "visibility": "Private",
        "odata_type": "#microsoft.graph.group",
    }
    data.update(overrides)
    return SimpleNamespace(**data)


__all__ = ["group_obj", "page", "user_obj"]
