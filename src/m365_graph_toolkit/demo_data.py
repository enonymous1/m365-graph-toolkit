"""Fictional demo data for the offline ``demo`` command and the tests.

Every person, address, and tenant here is invented. Names are
"Example" surnames, mail domains are ``example.com`` /
``example.onmicrosoft.com``, and the tenant is "Contoso" — Microsoft's
long-standing fictional company — on an ``.example`` domain. Nothing
in this module corresponds to a real user, group, site, or tenant.
Real exports must never be committed to this repository; see README.
"""

from __future__ import annotations

from typing import Any


def demo_users() -> list[dict[str, Any]]:
    """Three fictional users."""
    return [
        {
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
            "business_phones": ["+1 555 010 0001"],
            "created_date_time": "2024-01-15T09:00:00+00:00",
        },
        {
            "id": "22222222-2222-2222-2222-222222222222",
            "display_name": "Ben Example",
            "given_name": "Ben",
            "surname": "Example",
            "user_principal_name": "ben.example@example.onmicrosoft.com",
            "mail": "ben.example@example.com",
            "job_title": "Manager",
            "department": "Operations",
            "office_location": "Building A",
            "account_enabled": True,
            "user_type": "Member",
            "mobile_phone": None,
            "business_phones": [],
            "created_date_time": "2023-06-01T09:00:00+00:00",
        },
        {
            "id": "33333333-3333-3333-3333-333333333333",
            "display_name": "Guest Example",
            "given_name": "Guest",
            "surname": "Example",
            "user_principal_name": "guest.example_external@example.onmicrosoft.com",
            "mail": "guest@example.org",
            "job_title": None,
            "department": None,
            "office_location": None,
            "account_enabled": True,
            "user_type": "Guest",
            "mobile_phone": None,
            "business_phones": [],
            "created_date_time": "2024-03-01T09:00:00+00:00",
        },
    ]


def demo_groups() -> list[dict[str, Any]]:
    """Two fictional groups."""
    return [
        {
            "id": "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
            "display_name": "Operations",
            "description": "Operations team",
            "mail": "operations@example.com",
            "mail_nickname": "operations",
            "mail_enabled": True,
            "visibility": "Private",
            "group_types": ["Unified"],
            "created_date_time": "2023-01-01T09:00:00+00:00",
        },
        {
            "id": "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb",
            "display_name": "All Staff",
            "description": "Everyone",
            "mail": "allstaff@example.com",
            "mail_nickname": "allstaff",
            "mail_enabled": True,
            "visibility": "Private",
            "group_types": ["Unified"],
            "created_date_time": "2023-01-01T09:00:00+00:00",
        },
    ]


def demo_memberships() -> list[dict[str, Any]]:
    """Fictional user -> group rows, as ``get_user_memberships_batch`` returns them."""
    users = {user["id"]: user for user in demo_users()}
    groups = {group["id"]: group for group in demo_groups()}
    pairs = [
        ("11111111-1111-1111-1111-111111111111", "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"),
        ("11111111-1111-1111-1111-111111111111", "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"),
        ("22222222-2222-2222-2222-222222222222", "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"),
        ("22222222-2222-2222-2222-222222222222", "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"),
        ("33333333-3333-3333-3333-333333333333", "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"),
    ]
    return [
        {
            "user_id": user_id,
            "user_display_name": users[user_id]["display_name"],
            "user_principal_name": users[user_id]["user_principal_name"],
            "group_id": group_id,
            "group_display_name": groups[group_id]["display_name"],
            "group_type": "Unified",
        }
        for user_id, group_id in pairs
    ]


def demo_sites() -> list[dict[str, Any]]:
    """Two fictional SharePoint sites."""
    return [
        {
            "id": "contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222",
            "display_name": "Operations Hub",
            "name": "operations-hub",
            "description": "Fictional operations site",
            "web_url": "https://contoso.example/sites/operations-hub",
            "created_date_time": "2023-02-01T09:00:00+00:00",
            "last_modified_date_time": "2024-05-01T09:00:00+00:00",
        },
        {
            "id": "contoso.sharepoint.com,33333333-3333-3333-3333-333333333333,44444444-4444-4444-4444-444444444444",
            "display_name": "Team Site",
            "name": "team-site",
            "description": "Fictional team site",
            "web_url": "https://contoso.example/sites/team-site",
            "created_date_time": "2023-02-01T09:00:00+00:00",
            "last_modified_date_time": "2024-05-01T09:00:00+00:00",
        },
    ]


def demo_onedrive_stats() -> list[dict[str, Any]]:
    """Fictional per-user OneDrive stats, including one error entry."""
    return [
        {
            "user_id": "11111111-1111-1111-1111-111111111111",
            "display_name": "Ada Example",
            "storage": {
                "used_gb": 12.5,
                "total_gb": 100.0,
                "remaining_gb": 87.5,
                "state": "normal",
            },
            "items": {"total_items": 42, "folders": 5, "files": 37},
        },
        {
            "user_id": "22222222-2222-2222-2222-222222222222",
            "display_name": "Ben Example",
            "storage": {
                "used_gb": 80.0,
                "total_gb": 100.0,
                "remaining_gb": 20.0,
                "state": "normal",
            },
            "items": {"total_items": 120, "folders": 10, "files": 110},
        },
        {
            "user_id": "33333333-3333-3333-3333-333333333333",
            "display_name": "Guest Example",
            "error": "OneDrive not provisioned for this user",
        },
    ]


def demo_sign_in_logs() -> list[dict[str, Any]]:
    """Fictional sign-in rows: 3 successful, 1 failed."""
    base = {
        "app_display_name": "Fictional App",
        "app_id": "55555555-5555-5555-5555-555555555555",
        "ip_address": "203.0.113.10",
        "client_app_used": "Browser",
        "is_interactive": True,
        "risk_state": "none",
        "risk_level_aggregated": "low",
        "risk_level_during_sign_in": "low",
    }
    return [
        {
            **base,
            "id": "log-1",
            "created_date_time": "2026-10-20T09:00:00+00:00",
            "user_display_name": "Ada Example",
            "user_principal_name": "ada.example@example.onmicrosoft.com",
            "user_id": "11111111-1111-1111-1111-111111111111",
            "status": {"error_code": 0, "failure_reason": None},
        },
        {
            **base,
            "id": "log-2",
            "created_date_time": "2026-10-20T09:05:00+00:00",
            "user_display_name": "Ada Example",
            "user_principal_name": "ada.example@example.onmicrosoft.com",
            "user_id": "11111111-1111-1111-1111-111111111111",
            "status": {"error_code": 0, "failure_reason": None},
        },
        {
            **base,
            "id": "log-3",
            "created_date_time": "2026-10-20T09:10:00+00:00",
            "user_display_name": "Ben Example",
            "user_principal_name": "ben.example@example.onmicrosoft.com",
            "user_id": "22222222-2222-2222-2222-222222222222",
            "status": {"error_code": 0, "failure_reason": None},
        },
        {
            **base,
            "id": "log-4",
            "created_date_time": "2026-10-20T09:15:00+00:00",
            "user_display_name": "Ben Example",
            "user_principal_name": "ben.example@example.onmicrosoft.com",
            "user_id": "22222222-2222-2222-2222-222222222222",
            "status": {"error_code": 50126, "failure_reason": "Invalid username or password."},
        },
    ]


__all__ = [
    "demo_groups",
    "demo_memberships",
    "demo_onedrive_stats",
    "demo_sign_in_logs",
    "demo_sites",
    "demo_users",
]
