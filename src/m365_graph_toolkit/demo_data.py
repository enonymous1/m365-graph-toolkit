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
    "demo_channel_messages",
    "demo_events",
    "demo_groups",
    "demo_mail_messages",
    "demo_memberships",
    "demo_onedrive_stats",
    "demo_planner_tasks",
    "demo_plans",
    "demo_sign_in_logs",
    "demo_sites",
    "demo_stig_plan",
    "demo_users",
]


def demo_plans() -> list[dict[str, Any]]:
    """Two fictional Planner plans."""
    return [
        {
            "id": "plan-1",
            "title": "Office Move",
            "created_date_time": "2024-02-01T09:00:00+00:00",
            "created_by_user_id": "11111111-1111-1111-1111-111111111111",
            "container_id": "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
            "container_type": "group",
            "container_url": "https://contoso.example/planner/plan-1",
            "group_name": "Operations",
        },
        {
            "id": "plan-2",
            "title": "Audit Prep",
            "created_date_time": "2024-03-01T09:00:00+00:00",
            "created_by_user_id": "22222222-2222-2222-2222-222222222222",
            "container_id": "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
            "container_type": "group",
            "container_url": "https://contoso.example/planner/plan-2",
            "group_name": "Operations",
        },
    ]


def demo_planner_tasks() -> list[dict[str, Any]]:
    """Two fictional Planner tasks, as ``PlannerClient.get_plan_tasks`` returns them."""
    return [
        {
            "id": "task-1",
            "title": "Book movers",
            "bucket_id": "bucket-1",
            "percent_complete": 50,
            "priority": 5,
            "created_date_time": "2024-02-02T09:00:00+00:00",
            "start_date_time": "2024-02-05T09:00:00+00:00",
            "due_date_time": "2024-02-20T17:00:00+00:00",
            "completed_date_time": None,
            "plan_id": "plan-1",
            "bucket_name": "To do",
            "created_by_user_id": "11111111-1111-1111-1111-111111111111",
            "completed_by_user_id": None,
            "assigned_to": ["22222222-2222-2222-2222-222222222222"],
            "labels": ["category1"],
            "description": "Get three quotes before booking.",
            "checklist_items": [
                {"title": "Request quotes", "is_checked": True},
                {"title": "Compare quotes", "is_checked": False},
            ],
            "checklist_total": 2,
            "checklist_completed": 1,
        },
        {
            "id": "task-2",
            "title": "Gather evidence",
            "bucket_id": "bucket-2",
            "percent_complete": 100,
            "priority": 3,
            "created_date_time": "2024-03-02T09:00:00+00:00",
            "start_date_time": None,
            "due_date_time": "2024-03-15T17:00:00+00:00",
            "completed_date_time": "2024-03-14T16:00:00+00:00",
            "plan_id": "plan-2",
            "bucket_name": "Done",
            "created_by_user_id": "22222222-2222-2222-2222-222222222222",
            "completed_by_user_id": "22222222-2222-2222-2222-222222222222",
            "assigned_to": [],
            "labels": [],
            "description": None,
            "checklist_items": [],
            "checklist_total": 0,
            "checklist_completed": 0,
        },
    ]


def demo_events() -> list[dict[str, Any]]:
    """Two fictional calendar events."""
    return [
        {
            "id": "event-1",
            "subject": "Operations standup",
            "body_preview": "Daily standup.",
            "importance": "normal",
            "is_all_day": False,
            "is_cancelled": False,
            "is_online_meeting": True,
            "online_meeting_url": "https://contoso.example/meet/standup",
            "created_date_time": "2024-01-10T09:00:00+00:00",
            "last_modified_date_time": "2024-01-10T09:00:00+00:00",
            "start_date_time": "2026-10-28T14:00:00",
            "start_time_zone": "UTC",
            "end_date_time": "2026-10-28T14:30:00",
            "end_time_zone": "UTC",
            "location": "Virtual",
            "organizer_email": "ada.example@example.com",
            "attendee_emails": ["ben.example@example.com"],
        },
        {
            "id": "event-2",
            "subject": "Site visit",
            "body_preview": "Quarterly site visit.",
            "importance": "high",
            "is_all_day": True,
            "is_cancelled": False,
            "is_online_meeting": False,
            "online_meeting_url": None,
            "created_date_time": "2024-01-11T09:00:00+00:00",
            "last_modified_date_time": "2024-01-11T09:00:00+00:00",
            "start_date_time": "2026-10-29T00:00:00",
            "start_time_zone": "UTC",
            "end_date_time": "2026-10-29T23:59:00",
            "end_time_zone": "UTC",
            "location": "Building A",
            "organizer_email": "ben.example@example.com",
            "attendee_emails": [],
        },
    ]


def demo_mail_messages() -> list[dict[str, Any]]:
    """Two fictional Outlook messages."""
    return [
        {
            "id": "mail-1",
            "subject": "Q3 numbers",
            "body_preview": "Attached are the Q3 numbers.",
            "importance": "normal",
            "is_read": True,
            "has_attachments": True,
            "received_date_time": "2026-10-20T09:00:00+00:00",
            "sent_date_time": "2026-10-20T08:55:00+00:00",
            "created_date_time": "2026-10-20T09:00:00+00:00",
            "from_email": "ada.example@example.com",
            "from_name": "Ada Example",
            "to_emails": ["ben.example@example.com"],
            "cc_emails": [],
            "bcc_emails": [],
            "body": "<p>Attached are the Q3 numbers.</p>",
            "body_content_type": "html",
        },
        {
            "id": "mail-2",
            "subject": "Welcome",
            "body_preview": "Welcome to the team.",
            "importance": "low",
            "is_read": False,
            "has_attachments": False,
            "received_date_time": "2026-10-21T09:00:00+00:00",
            "sent_date_time": "2026-10-21T08:00:00+00:00",
            "created_date_time": "2026-10-21T09:00:00+00:00",
            "from_email": "ben.example@example.com",
            "from_name": "Ben Example",
            "to_emails": ["ada.example@example.com"],
            "cc_emails": ["guest@example.org"],
            "bcc_emails": [],
            "body": "Welcome to the team.",
            "body_content_type": "text",
        },
    ]


def demo_channel_messages() -> list[dict[str, Any]]:
    """Two fictional Teams channel messages."""
    return [
        {
            "id": "chat-1",
            "subject": None,
            "importance": "normal",
            "message_type": "message",
            "created_date_time": "2026-10-20T09:00:00+00:00",
            "last_modified_date_time": "2026-10-20T09:00:00+00:00",
            "deleted_date_time": None,
            "body": "<p>Standup notes are posted.</p>",
            "body_content_type": "html",
            "sender_type": "user",
            "sender_id": "11111111-1111-1111-1111-111111111111",
            "sender_display_name": "Ada Example",
            "reply_count": 1,
            "reaction_count": 1,
            "reaction_types": ["like"],
            "team_id": "team-1",
            "channel_id": "channel-1",
        },
        {
            "id": "chat-2",
            "subject": None,
            "importance": "normal",
            "message_type": "message",
            "created_date_time": "2026-10-20T09:10:00+00:00",
            "last_modified_date_time": "2026-10-20T09:10:00+00:00",
            "deleted_date_time": None,
            "body": "Thanks!",
            "body_content_type": "text",
            "sender_type": "user",
            "sender_id": "22222222-2222-2222-2222-222222222222",
            "sender_display_name": "Ben Example",
            "reply_count": 0,
            "reaction_count": 0,
            "reaction_types": [],
            "team_id": "team-1",
            "channel_id": "channel-1",
        },
    ]


def demo_stig_plan() -> dict[str, Any]:
    """A fictional STIG list plan, in the shape ``stig.plan_lists_from_directory`` returns."""
    return {
        "list_name": "Fictional Example STIG",
        "source_file": "Fictional Example STIG.csv",
        "row_count": 2,
        "title_header": "Vuln ID",
        "columns": [
            {"header": "Vuln ID", "is_title": True},
            {"header": "Severity", "is_title": False},
            {"header": "Rule Title", "is_title": False},
            {"header": "Fix Text", "is_title": False},
        ],
        "skipped_headers": ["_UIVersionString"],
    }
