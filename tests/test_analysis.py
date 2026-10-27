from __future__ import annotations

import pytest

from m365_graph_toolkit.analysis import (
    analyze_onedrive_usage,
    analyze_sign_in_logs,
    analyze_team_membership,
    group_logs_by_user,
    is_internal_user,
)
from m365_graph_toolkit.demo_data import (
    demo_memberships,
    demo_onedrive_stats,
    demo_sign_in_logs,
)


def test_is_internal_user_checks_mail_or_upn() -> None:
    # v1 checked mail only and hardcoded the domain.
    assert is_internal_user({"mail": "ada@example.com", "userPrincipalName": None}, "example.com")
    assert is_internal_user(
        {"mail": None, "userPrincipalName": "ada@example.onmicrosoft.com"},
        "example.onmicrosoft.com",
    )
    assert not is_internal_user({"mail": "ada@notexample.com"}, "example.com")
    assert is_internal_user({"mail": "ADA@EXAMPLE.COM"}, "example.com")
    assert not is_internal_user({}, "example.com")
    assert not is_internal_user({"mail": "ada@example.com"}, "")


def test_analyze_team_membership_dedups_and_counts_empty() -> None:
    members_by_team = {
        "t1": {"team_name": "A", "members": [{"id": "u1"}, {"id": "u1"}, {"id": "u2"}]},
        "t2": {"team_name": "B", "members": [{"id": "u1"}]},
        "t3": {"team_name": "C", "members": []},
    }
    result = analyze_team_membership(members_by_team)
    assert result["member_counts"] == {"A": 2, "B": 1, "C": 0}
    assert result["empty_teams"] == 1
    assert result["users_in_multiple_teams"] == 1
    assert result["total_unique_users"] == 2


def test_analyze_onedrive_usage_demo_fixture() -> None:
    result = analyze_onedrive_usage(demo_onedrive_stats())
    summary = result["summary"]
    assert summary["total_users"] == 3
    assert summary["users_with_onedrive"] == 2
    assert summary["users_with_errors"] == 1
    assert summary["total_storage_used_gb"] == 92.5
    assert summary["users_over_75_percent"] == 1
    assert summary["users_by_storage_tier"] == {
        "0-10GB": 0,
        "10-50GB": 1,
        "50-100GB": 1,
        "100GB+": 0,
    }


def test_analyze_onedrive_usage_rejects_negative_storage() -> None:
    with pytest.raises(ValueError):
        analyze_onedrive_usage([{"user_id": "u1", "storage": {"used_gb": -1, "total_gb": 10}}])


def test_analyze_sign_in_logs_demo_fixture() -> None:
    result = analyze_sign_in_logs(demo_sign_in_logs())
    assert result["total_sign_ins"] == 4
    assert result["successful_sign_ins"] == 3
    assert result["failed_sign_ins"] == 1
    assert result["success_rate"] == 75.0
    assert result["error_distribution"] == [("50126", 1)]


def test_analyze_sign_in_logs_without_status_counts_total_only() -> None:
    result = analyze_sign_in_logs([{"user_display_name": "Ada Example"}])
    assert result["total_sign_ins"] == 1
    assert result["successful_sign_ins"] == 0
    assert result["failed_sign_ins"] == 0


def test_group_logs_by_user_drops_rows_without_user() -> None:
    logs = demo_sign_in_logs() + [{"id": "orphan"}]
    grouped = group_logs_by_user(logs)
    assert set(grouped) == {
        "11111111-1111-1111-1111-111111111111",
        "22222222-2222-2222-2222-222222222222",
    }
    assert len(grouped["11111111-1111-1111-1111-111111111111"]) == 2


def test_demo_memberships_are_consistent() -> None:
    # Every membership row names a demo user and a demo group.
    from m365_graph_toolkit.demo_data import demo_groups, demo_users

    user_ids = {user["id"] for user in demo_users()}
    group_ids = {group["id"] for group in demo_groups()}
    for row in demo_memberships():
        assert row["user_id"] in user_ids
        assert row["group_id"] in group_ids
