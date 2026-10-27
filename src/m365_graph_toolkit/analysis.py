"""Pure analysis functions over already-fetched dict rows.

Everything here is a pure function of its inputs: no network, no
files, no global state. That is what makes them unit-testable with
fictional fixtures, and what v1's analysers were not — they mixed
fetching, file discovery, and statistics in single methods that
assumed specific CSV column names (``user_id`` in one place,
``displayName`` in another) and crashed when the exporter's naming
didn't match.

Input contract: rows use the ``snake_case`` field names produced by
the clients (``display_name``, ``user_principal_name``,
``created_date_time``, ...). Rows missing a required field are
skipped for that statistic and counted in ``*_skipped`` where the
result carries one; a row is never silently read as zero.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from typing import Any


def is_internal_user(user: dict[str, Any], domain: str) -> bool:
    """Whether a user's mail or UPN belongs to ``domain``.

    ``domain`` is an explicit argument: v1 hardcoded one employer's
    domain and then ignored the UPN entirely (its docstring promised
    a mail/UPN match and its ``all([...])`` checked mail only). A
    caller that passes no domain has not said what "internal" means.
    """
    if not user or not domain:
        return False
    domain = domain.strip().lower().lstrip("@")
    # Clients normalise to snake_case, but callers may hand this rows
    # straight from Graph JSON (camelCase) — accept both spellings.
    for key in ("mail", "user_principal_name", "userPrincipalName"):
        value = user.get(key)
        if isinstance(value, str) and value.strip().lower().endswith("@" + domain):
            return True
    return False


def analyze_team_membership(members_by_team: dict[str, dict[str, Any]]) -> dict[str, Any]:
    """Statistics over ``{team_id: {"team_name": ..., "members": [...]}}``.

    Membership rows are dicts with an ``id`` key. A team with no
    members counts as empty; a user counted twice in one team counts
    once for that team.
    """
    member_counts: dict[str, int] = {}
    user_teams: dict[str, set[str]] = defaultdict(set)
    for team_id, data in members_by_team.items():
        members = data.get("members") or []
        unique_ids = {str(member.get("id")) for member in members if member.get("id")}
        member_counts[data.get("team_name") or team_id] = len(unique_ids)
        for user_id in unique_ids:
            user_teams[user_id].add(team_id)
    counts = list(member_counts.values())
    return {
        "total_teams": len(members_by_team),
        "teams_with_members": sum(1 for count in counts if count > 0),
        "empty_teams": sum(1 for count in counts if count == 0),
        "member_counts": member_counts,
        "users_in_multiple_teams": sum(1 for teams in user_teams.values() if len(teams) > 1),
        "total_unique_users": len(user_teams),
        "avg_members_per_team": (sum(counts) / len(counts)) if counts else 0.0,
        "max_members_per_team": max(counts) if counts else 0,
    }


def analyze_onedrive_usage(stats: list[dict[str, Any]]) -> dict[str, Any]:
    """Summarise per-user OneDrive stats (``storage``/``items`` shape).

    Each stat is ``{"user_id": ..., "storage": {"used_gb", "total_gb",
    "remaining_gb", "state"}, "items": {...}}`` or an ``{"error": ...}``
    entry for a user whose drive could not be read. Negative storage
    values are a fail-loud error: a derived field that violates its
    domain must not be written into a report as if it were measured.
    """
    summary: dict[str, Any] = {
        "total_users": len(stats),
        "users_with_onedrive": 0,
        "users_with_errors": 0,
        "total_storage_used_gb": 0.0,
        "average_usage_gb": 0.0,
        "users_over_75_percent": 0,
        "users_by_storage_tier": {"0-10GB": 0, "10-50GB": 0, "50-100GB": 0, "100GB+": 0},
    }
    used_values: list[float] = []
    for stat in stats:
        storage = stat.get("storage")
        if not isinstance(storage, dict):
            summary["users_with_errors"] += 1
            continue
        used = storage.get("used_gb")
        total = storage.get("total_gb")
        if used is None or total is None:
            summary["users_with_errors"] += 1
            continue
        if float(used) < 0 or float(total) < 0:
            raise ValueError(f"Negative storage value for user {stat.get('user_id')}")
        summary["users_with_onedrive"] += 1
        used_values.append(float(used))
        if float(used) <= 10:
            summary["users_by_storage_tier"]["0-10GB"] += 1
        elif float(used) <= 50:
            summary["users_by_storage_tier"]["10-50GB"] += 1
        elif float(used) <= 100:
            summary["users_by_storage_tier"]["50-100GB"] += 1
        else:
            summary["users_by_storage_tier"]["100GB+"] += 1
        if float(total) > 0 and (float(used) / float(total)) > 0.75:
            summary["users_over_75_percent"] += 1
    if used_values:
        summary["total_storage_used_gb"] = round(sum(used_values), 2)
        summary["average_usage_gb"] = round(sum(used_values) / len(used_values), 2)
    return {"summary": summary, "details": stats}


def analyze_sign_in_logs(logs: list[dict[str, Any]]) -> dict[str, Any]:
    """Summary statistics for sign-in log rows (client field names)."""
    if not logs:
        return {"total_sign_ins": 0, "message": "No logs to analyze"}
    successful = 0
    failed = 0
    users: Counter[str] = Counter()
    apps: Counter[str] = Counter()
    error_codes: Counter[str] = Counter()
    ip_addresses: Counter[str] = Counter()
    risk_levels: Counter[str] = Counter()
    for log in logs:
        status = log.get("status")
        error_code = status.get("error_code") if isinstance(status, dict) else log.get("error_code")
        if error_code is None:
            # Rows without a status are neither success nor failure;
            # they are counted in the total only.
            pass
        elif error_code == 0 or error_code == "0":
            successful += 1
        else:
            failed += 1
            error_codes[str(error_code)] += 1
        users[str(log.get("user_display_name") or "Unknown")] += 1
        apps[str(log.get("app_display_name") or "Unknown")] += 1
        ip_addresses[str(log.get("ip_address") or "Unknown")] += 1
        risk_levels[str(log.get("risk_level_aggregated") or "Unknown")] += 1
    total = len(logs)
    return {
        "total_sign_ins": total,
        "successful_sign_ins": successful,
        "failed_sign_ins": failed,
        "success_rate": round((successful / total) * 100, 2) if total else 0.0,
        "top_users": users.most_common(10),
        "top_applications": apps.most_common(10),
        "top_ip_addresses": ip_addresses.most_common(10),
        "error_distribution": error_codes.most_common(),
        "risk_levels": risk_levels.most_common(),
    }


def group_logs_by_user(logs: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    """Group log rows by ``user_id``; rows without one are dropped."""
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for log in logs:
        user_id = log.get("user_id")
        if user_id:
            grouped[str(user_id)].append(log)
    return dict(grouped)


__all__ = [
    "analyze_onedrive_usage",
    "analyze_sign_in_logs",
    "analyze_team_membership",
    "group_logs_by_user",
    "is_internal_user",
]
