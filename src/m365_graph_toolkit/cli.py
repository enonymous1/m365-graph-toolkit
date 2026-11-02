"""Command-line interface: read-only exports, audit logs, STIG preview, demo.

Live commands need an Entra app registration. Set these variables
(``.env`` files are honoured — the CLI loads one from the working
directory if ``python-dotenv`` is installed):

* ``MS_TENANT_ID``
* ``MS_CLIENT_ID``
* ``MS_CLIENT_SECRET``

The CLI never prints a secret or a token. ``check-auth`` prints only
the token's expiry. Every live command is read-only except ``sync``,
which writes to Google Sheets and **clears the target worksheet
first** — treat that worksheet as owned by the sync, and try
``--dry-run`` (which fetches from Graph but never touches Sheets)
before a first real sync. There is no CLI path to create, update, or
delete users, groups, sites, files, plans, or lists. Those
operations exist in the library for code that needs them, behind its
own confirmation discipline. ``stig preview`` reads local CSV files
only and needs no credentials; the STIG import itself is library-only
(see ``examples/stig_import.py``).

``demo`` needs no credentials and touches no tenant. It writes the
fictional data from :mod:`m365_graph_toolkit.demo_data` so the export
and analysis paths can be exercised end to end offline.
"""

from __future__ import annotations

import argparse
import asyncio
import os
import sys
from pathlib import Path
from typing import Any, TypeVar

from . import __version__
from .analysis import analyze_onedrive_usage, analyze_sign_in_logs, analyze_team_membership
from .auth import AuthManager
from .clients import (
    AuditLogsClient,
    BaseClient,
    CalendarClient,
    GroupsClient,
    MailClient,
    OneDriveClient,
    PlannerClient,
    SharePointClient,
    TeamsMessagesClient,
    UsersClient,
)
from .clients.calendar import EVENT_FIELDS
from .clients.mail import MESSAGE_FIELDS
from .clients.planner import PLAN_FIELDS, TASK_FIELDS
from .clients.users import USER_FIELDS
from .demo_data import (
    demo_channel_messages,
    demo_events,
    demo_groups,
    demo_mail_messages,
    demo_memberships,
    demo_onedrive_stats,
    demo_planner_tasks,
    demo_plans,
    demo_sign_in_logs,
    demo_sites,
    demo_stig_plan,
    demo_users,
)
from .errors import M365Error, ValidationError
from .exporters import write_csv, write_json
from .sheets import SheetsExporter, sanitize_worksheet_title
from .stig import plan_lists_from_directory

_PLAN_EXPORT_FIELDS = [
    *PLAN_FIELDS,
    "created_by_user_id",
    "container_id",
    "container_type",
    "container_url",
    "group_name",
]
_TASK_EXPORT_FIELDS = [
    *TASK_FIELDS,
    "plan_id",
    "bucket_name",
    "created_by_user_id",
    "completed_by_user_id",
    "assigned_to",
    "labels",
    "description",
    "checklist_items",
    "checklist_total",
    "checklist_completed",
]
_EVENT_EXPORT_FIELDS = [
    *EVENT_FIELDS,
    "start_date_time",
    "start_time_zone",
    "end_date_time",
    "end_time_zone",
    "location",
    "organizer_email",
    "attendee_emails",
]
_MAIL_EXPORT_FIELDS = [
    *MESSAGE_FIELDS,
    "from_email",
    "from_name",
    "to_emails",
    "cc_emails",
    "bcc_emails",
    "body",
    "body_content_type",
]
_CHANNEL_MESSAGE_EXPORT_FIELDS = [
    "id",
    "subject",
    "importance",
    "message_type",
    "created_date_time",
    "last_modified_date_time",
    "deleted_date_time",
    "body",
    "body_content_type",
    "sender_type",
    "sender_id",
    "sender_display_name",
    "reply_count",
    "reaction_count",
    "reaction_types",
    "team_id",
    "channel_id",
]


def _load_dotenv() -> None:
    try:
        from dotenv import load_dotenv
    except ImportError:  # pragma: no cover - optional dependency
        return
    load_dotenv()


_ClientT = TypeVar("_ClientT", bound=BaseClient)


def _make_client(client_cls: type[_ClientT]) -> _ClientT:
    """Build one live client from the environment (single AuthManager path).

    The nine ``_make_*_client`` wrappers below exist so tests can
    monkeypatch one client type at a time; they are one-liners over
    this helper instead of nine copies of the same construction.
    """
    return client_cls.from_auth(AuthManager())


def _make_users_client() -> UsersClient:
    return _make_client(UsersClient)


def _make_groups_client() -> GroupsClient:
    return _make_client(GroupsClient)


def _make_sharepoint_client() -> SharePointClient:
    return _make_client(SharePointClient)


def _make_onedrive_client() -> OneDriveClient:
    return _make_client(OneDriveClient)


def _make_audit_client() -> AuditLogsClient:
    return _make_client(AuditLogsClient)


def _make_planner_client() -> PlannerClient:
    return _make_client(PlannerClient)


def _make_calendar_client() -> CalendarClient:
    return _make_client(CalendarClient)


def _make_mail_client() -> MailClient:
    return _make_client(MailClient)


def _make_teams_messages_client() -> TeamsMessagesClient:
    return _make_client(TeamsMessagesClient)


def _make_sheets_exporter(
    service_account_key_json_or_path: str, spreadsheet_id: str
) -> SheetsExporter:
    return SheetsExporter(service_account_key_json_or_path, spreadsheet_id)


def _resolve_sheets_config(args: argparse.Namespace) -> tuple[str, str]:
    """Sheets credentials for a real (non-dry-run) sync.

    CLI flags win over the environment. The original project's
    variable names are kept: ``GCS_SERVICE_ACCOUNT_KEY_JSON`` (a JSON
    string or a path to the key file) and ``TARGET_GOOGLE_SHEET_ID``.
    A missing value is an error that exits non-zero — the old CLI
    logged it and exited successfully, so a misconfigured scheduled
    run looked green.
    """
    service_account = args.service_account or os.environ.get("GCS_SERVICE_ACCOUNT_KEY_JSON", "")
    spreadsheet_id = args.spreadsheet or os.environ.get("TARGET_GOOGLE_SHEET_ID", "")
    missing = [
        name
        for name, value in (
            ("GCS_SERVICE_ACCOUNT_KEY_JSON", service_account),
            ("TARGET_GOOGLE_SHEET_ID", spreadsheet_id),
        )
        if not value
    ]
    if missing:
        raise ValidationError(
            "Missing required Google Sheets configuration: "
            + ", ".join(missing)
            + ". Pass --service-account/--spreadsheet or set the variables."
        )
    return service_account, spreadsheet_id


def _write_both(rows: list[dict[str, Any]], out_dir: Path, stem: str) -> list[Path]:
    return [write_csv(rows, out_dir, stem), write_json(rows, out_dir, stem)]


def _write_rows(
    rows: list[dict[str, Any]], out_dir: Path, stem: str, *, fieldnames: list[str]
) -> list[Path]:
    """Write CSV + JSON, including a header-only CSV for an empty result."""
    if rows:
        return _write_both(rows, out_dir, stem)
    return [write_csv(rows, out_dir, stem, fieldnames=fieldnames), write_json(rows, out_dir, stem)]


async def _export_users(out_dir: Path, *, top: int | None) -> list[Path]:
    client = _make_users_client()
    users = await client.get_users(top=top)
    paths = _write_both(users, out_dir, "users")
    print(f"Exported {len(users)} users")
    return paths


async def _export_memberships(out_dir: Path) -> list[Path]:
    client = _make_users_client()
    users = await client.get_users()
    memberships = await client.get_user_memberships_batch(users)
    paths = _write_both(memberships, out_dir, "user_memberships")
    print(f"Exported {len(memberships)} membership rows for {len(users)} users")
    return paths


async def _export_teams(out_dir: Path) -> list[Path]:
    client = _make_groups_client()
    teams = await client.get_teams()
    members_by_team: dict[str, dict[str, Any]] = {}
    for team in teams:
        members = await client.get_team_members(str(team["id"]))
        members_by_team[str(team["id"])] = {
            "team_name": team.get("display_name"),
            "members": members,
        }
    analysis = analyze_team_membership(members_by_team)
    paths = _write_both(teams, out_dir, "teams")
    paths.append(
        write_json({"teams": members_by_team, "analysis": analysis}, out_dir, "teams_composition")
    )
    print(f"Exported {len(teams)} teams")
    return paths


async def _export_sites(out_dir: Path) -> list[Path]:
    client = _make_sharepoint_client()
    sites = await client.get_all_sites()
    paths = _write_both(sites, out_dir, "sites")
    print(f"Exported {len(sites)} sites")
    return paths


async def _export_onedrive(out_dir: Path) -> list[Path]:
    users_client = _make_users_client()
    onedrive = _make_onedrive_client()
    users = await users_client.get_users()
    stats: list[dict[str, Any]] = []
    for user in users:
        user_id = str(user["id"])
        try:
            client_for_user = onedrive.for_user(user_id)
            info = await client_for_user.get_drive_info()
            quota = info.get("quota") or {}
            counts = await client_for_user.count_items()
            stats.append(
                {
                    "user_id": user_id,
                    "display_name": user.get("display_name"),
                    "storage": {
                        "used_gb": round(float(quota.get("used", 0)) / (1024**3), 2),
                        "total_gb": round(float(quota.get("total", 0)) / (1024**3), 2),
                        "remaining_gb": round(float(quota.get("remaining", 0)) / (1024**3), 2),
                        "state": quota.get("state", ""),
                    },
                    "items": counts,
                }
            )
        except Exception as error:  # noqa: BLE001 - per-user error entry, run continues
            stats.append(
                {
                    "user_id": user_id,
                    "display_name": user.get("display_name"),
                    "error": str(error),
                }
            )
    analysis = analyze_onedrive_usage(stats)
    paths = _write_both(stats, out_dir, "onedrive_usage")
    paths.append(write_json(analysis, out_dir, "onedrive_analysis"))
    print(f"Exported OneDrive usage for {len(stats)} users")
    return paths


async def _export_sign_in_logs(out_dir: Path, *, days: int, user_id: str | None) -> list[Path]:
    client = _make_audit_client()
    logs = await client.get_audit_logs(log_type="signIn", days=days, user_id=user_id)
    analysis = analyze_sign_in_logs(logs)
    paths = _write_both(logs, out_dir, "sign_in_logs")
    paths.append(write_json(analysis, out_dir, "sign_in_analysis"))
    print(f"Exported {len(logs)} sign-in log entries")
    return paths


async def _export_planner(out_dir: Path) -> list[Path]:
    groups_client = _make_groups_client()
    planner = _make_planner_client()
    groups = await groups_client.get_groups()
    plans = await planner.get_all_plans(groups)
    tasks: list[dict[str, Any]] = []
    for plan in plans:
        tasks.extend(await planner.get_plan_tasks(str(plan["id"])))
    paths = _write_rows(plans, out_dir, "planner_plans", fieldnames=_PLAN_EXPORT_FIELDS)
    paths += _write_rows(tasks, out_dir, "planner_tasks", fieldnames=_TASK_EXPORT_FIELDS)
    print(f"Exported {len(plans)} plans and {len(tasks)} tasks")
    return paths


async def _export_calendar_events(out_dir: Path, *, user_id: str) -> list[Path]:
    client = _make_calendar_client()
    events = await client.get_user_events(user_id)
    paths = _write_rows(events, out_dir, "calendar_events", fieldnames=_EVENT_EXPORT_FIELDS)
    print(f"Exported {len(events)} calendar events for {user_id}")
    return paths


async def _export_mail(out_dir: Path, *, user_id: str) -> list[Path]:
    client = _make_mail_client()
    messages = await client.get_user_messages(user_id)
    paths = _write_rows(messages, out_dir, "mail", fieldnames=_MAIL_EXPORT_FIELDS)
    print(f"Exported {len(messages)} messages for {user_id}")
    return paths


async def _export_channel_messages(out_dir: Path, *, team_id: str, channel_id: str) -> list[Path]:
    client = _make_teams_messages_client()
    messages = await client.get_channel_messages(team_id, channel_id)
    paths = _write_rows(
        messages, out_dir, "channel_messages", fieldnames=_CHANNEL_MESSAGE_EXPORT_FIELDS
    )
    print(f"Exported {len(messages)} channel messages for {channel_id}")
    return paths


async def _sync_users(
    *,
    service_account: str | None,
    spreadsheet_id: str | None,
    worksheet: str,
    dry_run: bool,
    top: int | None,
) -> None:
    """Fetch users from Graph and replace a Google Sheets worksheet.

    The Sheets exporter is built only after a successful fetch and
    only for a real sync: a dry run needs no Google credentials at
    all, fixing the old CLI, which authenticated to Sheets before
    honouring ``--dry-run``.
    """
    client = _make_users_client()
    users = await client.get_users(top=top)
    if dry_run:
        print(f"Dry-run: fetched {len(users)} users; skipping Google Sheets write")
        return
    if service_account is None or spreadsheet_id is None:
        # An assert would vanish under ``python -O`` and surface as a
        # confusing TypeError inside the exporter; fail explicitly.
        raise ValidationError("Google Sheets configuration is required for a real sync")
    exporter = _make_sheets_exporter(service_account, spreadsheet_id)
    title = sanitize_worksheet_title(worksheet)
    written = exporter.write_rows(users, title, fieldnames=list(USER_FIELDS))
    print(f"Synced {written} users to worksheet {title!r}")


async def _sync_mail(
    *,
    user_id: str,
    service_account: str | None,
    spreadsheet_id: str | None,
    worksheet: str | None,
    dry_run: bool,
    top: int | None,
) -> None:
    """Fetch one user's messages and replace a Google Sheets worksheet."""
    client = _make_mail_client()
    messages = await client.get_user_messages(user_id, top=top)
    if dry_run:
        print(
            f"Dry-run: fetched {len(messages)} messages for {user_id}; skipping Google Sheets write"
        )
        return
    if service_account is None or spreadsheet_id is None:
        # An assert would vanish under ``python -O`` and surface as a
        # confusing TypeError inside the exporter; fail explicitly.
        raise ValidationError("Google Sheets configuration is required for a real sync")
    exporter = _make_sheets_exporter(service_account, spreadsheet_id)
    title = sanitize_worksheet_title(worksheet or f"mail-{user_id}")
    written = exporter.write_rows(messages, title, fieldnames=_MAIL_EXPORT_FIELDS)
    print(f"Synced {written} messages for {user_id} to worksheet {title!r}")


def _run_stig_preview(directory: Path, out_dir: Path | None) -> list[Path]:
    """Plan SharePoint lists for local STIG CSVs. No credentials, no writes to a tenant."""
    plans = plan_lists_from_directory(directory)
    for plan in plans:
        print(
            f"{plan['list_name']}: {plan['row_count']} rows, "
            f"{len(plan['columns'])} columns"
            + (
                f" (skipped: {', '.join(plan['skipped_headers'])})"
                if plan["skipped_headers"]
                else ""
            )
        )
    if out_dir is None:
        return []
    path = write_json(plans, out_dir, "stig_preview")
    print(f"Wrote preview for {len(plans)} lists to {path}")
    return [path]


def _run_demo(out_dir: Path) -> list[Path]:
    """Write fictional exports + analyses. No credentials, no network."""
    users = demo_users()
    memberships = demo_memberships()
    sites = demo_sites()
    groups = demo_groups()
    onedrive = demo_onedrive_stats()
    logs = demo_sign_in_logs()
    members_by_team = {
        str(group["id"]): {
            "team_name": group["display_name"],
            "members": [
                {"id": row["user_id"]} for row in memberships if row["group_id"] == group["id"]
            ],
        }
        for group in groups
    }
    paths: list[Path] = []
    paths += _write_both(users, out_dir, "users_demo")
    paths += _write_both(memberships, out_dir, "user_memberships_demo")
    paths += _write_both(sites, out_dir, "sites_demo")
    paths += _write_both(demo_plans(), out_dir, "planner_plans_demo")
    paths += _write_both(demo_planner_tasks(), out_dir, "planner_tasks_demo")
    paths += _write_both(demo_events(), out_dir, "calendar_events_demo")
    paths += _write_both(demo_mail_messages(), out_dir, "mail_demo")
    paths += _write_both(demo_channel_messages(), out_dir, "channel_messages_demo")
    paths.append(write_json([demo_stig_plan()], out_dir, "stig_preview_demo"))
    paths.append(write_json(demo_onedrive_stats(), out_dir, "onedrive_usage_demo"))
    paths.append(write_json(analyze_onedrive_usage(onedrive), out_dir, "onedrive_analysis_demo"))
    paths += _write_both(logs, out_dir, "sign_in_logs_demo")
    paths.append(write_json(analyze_sign_in_logs(logs), out_dir, "sign_in_analysis_demo"))
    paths.append(
        write_json(
            {"teams": members_by_team, "analysis": analyze_team_membership(members_by_team)},
            out_dir,
            "teams_composition_demo",
        )
    )
    report = out_dir / "DEMO_README.txt"
    report.write_text(
        "This output was produced by `m365-toolkit demo` from fictional\n"
        "data bundled with the package. It is not from a real tenant.\n",
        encoding="utf-8",
    )
    paths.append(report)
    print(f"Wrote {len(paths)} demo files (fictional data) to {out_dir}")
    return paths


def _check_auth() -> None:
    auth = AuthManager()
    token = auth.get_access_token()
    expires_on = getattr(token, "expires_on", None)
    print(f"Authentication OK. Token expires_on={expires_on} (token value not shown).")


def build_parser() -> argparse.ArgumentParser:
    """Build the argument parser for every subcommand (export/sync/logs/demo/stig)."""
    parser = argparse.ArgumentParser(
        prog="m365-toolkit",
        description="Export and analyse Microsoft 365 data via the Graph API.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    subparsers = parser.add_subparsers(dest="command", required=True)

    export = subparsers.add_parser("export", help="Export data to CSV + JSON.")
    export_sub = export.add_subparsers(dest="export_target", required=True)
    users_parser = export_sub.add_parser("users", help="Export all users.")
    users_parser.add_argument("--out", type=Path, required=True)
    users_parser.add_argument("--top", type=int, default=None, help="Page size (not a total cap).")
    memberships_parser = export_sub.add_parser("memberships", help="Export user memberships.")
    memberships_parser.add_argument("--out", type=Path, required=True)
    teams_parser = export_sub.add_parser("teams", help="Export teams and their composition.")
    teams_parser.add_argument("--out", type=Path, required=True)
    sites_parser = export_sub.add_parser("sites", help="Export SharePoint sites.")
    sites_parser.add_argument("--out", type=Path, required=True)
    onedrive_parser = export_sub.add_parser("onedrive", help="Export OneDrive usage per user.")
    onedrive_parser.add_argument("--out", type=Path, required=True)
    planner_parser = export_sub.add_parser("planner", help="Export Planner plans and their tasks.")
    planner_parser.add_argument("--out", type=Path, required=True)
    calendar_parser = export_sub.add_parser(
        "calendar-events", help="Export one user's calendar events."
    )
    calendar_parser.add_argument("--out", type=Path, required=True)
    calendar_parser.add_argument("--user", required=True, help="User id or UPN.")
    mail_parser = export_sub.add_parser("mail", help="Export one user's Outlook messages.")
    mail_parser.add_argument("--out", type=Path, required=True)
    mail_parser.add_argument("--user", required=True, help="User id or UPN.")
    channel_parser = export_sub.add_parser(
        "channel-messages", help="Export one Teams channel's messages."
    )
    channel_parser.add_argument("--out", type=Path, required=True)
    channel_parser.add_argument("--team", required=True, help="Team id.")
    channel_parser.add_argument("--channel", required=True, help="Channel id.")

    sync = subparsers.add_parser(
        "sync",
        help="Sync data to Google Sheets (replaces worksheet content).",
    )
    sync_sub = sync.add_subparsers(dest="sync_target", required=True)
    sync_users = sync_sub.add_parser("users", help="Sync all users to a worksheet.")
    sync_users.add_argument(
        "--spreadsheet",
        default=None,
        help="Spreadsheet id (default: $TARGET_GOOGLE_SHEET_ID).",
    )
    sync_users.add_argument(
        "--service-account",
        default=None,
        help="Service-account key: JSON string or path to a key file "
        "(default: $GCS_SERVICE_ACCOUNT_KEY_JSON).",
    )
    sync_users.add_argument("--worksheet", default="users", help="Worksheet title.")
    sync_users.add_argument("--dry-run", action="store_true", help="Fetch but do not write.")
    sync_users.add_argument("--top", type=int, default=None, help="Page size (not a total cap).")
    sync_mail = sync_sub.add_parser("mail", help="Sync one user's Outlook messages to a worksheet.")
    sync_mail.add_argument(
        "--spreadsheet",
        default=None,
        help="Spreadsheet id (default: $TARGET_GOOGLE_SHEET_ID).",
    )
    sync_mail.add_argument(
        "--service-account",
        default=None,
        help="Service-account key: JSON string or path to a key file "
        "(default: $GCS_SERVICE_ACCOUNT_KEY_JSON).",
    )
    sync_mail.add_argument("--user", required=True, help="User id or UPN.")
    sync_mail.add_argument(
        "--worksheet",
        default=None,
        help="Worksheet title (default: mail-<user>).",
    )
    sync_mail.add_argument("--dry-run", action="store_true", help="Fetch but do not write.")
    sync_mail.add_argument("--top", type=int, default=None, help="Page size (not a total cap).")

    logs = subparsers.add_parser("logs", help="Export audit logs.")
    logs_sub = logs.add_subparsers(dest="logs_target", required=True)
    sign_ins = logs_sub.add_parser("sign-ins", help="Export sign-in logs.")
    sign_ins.add_argument("--out", type=Path, required=True)
    sign_ins.add_argument("--days", type=int, default=7)
    sign_ins.add_argument("--user", type=str, default=None, help="User id to filter by.")

    demo = subparsers.add_parser("demo", help="Write fictional demo exports (no credentials).")
    demo.add_argument("--out", type=Path, required=True)

    stig = subparsers.add_parser(
        "stig", help="Plan SharePoint lists from local STIG CSVs (no tenant writes)."
    )
    stig_sub = stig.add_subparsers(dest="stig_target", required=True)
    preview = stig_sub.add_parser(
        "preview", help="Show the lists that would be created from a CSV directory."
    )
    preview.add_argument("--dir", type=Path, required=True, help="Directory of STIG CSVs.")
    preview.add_argument("--out", type=Path, default=None, help="Write the preview JSON here.")

    subparsers.add_parser("check-auth", help="Verify credentials by acquiring a token.")
    return parser


async def _dispatch(args: argparse.Namespace) -> None:
    if args.command == "export":
        if args.export_target == "users":
            await _export_users(args.out, top=args.top)
        elif args.export_target == "memberships":
            await _export_memberships(args.out)
        elif args.export_target == "teams":
            await _export_teams(args.out)
        elif args.export_target == "sites":
            await _export_sites(args.out)
        elif args.export_target == "onedrive":
            await _export_onedrive(args.out)
        elif args.export_target == "planner":
            await _export_planner(args.out)
        elif args.export_target == "calendar-events":
            await _export_calendar_events(args.out, user_id=args.user)
        elif args.export_target == "mail":
            await _export_mail(args.out, user_id=args.user)
        elif args.export_target == "channel-messages":
            await _export_channel_messages(args.out, team_id=args.team, channel_id=args.channel)
    elif args.command == "sync":
        # Resolve Sheets config before any Graph fetch for a real
        # sync, so a misconfigured run fails fast. A dry run needs
        # no Sheets config at all.
        if args.dry_run:
            service_account, spreadsheet_id = None, None
        else:
            service_account, spreadsheet_id = _resolve_sheets_config(args)
        if args.sync_target == "users":
            await _sync_users(
                service_account=service_account,
                spreadsheet_id=spreadsheet_id,
                worksheet=args.worksheet,
                dry_run=args.dry_run,
                top=args.top,
            )
        elif args.sync_target == "mail":
            await _sync_mail(
                user_id=args.user,
                service_account=service_account,
                spreadsheet_id=spreadsheet_id,
                worksheet=args.worksheet,
                dry_run=args.dry_run,
                top=args.top,
            )
    elif args.command == "logs" and args.logs_target == "sign-ins":
        await _export_sign_in_logs(args.out, days=args.days, user_id=args.user)
    elif args.command == "stig" and args.stig_target == "preview":
        _run_stig_preview(args.dir, args.out)
    elif args.command == "demo":
        _run_demo(args.out)
    elif args.command == "check-auth":
        _check_auth()


def main(argv: list[str] | None = None) -> int:
    """CLI entry point. Returns a process exit code."""
    _load_dotenv()
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        asyncio.run(_dispatch(args))
    except M365Error as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1
    except OSError as error:
        # File writes (demo/export output) are the one non-library
        # failure path; a traceback here would hide the path that
        # failed behind a stack trace, so report it like any error.
        print(f"Error: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
