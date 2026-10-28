"""Export Planner plans/tasks, calendar events, and mail for one user.

Run from the project root with credentials in the environment (or a
``.env`` file)::

    python examples/export_planner_calendar_mail.py --out ./exports --user <user-id>

Planner plans are tenant-wide (via the user's groups); calendar
events and mail are for the ``--user`` given. Nothing here writes to
a tenant.
"""

from __future__ import annotations

import argparse
import asyncio
from pathlib import Path

from dotenv import load_dotenv

from m365_graph_toolkit.auth import AuthManager
from m365_graph_toolkit.clients import (
    CalendarClient,
    GroupsClient,
    MailClient,
    PlannerClient,
)
from m365_graph_toolkit.exporters import write_csv, write_json


async def main(out_dir: Path, user_id: str) -> None:
    auth = AuthManager()
    groups = GroupsClient.from_auth(auth)  # type: ignore[arg-type]
    planner = PlannerClient.from_auth(auth)  # type: ignore[arg-type]
    calendar = CalendarClient.from_auth(auth)  # type: ignore[arg-type]
    mail = MailClient.from_auth(auth)  # type: ignore[arg-type]

    all_groups = await groups.get_groups()
    plans = await planner.get_all_plans(all_groups)
    tasks: list[dict] = []
    for plan in plans:
        tasks.extend(await planner.get_plan_tasks(str(plan["id"])))
    events = await calendar.get_user_events(user_id)
    messages = await mail.get_user_messages(user_id)

    print(write_csv(plans, out_dir, "planner_plans"))
    print(write_csv(tasks, out_dir, "planner_tasks"))
    print(write_csv(events, out_dir, "calendar_events"))
    print(write_json(messages, out_dir, "mail"))
    print(
        f"Exported {len(plans)} plans, {len(tasks)} tasks, "
        f"{len(events)} events, {len(messages)} messages"
    )


if __name__ == "__main__":
    load_dotenv()
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--user", required=True, help="User id or UPN for calendar/mail.")
    args = parser.parse_args()
    asyncio.run(main(args.out, args.user))
