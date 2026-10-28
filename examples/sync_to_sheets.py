"""Sync users and one user's mail to Google Sheets worksheets.

Run from the project root with credentials in the environment (or a
``.env`` file)::

    python examples/sync_to_sheets.py --user <user-id> --dry-run
    python examples/sync_to_sheets.py --user <user-id>

The sync **replaces** the target worksheets' contents (they are
cleared first), so run ``--dry-run`` first and point the first real
run at a test spreadsheet. Needs the ``[sheets]`` extra and a Google
service account with Editor access to the spreadsheet. Nothing here
writes to a Microsoft 365 tenant.
"""

from __future__ import annotations

import argparse
import asyncio
import os

from dotenv import load_dotenv

from m365_graph_toolkit.auth import AuthManager
from m365_graph_toolkit.clients import MailClient, UsersClient
from m365_graph_toolkit.clients.mail import MESSAGE_FIELDS
from m365_graph_toolkit.clients.users import USER_FIELDS
from m365_graph_toolkit.sheets import SheetsExporter

_MAIL_FIELDS = [
    *MESSAGE_FIELDS,
    "from_email",
    "from_name",
    "to_emails",
    "cc_emails",
    "bcc_emails",
    "body",
    "body_content_type",
]


async def main(user_id: str, *, dry_run: bool) -> None:
    auth = AuthManager()
    users_client = UsersClient.from_auth(auth)  # type: ignore[arg-type]
    mail_client = MailClient.from_auth(auth)  # type: ignore[arg-type]

    users = await users_client.get_users()
    messages = await mail_client.get_user_messages(user_id)
    print(f"Fetched {len(users)} users and {len(messages)} messages for {user_id}")
    if dry_run:
        print("Dry-run: skipping Google Sheets write")
        return

    service_account = os.environ.get("GCS_SERVICE_ACCOUNT_KEY_JSON", "")
    spreadsheet_id = os.environ.get("TARGET_GOOGLE_SHEET_ID", "")
    if not service_account or not spreadsheet_id:
        raise SystemExit(
            "Set GCS_SERVICE_ACCOUNT_KEY_JSON and TARGET_GOOGLE_SHEET_ID "
            "for a real sync (see .env.example)."
        )
    exporter = SheetsExporter(service_account, spreadsheet_id)
    print(f"Wrote {exporter.write_rows(users, 'users', fieldnames=list(USER_FIELDS))} users")
    print(
        f"Wrote {exporter.write_rows(messages, f'mail-{user_id}', fieldnames=_MAIL_FIELDS)} "
        "messages"
    )


if __name__ == "__main__":
    load_dotenv()
    parser = argparse.ArgumentParser()
    parser.add_argument("--user", required=True, help="User id or UPN for mail.")
    parser.add_argument("--dry-run", action="store_true", help="Fetch but do not write.")
    args = parser.parse_args()
    asyncio.run(main(args.user, dry_run=args.dry_run))
