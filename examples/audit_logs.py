"""Fetch sign-in logs for the last 7 days and print a summary.

Run from the project root with credentials in the environment (or a
``.env`` file)::

    python examples/audit_logs.py --days 7
"""

from __future__ import annotations

import argparse
import asyncio
import json

from dotenv import load_dotenv

from m365_graph_toolkit.analysis import analyze_sign_in_logs
from m365_graph_toolkit.auth import AuthManager
from m365_graph_toolkit.clients import AuditLogsClient


async def main(days: int) -> None:
    client = AuditLogsClient.from_auth(AuthManager())  # type: ignore[arg-type]
    logs = await client.get_audit_logs(log_type="signIn", days=days)
    print(json.dumps(analyze_sign_in_logs(logs), indent=2, default=str))


if __name__ == "__main__":
    load_dotenv()
    parser = argparse.ArgumentParser()
    parser.add_argument("--days", type=int, default=7)
    args = parser.parse_args()
    asyncio.run(main(args.days))
