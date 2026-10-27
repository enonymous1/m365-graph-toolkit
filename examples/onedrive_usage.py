"""Summarise OneDrive usage across the tenant.

Run from the project root with credentials in the environment (or a
``.env`` file)::

    python examples/onedrive_usage.py
"""

from __future__ import annotations

import asyncio
import json

from dotenv import load_dotenv

from m365_graph_toolkit.analysis import analyze_onedrive_usage
from m365_graph_toolkit.auth import AuthManager
from m365_graph_toolkit.clients import OneDriveClient, UsersClient


async def main() -> None:
    auth = AuthManager()
    users = await UsersClient.from_auth(auth).get_users()  # type: ignore[arg-type]
    drive = OneDriveClient.from_auth(auth)  # type: ignore[arg-type]
    stats = []
    for user in users:
        try:
            info = await drive.for_user(str(user["id"])).get_drive_info()
            quota = info.get("quota") or {}
            stats.append(
                {
                    "user_id": user["id"],
                    "storage": {
                        "used_gb": round(float(quota.get("used", 0)) / (1024**3), 2),
                        "total_gb": round(float(quota.get("total", 0)) / (1024**3), 2),
                        "remaining_gb": round(float(quota.get("remaining", 0)) / (1024**3), 2),
                        "state": quota.get("state", ""),
                    },
                }
            )
        except Exception as error:  # noqa: BLE001 - per-user error entry
            stats.append({"user_id": user["id"], "error": str(error)})
    print(json.dumps(analyze_onedrive_usage(stats)["summary"], indent=2))


if __name__ == "__main__":
    load_dotenv()
    asyncio.run(main())
