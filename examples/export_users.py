"""Export all users and their group memberships to CSV + JSON.

Run from the project root with credentials in the environment (or a
``.env`` file)::

    python examples/export_users.py --out ./exports
"""

from __future__ import annotations

import argparse
import asyncio
from pathlib import Path

from dotenv import load_dotenv

from m365_graph_toolkit.auth import AuthManager
from m365_graph_toolkit.clients import UsersClient
from m365_graph_toolkit.exporters import write_csv, write_json


async def main(out_dir: Path) -> None:
    client = UsersClient.from_auth(AuthManager())  # type: ignore[arg-type]
    users = await client.get_users()
    memberships = await client.get_user_memberships_batch(users)
    print(write_csv(users, out_dir, "users"))
    print(write_json(users, out_dir, "users"))
    print(write_csv(memberships, out_dir, "user_memberships"))
    print(f"Exported {len(users)} users, {len(memberships)} membership rows")


if __name__ == "__main__":
    load_dotenv()
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()
    asyncio.run(main(args.out))
