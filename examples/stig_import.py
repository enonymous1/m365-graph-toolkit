"""Preview, then import, STIG CSVs as SharePoint lists. Library only.

This writes to a tenant: it creates one SharePoint list per CSV and
imports every row. There is no CLI path to it. Preview first with::

    m365-toolkit stig preview --dir <stig-csv-directory>

Then, only after checking the preview and confirming the site id,
run this script with the real site id and directory::

    python examples/stig_import.py --site <site-id> --dir <stig-csv-directory>

The site id is the Graph site id (``sites/root`` or a site from
``SharePointClient.search_sites``), not a display name. STIG CSVs
are not bundled with this repository; download the current export
from https://public.cyber.mil/stigs/ first. ``examples/Fictional Example STIG.csv``
is fictional data for preview testing only — do not import it as if
it were a real STIG.
"""

from __future__ import annotations

import argparse
import asyncio
import json
from pathlib import Path

from dotenv import load_dotenv

from m365_graph_toolkit.auth import AuthManager
from m365_graph_toolkit.clients import SharePointClient
from m365_graph_toolkit.stig import plan_lists_from_directory


async def main(site_id: str, directory: Path) -> None:
    plans = plan_lists_from_directory(directory)
    print(f"Planning {len(plans)} lists from {directory}:")
    for plan in plans:
        print(f"  {plan['list_name']}: {plan['row_count']} rows, {len(plan['columns'])} columns")
    answer = input("Import these lists into the site? Type YES to continue: ").strip()
    if answer != "YES":
        print("Aborted; nothing was written.")
        return
    client = SharePointClient.from_auth(AuthManager())  # type: ignore[arg-type]
    summaries = await client.import_lists_from_directory(site_id, directory)
    print(json.dumps(summaries, indent=2, default=str))
    failed = [s for s in summaries if s.get("error") or s["items_failed"]]
    if failed:
        print(f"{len(failed)} list(s) had errors or failed rows; see the summary above.")


if __name__ == "__main__":
    load_dotenv()
    parser = argparse.ArgumentParser()
    parser.add_argument("--site", required=True, help="Graph site id.")
    parser.add_argument("--dir", type=Path, required=True, help="Directory of STIG CSVs.")
    args = parser.parse_args()
    asyncio.run(main(args.site, args.dir))
