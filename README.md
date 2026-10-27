# M365 Graph Toolkit

Export and analyse Microsoft 365 data — users, groups and teams, SharePoint sites and lists, OneDrive usage, and audit logs — through the Microsoft Graph API, from one installable Python package and CLI.

This is a rewrite of a July 2026 codebase that did not run: its clients imported a module that never existed, its pagination and raw-HTTP helpers were never defined, and its exporters mixed SDK objects with dicts. Version 2.0.0 keeps the scope and fixes the plumbing. See [CHANGELOG.md](CHANGELOG.md) for the full list.

> **No live tenant run yet.** The SDK builder chains were checked against the installed `msgraph-sdk` source and every offline behaviour is tested, but auth, pagination, `$search`, photo content, and upload/download against a real tenant still need a maintainer run. See [Live check](#live-check).

## Install

```bash
pip install -e ".[graph]"
```

The base install (no `[graph]`) is enough for the offline `demo` command, the exporters, and the analysers. Live commands need the Graph dependencies (`azure-identity`, `msgraph-sdk`).

Python 3.10+.

## Setup

Create an Entra ID app registration with **application** permissions for what you need (for example `User.Read.All`, `Group.Read.All`, `Sites.Read.All`, `AuditLog.Read.All`, `Files.Read.All`), grant admin consent, and create a client secret. Then:

```bash
cp .env.example .env
# fill in MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET
```

Or export the three variables directly. The CLI loads a `.env` file from the working directory when `python-dotenv` is installed. Nothing in this project logs or prints a secret or a token value.

Check the credentials work (prints only the token expiry):

```bash
m365-toolkit check-auth
```

## Use

Every export writes a timestamped CSV **and** JSON file into the directory you name. Nothing is deleted, overwritten, or written inside the package directory.

```bash
# All users
m365-toolkit export users --out ./exports

# Group memberships for every user
m365-toolkit export memberships --out ./exports

# Teams and their composition + analysis
m365-toolkit export teams --out ./exports

# SharePoint sites
m365-toolkit export sites --out ./exports

# OneDrive quota + item counts per user
m365-toolkit export onedrive --out ./exports

# Sign-in logs for the last 7 days (plus an analysis file)
m365-toolkit logs sign-ins --out ./exports --days 7
m365-toolkit logs sign-ins --out ./exports --days 30 --user <user-id>
```

Try it without credentials or a tenant — fictional data only:

```bash
m365-toolkit demo --out ./demo-output
```

## Library

```python
import asyncio
from m365_graph_toolkit.auth import AuthManager
from m365_graph_toolkit.clients import UsersClient, AuditLogsClient

async def main() -> None:
    users = UsersClient.from_auth(AuthManager())
    all_users = await users.get_users()
    memberships = await users.get_user_memberships_batch(all_users)

    logs = AuditLogsClient.from_auth(AuthManager())
    sign_ins = await logs.get_audit_logs(log_type="signIn", days=7)

asyncio.run(main())
```

Clients return plain `dict` rows with `snake_case` field names, so the same rows feed the exporters, the analysers, and your own code without shape conversions. Examples live in [`examples/`](examples/).

Write operations (`create_user`, `update_user`, `delete_user`, list-item and drive-item writes) exist in the library only — there is no CLI path to them. Treat deletes as permanent: confirm before calling.

## What the analysers expect

Analysis functions in `m365_graph_toolkit.analysis` are pure functions over the clients' rows:

- `analyze_team_membership(members_by_team)` — team sizes, empty teams, users in multiple teams.
- `analyze_onedrive_usage(stats)` — storage tiers, users over 75% of quota, error entries counted separately. Negative storage values raise instead of being reported as measurements.
- `analyze_sign_in_logs(logs)` — success/failure counts from `status.error_code`, top users/apps/IPs, error distribution.
- `is_internal_user(user, domain)` — mail **or** UPN ending in the domain you pass. The domain is always an explicit argument; nothing is hardcoded.

## Rate limits

List methods page through `odata_next_link` and retry 429/502/503/504 responses with exponential backoff (capped at 30s, with jitter), honouring `Retry-After` when Graph sends one. Other failures (400/401/403/404) fail immediately — a permission problem should never look like a slow tenant. A run that retried is logged as such; do not present a throttled collection as a complete measurement.

OneDrive simple upload supports files under 4 MB. Larger files raise an error: Graph requires a resumable upload session for them, which this client does not implement.

`GET /sites` without search returns only a subset of sites in some tenants (followed/recently accessed). For a full tenant inventory, use search-based discovery against your tenant and treat `export sites` output as what Graph returned, not as a census.

## Live check

After pushing, run from a machine with tenant access:

```bash
pip install -e ".[graph]"
m365-toolkit check-auth
m365-toolkit export users --out ./exports
m365-toolkit logs sign-ins --out ./exports --days 7
```

Confirm: users paginate past the first page, `find_user` returns a match, a team-members export has no double `/groups/groups/` URL errors, and a photo/download round trip works. Report anything that fails with the command and the (redacted) error.

## Development

```bash
pip install -e ".[dev]"
ruff check src tests examples && ruff format --check src tests examples
mypy src
python -m pytest -q
```

Tests use fake Graph clients — no credentials, no network. The `demo` command is the offline end-to-end smoke test.

## Safety

- Real tenant exports are **never** committed. `.gitignore` excludes `exports/`, `raw/`, `*.csv`, and `.env`. The only data in this repository is fictional demo data (`example.com`, invented names, invented GUIDs).
- Exports can contain personal data (names, emails, sign-in IPs). Handle output directories accordingly and do not share them.
- Client-secret auth is appropriate for unattended exports. Prefer certificate credentials and least-privilege permissions in production; scope the app registration to only the permissions the commands you run need.

## License

MIT. See [LICENSE](LICENSE).
