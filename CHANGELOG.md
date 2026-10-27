# Changelog

All notable changes to this project are documented here. The format
follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [2.0.0] — 2026-10-27

A rewrite of the July 2026 `mp-user-project` codebase as a single,
installable, test-gated package. The original project name appears
only in this note; no original history, identifiers, or tenant data
were carried over.

### Added

- Installable package `m365_graph_toolkit` and CLI `m365-toolkit`
  (`export users|memberships|teams|sites|onedrive`, `logs sign-ins`,
  `demo`, `check-auth`).
- One authentication path (`AuthManager`, client credentials from
  `MS_TENANT_ID` / `MS_CLIENT_ID` / `MS_CLIENT_SECRET`), with lazy
  credential imports so offline use needs no Graph dependencies.
- SDK-only clients for users, groups/teams, SharePoint, OneDrive,
  and audit logs. Every read returns plain `dict` rows.
- Bounded retry (429/502/503/504, `Retry-After` honoured, jitter,
  exactly `max_retries` retries) and real `odata_next_link`
  pagination in every list method.
- Non-destructive CSV/JSON exporters with caller-chosen output
  directories, no mutation of input rows, and no file rotation.
- Pure analysis functions (team membership, OneDrive usage,
  sign-in logs) over the clients' `snake_case` rows.
- Offline `demo` command over fictional data, so the export and
  analysis paths run end to end without credentials or a tenant.
- Test suite with fake Graph clients covering every fixed bug below;
  ruff, mypy, and GitHub Actions CI; MIT LICENSE; `.gitattributes`
  pinning LF endings.

### Fixed (verified against the v1 tree)

- Every API client failed to import: the `auth.ms_graph` module it
  imported never existed. Removed; one auth path remains.
- `GraphTeamClient` and `GraphOneDriveClient` crashed on
  `self.endpoint`, which the base client never set.
- `_paginated_request`, `_make_request`, `_get_headers`, and
  `_get_all_pages` were called but never defined. The raw-`aiohttp`
  transport is gone; all calls use the SDK.
- `GraphUtils` built `/groups/groups/{id}/members`. Team members now
  go through the SDK groups builder.
- Two more `GraphUserClient` implementations (in `graph/users.py`)
  and a third auth design are removed; one `UsersClient` remains.
- Membership fetching no longer does two extra requests per member
  (N+1); groups are filtered from `transitive_member_of` directly.
- `find_user` no longer builds an invalid parenthesised `$search`
  string; it uses `$search` with `ConsistencyLevel: eventual`.
- SharePoint list-item writes no longer post plain dicts; they send
  SDK `ListItem` models with a `FieldValueSet`.
- `full_site_backup` no longer crashes on an empty site
  (`lists_data[0]`) and pages lists and items; replaced by
  `get_site_backup`, which returns data instead of writing files.
- `get_documents` now expands `fields` and handles empty results.
- Audit logs: the time filter now applies to directory audits too,
  and `get_directory_audit_logs(days=...)` honours `days` (v1
  hardcoded 30). "All users' sign-ins" is one tenant-wide query
  grouped locally, not one query per user.
- `save_onedrive_analysis_to_csv` awaited a synchronous method; the
  equivalent CLI path is synchronous where it should be.
- `save_all_*` helpers called a non-existent `get_all_users()` and
  passed SDK objects to dict-expecting savers; both are gone.
- `CsvFileManager.save_data` no longer mutates caller dicts via
  `pop()`; `DataExporter`'s silent file rotation and dead per-call
  cache are gone.
- `is_internal_user` no longer hardcodes one employer's domain or
  ignores the UPN; the domain is an explicit argument and either
  mail or UPN matching it counts.
- `update_user` no longer sends camelCase keys or `None` values that
  could clear fields; it validates `snake_case` fields and rejects
  unknown ones.
- `get_user_education` (which read a non-existent `.schools`
  attribute) is removed; education is not a Graph user property.
- Hardcoded tenant name, tenant site-id prefixes, and a real work
  email in code/examples are removed. The only tenant data in this
  repository is fictional demo data.
- Token handling: the v1 `__main__` block printed the first 10
  characters of an access token. Nothing in v2 logs or prints a
  token value.

### Not verified

- No live Graph run. The builder chains were checked against the
  installed `msgraph-sdk` source on the build machine, and every
  behaviour that can be tested offline is tested, but a real tenant
  run (auth, pagination, `$search`, photo content, upload/download)
  still needs the maintainer's tenant. See README, "Live check".

## [1.0.0] — 2026-07-20 (original, private)

Initial `mp-user-project` structure. Not published; superseded by 2.0.0.
