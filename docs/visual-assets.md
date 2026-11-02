# Visual assets: capture plan

No image files are shipped with this repository yet, deliberately.
Every capture below must use fictional data (`m365-toolkit demo`) or
a test tenant / test spreadsheet. Never capture a real tenant's users,
mail, sign-in IPs, or spreadsheet contents, and redact any tenant id,
client id, or spreadsheet id that appears in a terminal recording.

## 1. Demo run GIF — `docs/assets/demo-run.gif`

**What it depicts:** a terminal running the offline demo and listing
its output, proving the tool works with no credentials.

**Capture:**

```bash
pip install -e ".[graph]"
m365-toolkit demo --out ./demo-output
ls ./demo-output
```

Record the terminal (for example with `asciinema` or a screen
recorder), trim to the command, the "Wrote 24 demo files" line, and
the file listing. Keep it under ~30 seconds, no audio.

**README placement:** directly under the intro paragraph, before
the "Description" heading. It is the first visual a visitor sees.

## 2. Demo CSV screenshot — `docs/assets/demo-csv.png`

**What it depicts:** the demo users CSV open in a spreadsheet,
showing the `snake_case` header row (`id`, `display_name`,
`user_principal_name`, `mail`, ...) and fictional rows (Ada Example
and company, `example.com` addresses only).

**Capture:** run `m365-toolkit demo --out ./demo-output`, open
`users_demo_*.csv` in a spreadsheet, widen the first columns so the
headers are readable, and screenshot the header row plus 3–4 rows.
No real data may appear in this file by construction — it is demo
output — but check the visible rows before publishing.

**README placement:** in "Use", immediately after the export
command block, before the Sheets sync commands.

## 3. Sheets sync screenshot — `docs/assets/sheets-sync.png`

**What it depicts:** a test spreadsheet after a real
`m365-toolkit sync users` run: the `users` worksheet tab selected,
the header row, and fictional/test users below it.

**Capture:** against a test spreadsheet only, with a test tenant or
fictional rows written by hand if no test tenant exists:

```bash
m365-toolkit sync users --spreadsheet <test-sheet-id> --worksheet users
```

Screenshot the browser tab with the worksheet tab visible. Redact
the spreadsheet id from the URL bar (crop it out) before publishing.

**README placement:** in "Use", immediately after the sync command
block, before the dry-run paragraph.

## 4. Architecture diagram — already in the README

The Mermaid diagram in the README's "Architecture" section needs no
external file and renders on GitHub. Its source is
[`architecture.mmd`](architecture.mmd) in this directory; keep the
two identical when the layering changes.

## Rules for all assets

- Fictional or test data only. If a capture accidentally shows real
  tenant data, delete it and re-capture; do not blur real data and
  ship it.
- No `<img>` link for an asset may be added to the README until the
  file exists in `docs/assets/` — a broken image link is worse than
  no image.
- File names are fixed by this plan so the README references do not
  drift: `demo-run.gif`, `demo-csv.png`, `sheets-sync.png`.
