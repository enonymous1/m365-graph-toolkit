"""Google Sheets exporter: write client rows into a worksheet.

Folded in from the October 2025 / May 2026 ``mp-m365-py`` tree
(``m365-google-sheets-sync``), whose only parts this package did not
already cover were this Sheets sink and a mailbox-messages fetch —
the mailbox fetch is already covered by
:class:`~m365_graph_toolkit.clients.mail.MailClient` (paged, via
the SDK), so only the sink is carried over. Fixes over that code:

* Construction never touches the network. The old
  ``SheetsConnector.__init__`` authorised and opened the spreadsheet
  immediately, so ``--dry-run`` still needed valid Google credentials
  and a failed Sheets login leaked the already-created Graph session
  (``Unclosed client session``). Here the service-account key is
  parsed and validated at construction, but authorisation and
  ``open_by_key`` happen lazily on the first write. A caller that
  never writes (a dry run) never talks to Google.
* The key argument is a JSON object string *or* a path to a JSON key
  file, distinguished by whether the path exists on disk — not by
  the old ``startswith("{") is False`` test, which fell through to a
  confusing ``json.loads`` error for a mistyped path. A path that
  does not exist and is not valid JSON fails loudly naming the
  problem; key content is never included in an error message.
* No ``pandas``. The old writer ran ``json_normalize`` /
  ``fillna("").astype(str)`` over Graph payloads, stringifying every
  value. Toolkit rows are already flat ``dict`` rows, so values are
  converted the same way as the CSV exporter: ``None`` becomes
  ``""``, datetimes become ISO strings, lists/dicts become JSON
  strings, numbers and booleans keep their types.
* Empty input follows the CSV exporter: an error unless the caller
  supplies field names, in which case a header-only worksheet is
  written. The old writer wrote ``[[]]`` for an empty frame.
* Worksheet titles are sanitised (Google forbids ``[]*?/\\`` and
  caps titles at 100 characters). The old CLI wrote
  ``f"emails-{user_id}"`` unsanitised.
* Writes are chunked (1,000 data rows per ``update`` by default)
  instead of one unbounded ``update`` for an arbitrarily large
  tenant export.
* Retry covers only the shared retryable statuses (429 / 502 / 503
  / 504, honouring ``Retry-After``), via the same predicate as the
  Graph clients. The old ``_retry`` retried *every* exception,
  including authentication failures, three times with sleeps.

``gspread`` and ``google-auth`` are an optional extra:
``pip install "m365-graph-toolkit[sheets]"``. Importing this module
does not require them; a prebuilt client can be injected instead
(the pattern the Graph clients use), which is also how the tests run
offline. Writing **clears the target worksheet first**, as the old
code did — sync is a replace, not an append. Treat the target
worksheet as owned by the sync.
"""

from __future__ import annotations

import json
import logging
import random
import re
import time
from collections.abc import Callable, Mapping
from pathlib import Path
from typing import Any

from .errors import ExportError, ValidationError
from .exporters import field_names_for, render_cell_value
from .normalize import require_non_empty
from .retry import is_retryable

logger = logging.getLogger(__name__)

_SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets"
_FORBIDDEN_TITLE_CHARS = re.compile(r"[\[\]*?/\\]")
_MAX_TITLE_LENGTH = 100
_MAX_DELAY_SECONDS = 30.0
_DEFAULT_CHUNK_ROWS = 1000

__all__ = ["SheetsExporter", "sanitize_worksheet_title"]


def sanitize_worksheet_title(title: str) -> str:
    """Make ``title`` a valid Google Sheets worksheet title."""
    require_non_empty(title, name="worksheet_title")
    cleaned = _FORBIDDEN_TITLE_CHARS.sub("_", title.strip())
    cleaned = cleaned[:_MAX_TITLE_LENGTH].strip()
    if not cleaned:
        raise ValidationError("worksheet_title must contain at least one valid character")
    return cleaned


def _load_service_account_info(value: str | Mapping[str, Any]) -> dict[str, Any]:
    """Parse a service-account key given as a dict, JSON text, or file path."""
    if isinstance(value, Mapping):
        info = dict(value)
        if not info:
            raise ValidationError("service account key must not be empty")
        return info
    text = value.strip()
    if not text:
        raise ValidationError("service account key must not be empty")
    path = Path(text)
    if path.is_file():
        try:
            text = path.read_text(encoding="utf-8")
        except OSError as exc:
            raise ExportError(f"Could not read service account key file {path}: {exc}") from exc
        try:
            parsed = json.loads(text)
        except json.JSONDecodeError as exc:
            raise ExportError(f"Service account key file {path} is not valid JSON: {exc}") from exc
    else:
        try:
            parsed = json.loads(text)
        except json.JSONDecodeError as exc:
            raise ValidationError(
                "service account key must be a JSON object string or a path "
                f"to a JSON key file: {exc}"
            ) from exc
    if not isinstance(parsed, dict) or not parsed:
        raise ValidationError("service account key must be a non-empty JSON object")
    return parsed


def _retry_after_seconds(error: BaseException) -> float | None:
    """Best-effort ``Retry-After`` (seconds) from a Sheets API error, if any."""
    response = getattr(error, "response", None)
    headers = getattr(response, "headers", None)
    raw: Any = None
    if isinstance(headers, Mapping):
        raw = headers.get("Retry-After")
    try:
        value = float(raw) if raw is not None else None
    except (TypeError, ValueError):
        return None
    return value if value is not None and value >= 0 else None


def _is_worksheet_not_found(error: BaseException) -> bool:
    """Whether ``error`` is gspread's WorksheetNotFound, by class name.

    The name check (rather than an import) keeps this module usable
    with an injected fake client when ``gspread`` is not installed;
    the real exception class is ``gspread.exceptions.WorksheetNotFound``.
    """
    return type(error).__name__ == "WorksheetNotFound"


class SheetsExporter:
    """Write rows to one worksheet of one Google Sheets spreadsheet."""

    def __init__(
        self,
        service_account_key_json_or_path: str | Mapping[str, Any] | None = None,
        spreadsheet_id: str | None = None,
        *,
        client: Any = None,
        max_retries: int = 3,
        base_delay: float = 1.0,
        sleep: Callable[[float], None] = time.sleep,
    ) -> None:
        if max_retries < 0:
            raise ValidationError("max_retries must be >= 0")
        self.spreadsheet_id = require_non_empty(spreadsheet_id or "", name="spreadsheet_id")
        if client is not None:
            # Injected (test / prebuilt) client: no auth is attempted
            # and no service-account key is needed.
            self._client = client
            self._service_account_info: dict[str, Any] | None = None
        else:
            if service_account_key_json_or_path is None:
                raise ValidationError("service account key is required unless a client is injected")
            self._client = None
            self._service_account_info = _load_service_account_info(
                service_account_key_json_or_path
            )
        self.max_retries = max_retries
        self.base_delay = base_delay
        self._sleep = sleep
        self._spreadsheet: Any = None

    def _run_with_retry(self, operation: Callable[[], Any], *, operation_name: str) -> Any:
        """Run one Sheets call with bounded retries (retryable statuses only)."""
        delay = self.base_delay
        for attempt in range(self.max_retries + 1):
            try:
                return operation()
            except Exception as error:  # noqa: BLE001 - predicate decides
                if attempt >= self.max_retries or not is_retryable(error):
                    raise
                retry_after = _retry_after_seconds(error)
                wait = retry_after if retry_after is not None else delay
                wait = min(wait, _MAX_DELAY_SECONDS) + random.uniform(0, 0.5)
                logger.warning(
                    "%s failed (%s); retrying in %.1fs (%d/%d)",
                    operation_name,
                    error,
                    wait,
                    attempt + 1,
                    self.max_retries,
                )
                self._sleep(wait)
                delay = min(delay * 2, _MAX_DELAY_SECONDS)
        raise AssertionError("unreachable")  # pragma: no cover

    def _build_client(self) -> Any:
        """Authorise with the service-account key (lazy import)."""
        try:
            import gspread
            from google.oauth2.service_account import Credentials
        except ImportError as exc:
            raise ExportError(
                "gspread and google-auth are required for Google Sheets sync; "
                "install them with: pip install 'm365-graph-toolkit[sheets]'"
            ) from exc
        try:
            creds = Credentials.from_service_account_info(
                self._service_account_info, scopes=[_SHEETS_SCOPE]
            )
            return gspread.authorize(creds)
        except Exception as exc:  # noqa: BLE001 - wrapped, key content never shown
            raise ExportError(f"Failed to build Google Sheets client: {exc}") from exc

    def _get_spreadsheet(self) -> Any:
        if self._spreadsheet is None:
            client = self._client if self._client is not None else self._build_client()
            self._spreadsheet = self._run_with_retry(
                lambda: client.open_by_key(self.spreadsheet_id),
                operation_name="open spreadsheet",
            )
        return self._spreadsheet

    def _get_or_create_worksheet(
        self, spreadsheet: Any, title: str, *, total_rows: int, total_cols: int
    ) -> Any:
        try:
            worksheet = self._run_with_retry(
                lambda: spreadsheet.worksheet(title),
                operation_name="get worksheet",
            )
        except Exception as error:  # noqa: BLE001 - not-found is the create path
            if not _is_worksheet_not_found(error):
                raise
            rows = max(100, total_rows)
            cols = max(10, total_cols)
            return self._run_with_retry(
                lambda: spreadsheet.add_worksheet(title=title, rows=rows, cols=cols),
                operation_name="create worksheet",
            )
        # An existing worksheet keeps the grid it was created with. A
        # sync into a smaller grid used to fail on the Sheets API
        # update (or silently drop trailing rows, depending on the
        # client), so grow it before clearing and writing. A worksheet
        # is never shrunk: shrinking would delete data outside the
        # sync's own columns.
        existing_rows = getattr(worksheet, "row_count", None)
        existing_cols = getattr(worksheet, "col_count", None)
        if (
            isinstance(existing_rows, int)
            and isinstance(existing_cols, int)
            and (existing_rows < total_rows or existing_cols < total_cols)
        ):
            self._run_with_retry(
                lambda: worksheet.resize(
                    rows=max(total_rows, existing_rows),
                    cols=max(total_cols, existing_cols),
                ),
                operation_name="resize worksheet",
            )
        return worksheet

    def write_rows(
        self,
        rows: list[dict[str, Any]],
        worksheet_title: str,
        *,
        fieldnames: list[str] | None = None,
        chunk_size: int = _DEFAULT_CHUNK_ROWS,
    ) -> int:
        """Replace ``worksheet_title`` with ``rows``; return rows written.

        The worksheet is cleared first. Rows are copied, never mutated.
        """
        title = sanitize_worksheet_title(worksheet_title)
        row_list = [dict(row) for row in rows]
        if fieldnames is None:
            if not row_list:
                raise ValidationError("Cannot write an empty sheet without fieldnames")
            fieldnames = field_names_for(row_list)
        unknown = sorted({key for row in row_list for key in row} - set(fieldnames))
        if unknown:
            raise ValidationError(f"Rows contain fields not in fieldnames: {', '.join(unknown)}")
        if chunk_size < 1:
            raise ValidationError("chunk_size must be >= 1")
        values: list[list[Any]] = [list(fieldnames)]
        values.extend([render_cell_value(row.get(key)) for key in fieldnames] for row in row_list)
        try:
            spreadsheet = self._get_spreadsheet()
            worksheet = self._get_or_create_worksheet(
                spreadsheet,
                title,
                total_rows=len(values),
                total_cols=len(fieldnames),
            )
            self._run_with_retry(lambda: worksheet.clear(), operation_name="clear worksheet")
            data_rows = values[1:]
            chunks: list[list[list[Any]]] = []
            if not data_rows:
                chunks.append(values)
            else:
                for start in range(0, len(data_rows), chunk_size):
                    chunk = data_rows[start : start + chunk_size]
                    if start == 0:
                        chunk = [values[0], *chunk]
                    chunks.append(chunk)
            for index, chunk in enumerate(chunks):
                # First chunk starts at A1 (header row); later chunks
                # start after the header plus the rows already written.
                start_row = 1 if index == 0 else 2 + index * chunk_size
                range_name = f"A{start_row}"

                def _write_chunk(
                    chunk: list[list[Any]] = chunk, range_name: str = range_name
                ) -> Any:
                    return worksheet.update(chunk, range_name=range_name)

                self._run_with_retry(_write_chunk, operation_name="write worksheet")
        except ValidationError:
            raise
        except Exception as exc:  # noqa: BLE001 - wrapped as an export failure
            raise ExportError(
                f"Failed to write {len(row_list)} rows to worksheet {title!r}: {exc}"
            ) from exc
        return len(row_list)
