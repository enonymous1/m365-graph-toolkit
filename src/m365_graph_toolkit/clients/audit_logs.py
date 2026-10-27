"""Audit-log client: sign-in and directory audit logs.

Fixes over v1's ``UnifiedLogsClient``:

* The time filter is applied to *every* log type, and
  ``get_directory_audit_logs(days=...)`` honours its ``days``
  argument — v1 hardcoded 30 days and only filtered sign-ins.
* Pagination is real: ``_get_all_pages`` never existed in v1, so
  every call crashed. Pages are followed through ``odata_next_link``.
* ``get_audit_logs`` builds an OData ``$filter`` from datetimes
  formatted as UTC ISO-8601. Naive datetimes are rejected rather
  than silently treated as UTC.
* Fetching "all users' sign-ins" no longer issues one unfiltered
  tenant-wide query per user. It issues one tenant-wide, time-boxed
  query and groups the result by user (see
  :func:`m365_graph_toolkit.analysis.group_logs_by_user`). That is
  both faster and the only shape that survives Graph throttling at
  tenant scale.
"""

from __future__ import annotations

from collections.abc import Iterable
from datetime import datetime, timedelta, timezone
from typing import Any

from ..errors import ValidationError
from ..normalize import model_to_dict, require_non_empty
from .base import BaseClient

SIGN_IN_FIELDS = (
    "id",
    "created_date_time",
    "user_display_name",
    "user_principal_name",
    "user_id",
    "app_display_name",
    "app_id",
    "ip_address",
    "client_app_used",
    "is_interactive",
    "risk_state",
    "risk_level_aggregated",
    "risk_level_during_sign_in",
)

DIRECTORY_AUDIT_FIELDS = (
    "id",
    "activity_date_time",
    "activity_display_name",
    "category",
    "result",
    "logged_by_service",
    "initiated_by",
    "target_resources",
)


def _require_aware(value: datetime, *, name: str) -> datetime:
    if value.tzinfo is None:
        raise ValidationError(f"{name} must be timezone-aware (use UTC)")
    return value


def _format_utc(value: datetime) -> str:
    return value.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _time_range(
    start_time: datetime | None,
    end_time: datetime | None,
    days: int | None,
) -> tuple[datetime, datetime]:
    end = _require_aware(end_time, name="end_time") if end_time else datetime.now(timezone.utc)
    if days is not None:
        if days < 0:
            raise ValidationError("days must be >= 0")
        start = end - timedelta(days=days)
    elif start_time is not None:
        start = _require_aware(start_time, name="start_time")
    else:
        start = end - timedelta(days=7)
    if start > end:
        raise ValidationError("start_time must not be after end_time")
    return start, end


class AuditLogsClient(BaseClient):
    """Microsoft Graph audit-log operations (sign-ins, directory audits)."""

    async def get_audit_logs(
        self,
        *,
        log_type: str = "signIn",
        start_time: datetime | None = None,
        end_time: datetime | None = None,
        days: int | None = None,
        filter_query: str | None = None,
        user_id: str | None = None,
        top: int = 1000,
    ) -> list[dict[str, Any]]:
        """Fetch audit logs of one type for a time range.

        ``log_type`` is ``"signIn"`` or ``"directoryAudits"``.
        ``days`` overrides ``start_time``. The time filter always uses
        the log type's timestamp field (``createdDateTime`` for
        sign-ins, ``activityDateTime`` for directory audits).
        """
        if log_type not in {"signIn", "directoryAudits"}:
            raise ValidationError("log_type must be 'signIn' or 'directoryAudits'")
        if top < 1:
            raise ValidationError("top must be >= 1")
        start, end = _time_range(start_time, end_time, days)
        timestamp_field = "createdDateTime" if log_type == "signIn" else "activityDateTime"
        filters = [
            f"{timestamp_field} ge {_format_utc(start)} and {timestamp_field} le {_format_utc(end)}"
        ]
        if user_id:
            if log_type == "signIn":
                filters.append(f"userId eq '{user_id}'")
            else:
                filters.append(f"initiatedBy/user/userPrincipalName eq '{user_id}'")
        if filter_query:
            filters.append(f"({filter_query})")
        configuration = self._logs_configuration(
            log_type=log_type, filter_value=" and ".join(filters), top=top
        )

        if log_type == "signIn":
            collection = self.client.audit_logs.sign_ins
            fields: tuple[str, ...] = SIGN_IN_FIELDS
        else:
            collection = self.client.audit_logs.directory_audits
            fields = DIRECTORY_AUDIT_FIELDS

        async def first_page() -> Any:
            return await collection.get(request_configuration=configuration)

        async def next_page(url: str) -> Any:
            return await collection.with_url(url).get()

        logs = await self._collect(first_page, next_page, operation_name="get_audit_logs")
        return [self._normalise_log(log, fields) for log in logs]

    @staticmethod
    def _normalise_log(log: Any, fields: tuple[str, ...]) -> dict[str, Any]:
        """Normalise a log row, flattening its nested ``status`` object.

        Sign-in success/failure lives in ``status.errorCode``. Without
        this step the analysers cannot tell a failed sign-in from a
        successful one, because ``status`` is a nested SDK object.
        """
        row = model_to_dict(log, fields)
        status = getattr(log, "status", None)
        if status is None and isinstance(log, dict):
            status = log.get("status")
        if status is not None:
            if isinstance(status, dict):
                row["status"] = {
                    "error_code": status.get("errorCode", status.get("error_code")),
                    "failure_reason": status.get("failureReason", status.get("failure_reason")),
                }
            else:
                row["status"] = {
                    "error_code": getattr(status, "error_code", None),
                    "failure_reason": getattr(status, "failure_reason", None),
                }
        # Directory audits carry initiated_by / target_resources as
        # nested objects; keep them as-is for JSON, as JSON strings
        # for CSV (the exporter handles both).
        return row

    def _logs_configuration(self, *, log_type: str, filter_value: str, top: int) -> Any:
        if log_type == "signIn":
            from msgraph.generated.audit_logs.sign_ins.sign_ins_request_builder import (
                SignInsRequestBuilder,
            )

            query_params = SignInsRequestBuilder.SignInsRequestBuilderGetQueryParameters(
                filter=filter_value, top=top
            )
            return SignInsRequestBuilder.SignInsRequestBuilderGetRequestConfiguration(
                query_parameters=query_params
            )
        from msgraph.generated.audit_logs.directory_audits.directory_audits_request_builder import (
            DirectoryAuditsRequestBuilder,
        )

        audit_query_params = (
            DirectoryAuditsRequestBuilder.DirectoryAuditsRequestBuilderGetQueryParameters(
                filter=filter_value, top=top
            )
        )
        return DirectoryAuditsRequestBuilder.DirectoryAuditsRequestBuilderGetRequestConfiguration(
            query_parameters=audit_query_params
        )

    async def get_user_sign_in_logs(
        self,
        user_id: str,
        *,
        days: int = 30,
        success_only: bool = False,
    ) -> list[dict[str, Any]]:
        """Fetch sign-in logs for one user."""
        user_id = require_non_empty(user_id, name="user_id")
        filter_query = "status/errorCode eq 0" if success_only else None
        return await self.get_audit_logs(
            log_type="signIn",
            days=days,
            filter_query=filter_query,
            user_id=user_id,
        )

    async def get_all_sign_in_logs(
        self, *, days: int = 7, success_only: bool = False
    ) -> dict[str, list[dict[str, Any]]]:
        """Fetch all sign-ins for the range, grouped by user id.

        One tenant-wide query, grouped locally. v1 issued one query
        per user with no shared time box.
        """
        from ..analysis import group_logs_by_user

        filter_query = "status/errorCode eq 0" if success_only else None
        logs = await self.get_audit_logs(log_type="signIn", days=days, filter_query=filter_query)
        return group_logs_by_user(logs)

    async def get_directory_audit_logs(
        self,
        *,
        days: int = 30,
        filter_query: str | None = None,
    ) -> list[dict[str, Any]]:
        """Fetch directory audit logs for the last ``days`` days."""
        return await self.get_audit_logs(
            log_type="directoryAudits", days=days, filter_query=filter_query
        )

    async def get_sign_in_logs_for_users(
        self,
        user_ids: Iterable[str],
        *,
        days: int = 7,
        success_only: bool = False,
    ) -> dict[str, list[dict[str, Any]]]:
        """Fetch sign-ins for selected users via one query, filtered locally.

        ``user_ids`` must be non-empty; unknown ids simply contribute
        no rows. Prefer :meth:`get_all_sign_in_logs` when the whole
        tenant is wanted.
        """
        ids = [require_non_empty(uid, name="user_id") for uid in user_ids]
        if not ids:
            raise ValidationError("user_ids must not be empty")
        grouped = await self.get_all_sign_in_logs(days=days, success_only=success_only)
        wanted = set(ids)
        return {uid: grouped.get(uid, []) for uid in ids if uid in wanted}


__all__ = ["AuditLogsClient"]
