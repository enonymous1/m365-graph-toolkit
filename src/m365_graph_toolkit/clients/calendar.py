"""Calendar client: a user's events, normalised for export.

Folded in from the February/May 2026 ``mp-m365-app`` tree. Fixes over
that code:

* The old ``CalendarsManager`` fetched only the first page of events
  per user, retried every ``ODataError`` (including 403s) five times,
  and exported through a CSV helper that rendered nested SDK
  objects (``start``, ``organizer``, ``attendees``) as opaque
  strings. Events here are paged, retried only on retryable
  statuses, and returned as plain ``dict`` rows with the start/end
  time, organiser, attendees, and location flattened.
* No hardcoded home-directory paths and no ``sys.path`` hacks; the
  client takes an already-built Graph client.

Events come from ``GET /users/{id}/calendar/events`` (the user's
default calendar). Graph returns them in no documented order and
with no date filter applied here; callers that need a window should
filter the returned rows on ``start_date_time``.
"""

from __future__ import annotations

import asyncio
from collections.abc import Iterable
from typing import Any

from ..errors import ValidationError
from ..normalize import model_to_dict, require_non_empty, scalar_value
from .base import BaseClient

EVENT_FIELDS = (
    "id",
    "subject",
    "body_preview",
    "importance",
    "is_all_day",
    "is_cancelled",
    "is_online_meeting",
    "online_meeting_url",
    "created_date_time",
    "last_modified_date_time",
)


def _field(obj: Any, name: str) -> Any:
    if isinstance(obj, dict):
        return obj.get(name)
    return getattr(obj, name, None)


def _email_address(recipient: Any) -> str | None:
    if recipient is None:
        return None
    address = _field(recipient, "email_address")
    if address is None:
        return None
    value = _field(address, "address")
    return str(value) if value is not None else None


def _normalise_event(event: Any) -> dict[str, Any]:
    row = model_to_dict(event, EVENT_FIELDS)
    for prefix, key in (("start", "start"), ("end", "end")):
        window = _field(event, key)
        row[f"{prefix}_date_time"] = scalar_value(_field(window, "date_time")) if window else None
        row[f"{prefix}_time_zone"] = _field(window, "time_zone") if window else None
    location = _field(event, "location")
    row["location"] = _field(location, "display_name") if location is not None else None
    row["organizer_email"] = _email_address(_field(event, "organizer"))
    attendees = _field(event, "attendees") or []
    row["attendee_emails"] = [
        address for address in (_email_address(a) for a in attendees) if address is not None
    ]
    return row


class CalendarClient(BaseClient):
    """Microsoft Graph calendar operations."""

    async def get_user_events(
        self,
        user_id: str,
        *,
        select: Iterable[str] | None = None,
        top: int | None = None,
    ) -> list[dict[str, Any]]:
        """Return a user's calendar events, following pagination."""
        user_id = require_non_empty(user_id, name="user_id")
        if top is not None and top < 1:
            raise ValidationError("top must be >= 1")
        configuration = self._events_configuration(select=list(select) if select else None, top=top)

        async def first_page() -> Any:
            return await self.client.users.by_user_id(user_id).calendar.events.get(
                request_configuration=configuration
            )

        async def next_page(url: str) -> Any:
            return await self.client.users.by_user_id(user_id).calendar.events.with_url(url).get()

        events = await self._collect(first_page, next_page, operation_name="get_user_events")
        return [_normalise_event(event) for event in events]

    def _events_configuration(self, *, select: list[str] | None, top: int | None) -> Any:
        from msgraph.generated.users.item.calendar.events.events_request_builder import (
            EventsRequestBuilder,
        )

        query_params = EventsRequestBuilder.EventsRequestBuilderGetQueryParameters(
            select=select, top=top
        )
        return EventsRequestBuilder.EventsRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

    async def get_events_for_users(
        self,
        users: list[dict[str, Any]],
        *,
        max_concurrent: int = 5,
    ) -> list[dict[str, Any]]:
        """Fetch calendar events for many users concurrently, in input order.

        A user whose fetch fails contributes zero rows and a warning;
        failures never abort the whole batch.
        """
        if max_concurrent < 1:
            raise ValidationError("max_concurrent must be >= 1")
        semaphore = asyncio.Semaphore(max_concurrent)

        async def fetch_one(user: dict[str, Any]) -> list[dict[str, Any]]:
            user_id = user.get("id")
            if not user_id:
                return []
            async with semaphore:
                try:
                    events = await self.get_user_events(str(user_id))
                except Exception as error:  # noqa: BLE001 - logged, batch continues
                    self.logger.warning("Calendar fetch failed for user %s: %s", user_id, error)
                    return []
            for event in events:
                event["user_id"] = user.get("id")
                event["user_display_name"] = user.get("display_name")
            return events

        results = await asyncio.gather(*(fetch_one(user) for user in users))
        return [row for batch in results for row in batch]


__all__ = ["CalendarClient", "EVENT_FIELDS"]
