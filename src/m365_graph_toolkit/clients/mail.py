"""Mail client: a user's Outlook messages, normalised for export.

Folded in from the February/May 2026 ``mp-m365-app`` tree. Fixes over
that code:

* The old ``OutlookManager`` fetched only the first page of messages
  per user and exported through a CSV helper that rendered nested
  SDK objects (``from``, ``to_recipients``, ``body``) as opaque
  strings. Messages here are paged and returned as plain ``dict``
  rows with sender/recipient addresses and the body content
  resolved.
* Per-mailbox failures are handled in the batch method, not by
  string-matching error codes inside a retry loop. The old code
  skipped a mailbox on ``MailboxNotEnabledForRESTAPI`` /
  ``ErrorAccessDenied`` but retried everything else five times
  including 403s; here a single-user fetch surfaces the Graph
  error, and :meth:`MailClient.get_messages_for_users` logs and
  skips the failing mailbox while the batch continues.
* No hardcoded home-directory paths and no ``sys.path`` hacks.

``body`` is the message body's HTML/text content as Graph returns
it; ``body_preview`` is Graph's truncated preview. Both are kept
because the old export used the preview while the full body is what
an audit usually needs — callers that only want previews can drop
the ``body`` key from the rows.
"""

from __future__ import annotations

import asyncio
from collections.abc import Iterable
from typing import Any

from ..errors import ValidationError
from ..normalize import model_to_dict, require_non_empty
from .base import BaseClient

MESSAGE_FIELDS = (
    "id",
    "subject",
    "body_preview",
    "importance",
    "is_read",
    "has_attachments",
    "received_date_time",
    "sent_date_time",
    "created_date_time",
)


def _field(obj: Any, name: str) -> Any:
    if isinstance(obj, dict):
        return obj.get(name)
    return getattr(obj, name, None)


def _recipient_emails(recipients: Any) -> list[str]:
    emails: list[str] = []
    for recipient in recipients or []:
        address = _field(recipient, "email_address")
        value = _field(address, "address") if address is not None else None
        if value is not None:
            emails.append(str(value))
    return emails


def _normalise_message(message: Any) -> dict[str, Any]:
    row = model_to_dict(message, MESSAGE_FIELDS)
    sender = _field(message, "from_")
    sender_address = _field(sender, "email_address") if sender is not None else None
    row["from_email"] = _field(sender_address, "address") if sender_address is not None else None
    row["from_name"] = _field(sender_address, "name") if sender_address is not None else None
    row["to_emails"] = _recipient_emails(_field(message, "to_recipients"))
    row["cc_emails"] = _recipient_emails(_field(message, "cc_recipients"))
    row["bcc_emails"] = _recipient_emails(_field(message, "bcc_recipients"))
    body = _field(message, "body")
    row["body"] = _field(body, "content") if body is not None else None
    row["body_content_type"] = _field(body, "content_type") if body is not None else None
    return row


class MailClient(BaseClient):
    """Microsoft Graph Outlook mail operations."""

    async def get_user_messages(
        self,
        user_id: str,
        *,
        select: Iterable[str] | None = None,
        top: int | None = None,
    ) -> list[dict[str, Any]]:
        """Return a user's messages, following pagination."""
        user_id = require_non_empty(user_id, name="user_id")
        if top is not None and top < 1:
            raise ValidationError("top must be >= 1")
        configuration = self._messages_configuration(
            select=list(select) if select else None, top=top
        )

        async def first_page() -> Any:
            return await self.client.users.by_user_id(user_id).messages.get(
                request_configuration=configuration
            )

        async def next_page(url: str) -> Any:
            return await self.client.users.by_user_id(user_id).messages.with_url(url).get()

        messages = await self._collect(first_page, next_page, operation_name="get_user_messages")
        return [_normalise_message(message) for message in messages]

    def _messages_configuration(self, *, select: list[str] | None, top: int | None) -> Any:
        from msgraph.generated.users.item.messages.messages_request_builder import (
            MessagesRequestBuilder,
        )

        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            select=select, top=top
        )
        return MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

    async def get_messages_for_users(
        self,
        users: list[dict[str, Any]],
        *,
        max_concurrent: int = 5,
    ) -> list[dict[str, Any]]:
        """Fetch messages for many users concurrently, in input order.

        A user whose mailbox cannot be read (not provisioned, no
        access) contributes zero rows and a warning; failures never
        abort the whole batch.
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
                    messages = await self.get_user_messages(str(user_id))
                except Exception as error:  # noqa: BLE001 - logged, batch continues
                    self.logger.warning("Mail fetch failed for user %s: %s", user_id, error)
                    return []
            for message in messages:
                message["user_id"] = user.get("id")
                message["user_display_name"] = user.get("display_name")
            return messages

        results = await asyncio.gather(*(fetch_one(user) for user in users))
        return [row for batch in results for row in batch]


__all__ = ["MailClient", "MESSAGE_FIELDS"]
