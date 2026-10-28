"""Teams channel-messages client: teams, channels, and channel messages.

Folded in from the February/May 2026 ``mp-m365-app`` tree
(``graph/messages.py``). Fixes over that code:

* Errors are raised, not swallowed. The old exporter caught every
  error, printed it, and returned ``[]`` — an empty export looked
  identical to a team with no messages and to a 403. Every method
  here raises :class:`~m365_graph_toolkit.errors.GraphError` (or
  ``NotFoundError`` for 404s) through the shared retry/pagination
  plumbing, so a permission problem is visible.
* Pagination uses ``with_url`` on the same builder throughout.
  The old code paged teams and channels through raw request-builder
  constructors over ``client.request_adapter`` and messages through
  ``with_url`` — three patterns for one operation.
* Rows are plain ``dict``\\ s with the sender and body resolved.
  The old code rendered ``channel_identity`` and the raw ``from_``
  object into CSV cells and dropped ``print()`` output instead of
  logging.
* No hardcoded home-directory export path; callers decide where
  rows go (see :mod:`m365_graph_toolkit.exporters`).

This client uses the ``/teams`` endpoint, which is distinct from
:meth:`GroupsClient.get_teams` (groups with a team provisioned).
Reading channel messages needs ``ChannelMessage.Read.All`` (or
``ChannelMessage.Read.Group`` for delegated access); listing teams
and channels needs ``Team.ReadBasic.All`` at minimum.
"""

from __future__ import annotations

from collections.abc import Iterable
from typing import Any

from ..errors import ValidationError
from ..normalize import model_to_dict, require_non_empty
from .base import BaseClient

TEAM_FIELDS = (
    "id",
    "display_name",
    "description",
    "created_date_time",
    "is_archived",
)

CHANNEL_FIELDS = (
    "id",
    "display_name",
    "description",
    "membership_type",
    "created_date_time",
)

MESSAGE_FIELDS = (
    "id",
    "subject",
    "importance",
    "message_type",
    "created_date_time",
    "last_modified_date_time",
    "deleted_date_time",
)


def _field(obj: Any, name: str) -> Any:
    if isinstance(obj, dict):
        return obj.get(name)
    return getattr(obj, name, None)


def _normalise_team(team: Any) -> dict[str, Any]:
    return model_to_dict(team, TEAM_FIELDS)


def _normalise_channel(channel: Any) -> dict[str, Any]:
    return model_to_dict(channel, CHANNEL_FIELDS)


def _normalise_channel_message(message: Any) -> dict[str, Any]:
    row = model_to_dict(message, MESSAGE_FIELDS)
    body = _field(message, "body")
    row["body"] = _field(body, "content") if body is not None else None
    row["body_content_type"] = _field(body, "content_type") if body is not None else None
    sender = _field(message, "from_")
    row["sender_type"] = None
    row["sender_id"] = None
    row["sender_display_name"] = None
    if sender is not None:
        for kind in ("user", "application", "device"):
            identity = _field(sender, kind)
            if identity is not None:
                row["sender_type"] = kind
                row["sender_id"] = _field(identity, "id")
                row["sender_display_name"] = _field(identity, "display_name")
                break
    replies = _field(message, "replies")
    if isinstance(replies, list):
        row["reply_count"] = len(replies)
    else:
        row["reply_count"] = _field(message, "replies_count")
    reactions = _field(message, "reactions")
    if isinstance(reactions, list):
        row["reaction_count"] = len(reactions)
        row["reaction_types"] = sorted(
            str(_field(reaction, "reaction_type"))
            for reaction in reactions
            if _field(reaction, "reaction_type") is not None
        )
    else:
        row["reaction_count"] = _field(message, "reactions_count")
        row["reaction_types"] = []
    return row


class TeamsMessagesClient(BaseClient):
    """Microsoft Graph Teams channel-message operations."""

    async def get_teams(
        self,
        *,
        select: Iterable[str] | None = None,
        top: int | None = None,
    ) -> list[dict[str, Any]]:
        """Return teams via the ``/teams`` endpoint, following pagination."""
        if top is not None and top < 1:
            raise ValidationError("top must be >= 1")
        configuration = self._teams_configuration(select=list(select) if select else None, top=top)

        async def first_page() -> Any:
            return await self.client.teams.get(request_configuration=configuration)

        async def next_page(url: str) -> Any:
            return await self.client.teams.with_url(url).get()

        teams = await self._collect(first_page, next_page, operation_name="get_teams")
        return [_normalise_team(team) for team in teams]

    def _teams_configuration(self, *, select: list[str] | None, top: int | None) -> Any:
        from msgraph.generated.teams.teams_request_builder import TeamsRequestBuilder

        query_params = TeamsRequestBuilder.TeamsRequestBuilderGetQueryParameters(
            select=select, top=top
        )
        return TeamsRequestBuilder.TeamsRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

    async def get_channels(self, team_id: str, *, top: int | None = None) -> list[dict[str, Any]]:
        """Return a team's channels, following pagination."""
        team_id = require_non_empty(team_id, name="team_id")
        if top is not None and top < 1:
            raise ValidationError("top must be >= 1")

        async def first_page() -> Any:
            return await self.client.teams.by_team_id(team_id).channels.get()

        async def next_page(url: str) -> Any:
            return await self.client.teams.by_team_id(team_id).channels.with_url(url).get()

        channels = await self._collect(first_page, next_page, operation_name="get_channels")
        return [_normalise_channel(channel) for channel in channels]

    async def get_channel_messages(
        self,
        team_id: str,
        channel_id: str,
        *,
        select: Iterable[str] | None = None,
        top: int | None = None,
    ) -> list[dict[str, Any]]:
        """Return a channel's messages, following pagination."""
        team_id = require_non_empty(team_id, name="team_id")
        channel_id = require_non_empty(channel_id, name="channel_id")
        if top is not None and top < 1:
            raise ValidationError("top must be >= 1")
        configuration = self._messages_configuration(
            select=list(select) if select else None, top=top
        )

        async def first_page() -> Any:
            return await (
                self.client.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .messages.get(request_configuration=configuration)
            )

        async def next_page(url: str) -> Any:
            return (
                await self.client.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .messages.with_url(url)
                .get()
            )

        messages = await self._collect(first_page, next_page, operation_name="get_channel_messages")
        rows = [_normalise_channel_message(message) for message in messages]
        for row in rows:
            row["team_id"] = team_id
            row["channel_id"] = channel_id
        return rows

    def _messages_configuration(self, *, select: list[str] | None, top: int | None) -> Any:
        from msgraph.generated.teams.item.channels.item.messages.messages_request_builder import (
            MessagesRequestBuilder,
        )

        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            select=select, top=top
        )
        return MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )


__all__ = ["TeamsMessagesClient"]
