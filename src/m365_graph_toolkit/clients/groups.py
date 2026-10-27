"""Groups and Teams client.

Teams are Microsoft 365 groups with ``resourceProvisioningOptions``
containing ``Team``. v1's ``GraphUtils`` built the members URL as
``/groups/groups/{id}/members`` (double ``/groups``) and its
``GraphTeamClient`` crashed on construction; both are fixed here by
using the SDK builders for every call.
"""

from __future__ import annotations

from collections.abc import Iterable
from typing import Any

from ..errors import ValidationError
from ..normalize import model_to_dict, require_non_empty
from .base import BaseClient

GROUP_FIELDS = (
    "id",
    "display_name",
    "description",
    "mail",
    "mail_nickname",
    "mail_enabled",
    "visibility",
    "group_types",
    "created_date_time",
)

MEMBER_FIELDS = (
    "id",
    "display_name",
    "user_principal_name",
    "mail",
    "user_type",
    "account_enabled",
    "odata_type",
)


def _normalise_group(group: Any) -> dict[str, Any]:
    return model_to_dict(group, GROUP_FIELDS)


def _normalise_member(member: Any) -> dict[str, Any]:
    return model_to_dict(member, MEMBER_FIELDS)


class GroupsClient(BaseClient):
    """Microsoft Graph group and team operations."""

    async def get_groups(
        self,
        *,
        select: Iterable[str] | None = None,
        filter_query: str | None = None,
        top: int | None = None,
    ) -> list[dict[str, Any]]:
        """Return every group, following pagination."""
        if top is not None and top < 1:
            raise ValidationError("top must be >= 1")
        configuration = self._groups_configuration(
            select=list(select) if select else list(GROUP_FIELDS),
            filter_query=filter_query,
            top=top,
        )

        async def first_page() -> Any:
            return await self.client.groups.get(request_configuration=configuration)

        async def next_page(url: str) -> Any:
            return await self.client.groups.with_url(url).get()

        groups = await self._collect(first_page, next_page, operation_name="get_groups")
        return [_normalise_group(group) for group in groups]

    def _groups_configuration(
        self, *, select: list[str], filter_query: str | None, top: int | None
    ) -> Any:
        from msgraph.generated.groups.groups_request_builder import GroupsRequestBuilder

        query_params = GroupsRequestBuilder.GroupsRequestBuilderGetQueryParameters(
            select=select,
            filter=filter_query,
            top=top,
        )
        return GroupsRequestBuilder.GroupsRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

    async def get_teams(
        self, *, select: Iterable[str] | None = None, top: int | None = None
    ) -> list[dict[str, Any]]:
        """Return Microsoft 365 groups that have a team provisioned."""
        return await self.get_groups(
            select=select,
            filter_query="resourceProvisioningOptions/any(x:x eq 'Team')",
            top=top,
        )

    async def get_group(self, group_id: str) -> dict[str, Any]:
        """Return one group by id."""
        group_id = require_non_empty(group_id, name="group_id")
        group = await self._run(
            lambda: self.client.groups.by_group_id(group_id).get(),
            operation_name="get_group",
        )
        return _normalise_group(group)

    async def get_group_members(self, group_id: str) -> list[dict[str, Any]]:
        """Return the members of a group, following pagination."""
        group_id = require_non_empty(group_id, name="group_id")

        async def first_page() -> Any:
            return await self.client.groups.by_group_id(group_id).members.get()

        async def next_page(url: str) -> Any:
            return await self.client.groups.by_group_id(group_id).members.with_url(url).get()

        members = await self._collect(first_page, next_page, operation_name="get_group_members")
        return [_normalise_member(member) for member in members]

    async def get_team_members(self, team_id: str) -> list[dict[str, Any]]:
        """Return the members of a team (a team *is* a group)."""
        return await self.get_group_members(team_id)


__all__ = ["GROUP_FIELDS", "GroupsClient", "MEMBER_FIELDS"]
