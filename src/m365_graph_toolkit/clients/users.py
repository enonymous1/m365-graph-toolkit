"""Users client: read, search, memberships, and profile writes.

All read methods return plain ``dict`` rows (see
:mod:`m365_graph_toolkit.normalize`). Write methods accept plain dicts
with ``snake_case`` field names matching the Graph SDK model
(``display_name``, not ``displayName``) and reject unknown fields
instead of silently dropping them, which is what v1 did.

Membership fetching uses ``transitive_member_of`` directly and keeps
only directory objects whose ``odata_type`` is a group. v1 fetched
each membership twice more (directory object, then group) — an N+1
pattern that tripled request volume for large tenants.
"""

from __future__ import annotations

import asyncio
from collections.abc import Iterable
from typing import Any

from ..errors import ValidationError
from ..normalize import model_to_dict, require_non_empty
from .base import BaseClient

USER_FIELDS = (
    "id",
    "display_name",
    "given_name",
    "surname",
    "user_principal_name",
    "mail",
    "job_title",
    "department",
    "office_location",
    "account_enabled",
    "user_type",
    "mobile_phone",
    "business_phones",
    "created_date_time",
)

# Fields that may be patched with UsersClient.update_user.
_UPDATEABLE_FIELDS = frozenset(
    {
        "display_name",
        "given_name",
        "surname",
        "user_principal_name",
        "mobile_phone",
        "office_location",
        "job_title",
        "department",
        "account_enabled",
        "business_phones",
        "about_me",
        "interests",
        "skills",
        "responsibilities",
    }
)

_GROUP_FIELDS = ("id", "display_name", "mail", "group_types", "visibility", "odata_type")


def _normalise_user(user: Any) -> dict[str, Any]:
    return model_to_dict(user, USER_FIELDS)


def _normalise_group(group: Any) -> dict[str, Any]:
    row = model_to_dict(group, _GROUP_FIELDS)
    row["object_type"] = row.pop("odata_type", None)
    return row


class UsersClient(BaseClient):
    """Microsoft Graph user operations."""

    async def get_users(
        self,
        *,
        select: Iterable[str] | None = None,
        top: int | None = None,
        filter_query: str | None = None,
    ) -> list[dict[str, Any]]:
        """Return every user, following pagination. ``top`` caps page size, not total."""
        if top is not None and top < 1:
            raise ValidationError("top must be >= 1")
        first = self.client.users.get
        # The SDK query-parameter object is built lazily so unit tests
        # can substitute a fake client without the SDK installed.
        request_configuration = self._list_configuration(
            select=list(select) if select else list(USER_FIELDS),
            top=top,
            filter_query=filter_query,
        )

        async def first_page() -> Any:
            return await first(request_configuration=request_configuration)

        async def next_page(url: str) -> Any:
            return await self.client.users.with_url(url).get()

        pages = await self._collect(first_page, next_page, operation_name="get_users")
        return [_normalise_user(user) for user in pages]

    def _list_configuration(
        self,
        *,
        select: list[str],
        top: int | None,
        filter_query: str | None,
    ) -> Any:
        """Build the SDK list request configuration (live path)."""
        from msgraph.generated.users.users_request_builder import UsersRequestBuilder

        query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
            select=select,
            top=top,
            filter=filter_query,
        )
        return UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

    async def get_user(self, user_id: str) -> dict[str, Any]:
        """Return one user by id or userPrincipalName."""
        user_id = require_non_empty(user_id, name="user_id")
        user = await self._run(
            lambda: self.client.users.by_user_id(user_id).get(),
            operation_name=f"get_user {user_id}",
        )
        return _normalise_user(user)

    async def find_user(self, search_term: str) -> dict[str, Any] | None:
        """Search by display name / mail / UPN; return the first match or None.

        Uses Graph ``$search`` with advanced queries (``ConsistencyLevel:
        eventual`` + ``$count``), which is required for ``$search`` on
        users. The search string is a single quoted ``$search`` value;
        v1 wrapped it in parentheses and ``OR`` operators that Graph
        does not accept in that form.
        """
        search_term = require_non_empty(search_term, name="search_term")
        configuration = self._search_configuration(search_term)
        response = await self._run(
            lambda: self.client.users.get(request_configuration=configuration),
            operation_name="find_user",
        )
        values = getattr(response, "value", None) or []
        return _normalise_user(values[0]) if values else None

    def _search_configuration(self, search_term: str) -> Any:
        from msgraph.generated.users.users_request_builder import UsersRequestBuilder

        escaped = search_term.replace('"', '\\"')
        query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
            search=f'"displayName:{escaped}"',
            count=True,
            select=list(USER_FIELDS),
        )
        configuration = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )
        configuration.headers.add("ConsistencyLevel", "eventual")
        return configuration

    async def get_user_memberships(self, user_id: str) -> list[dict[str, Any]]:
        """Return the groups a user belongs to (transitive), without N+1 fetches."""
        user_id = require_non_empty(user_id, name="user_id")

        async def first_page() -> Any:
            return await self.client.users.by_user_id(user_id).transitive_member_of.get()

        async def next_page(url: str) -> Any:
            return (
                await self.client.users.by_user_id(user_id).transitive_member_of.with_url(url).get()
            )

        members = await self._collect(first_page, next_page, operation_name="get_user_memberships")
        groups = [
            _normalise_group(member)
            for member in members
            if getattr(member, "odata_type", None) == "#microsoft.graph.group"
            or (isinstance(member, dict) and member.get("odata_type") == "#microsoft.graph.group")
        ]
        return groups

    async def get_user_memberships_batch(
        self,
        users: list[dict[str, Any]],
        *,
        max_concurrent: int = 5,
    ) -> list[dict[str, Any]]:
        """Fetch memberships for many users concurrently, in input order.

        A user whose fetch fails contributes zero rows and a warning;
        failures never abort the whole batch and are never hidden.
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
                    groups = await self.get_user_memberships(str(user_id))
                except Exception as error:  # noqa: BLE001 - logged, batch continues
                    self.logger.warning("Membership fetch failed for user %s: %s", user_id, error)
                    return []
            rows: list[dict[str, Any]] = []
            for group in groups:
                rows.append(
                    {
                        "user_id": user.get("id"),
                        "user_display_name": user.get("display_name"),
                        "user_principal_name": user.get("user_principal_name"),
                        "group_id": group.get("id"),
                        "group_display_name": group.get("display_name"),
                        "group_type": ",".join(group.get("group_types") or []) or "security",
                    }
                )
            return rows

        results = await asyncio.gather(*(fetch_one(user) for user in users))
        return [row for batch in results for row in batch]

    async def get_direct_reports(self, user_id: str) -> list[dict[str, Any]]:
        """Return a user's direct reports, following pagination.

        Every other list method in this package pages; this one
        previously returned only the SDK's first page, silently
        truncating large reports. The ``with_url`` shape was checked
        against the installed ``msgraph-sdk`` builder (it exists on
        the direct-reports builder, as on every other collection
        builder used here).
        """
        user_id = require_non_empty(user_id, name="user_id")

        async def first_page() -> Any:
            return await self.client.users.by_user_id(user_id).direct_reports.get()

        async def next_page(url: str) -> Any:
            return await self.client.users.by_user_id(user_id).direct_reports.with_url(url).get()

        reports = await self._collect(first_page, next_page, operation_name="get_direct_reports")
        return [_normalise_user(report) for report in reports]

    async def get_manager(self, user_id: str) -> dict[str, Any] | None:
        """Return a user's manager, or None if the user has none."""
        user_id = require_non_empty(user_id, name="user_id")
        try:
            manager = await self._run(
                lambda: self.client.users.by_user_id(user_id).manager.get(),
                operation_name="get_manager",
            )
        except Exception as error:  # noqa: BLE001 - 404 means "no manager"
            if "404" in str(error) or getattr(error, "status_code", None) == 404:
                return None
            raise
        return _normalise_user(manager) if manager is not None else None

    async def get_user_photo(self, user_id: str) -> bytes | None:
        """Return a user's photo bytes, or None if the user has no photo."""
        user_id = require_non_empty(user_id, name="user_id")
        try:
            content = await self._run(
                lambda: self.client.users.by_user_id(user_id).photo.content.get(),
                operation_name="get_user_photo",
            )
        except Exception as error:  # noqa: BLE001 - 404 means "no photo"
            if "404" in str(error) or getattr(error, "status_code", None) == 404:
                return None
            raise
        if content is None:
            return None
        if isinstance(content, (bytes, bytearray)):
            return bytes(content)
        read = getattr(content, "read", None)
        if callable(read):
            data = read()
            if asyncio.iscoroutine(data):
                data = await data
            return bytes(data)
        return bytes(content)

    def _build_user_model(self, data: dict[str, Any]) -> Any:
        """Build an SDK User model from snake_case fields (live path).

        Validation runs before the SDK import, so an unknown field
        raises ``ValidationError`` even when ``msgraph-sdk`` is not
        installed (base install), instead of a misleading
        ``ModuleNotFoundError``.
        """
        unknown = sorted(set(data) - _UPDATEABLE_FIELDS - {"password_profile"})
        if unknown:
            raise ValidationError(f"Unknown user field(s): {', '.join(unknown)}")
        from msgraph.generated.models.user import User

        return User(**{key: value for key, value in data.items() if value is not None})

    async def update_user(self, user_id: str, data: dict[str, Any]) -> dict[str, Any]:
        """Patch a user. Only keys present in ``data`` are sent; ``None`` values are dropped."""
        user_id = require_non_empty(user_id, name="user_id")
        if not data:
            raise ValidationError("data must not be empty")
        model = self._build_user_model(dict(data))
        updated = await self._run(
            lambda: self.client.users.by_user_id(user_id).patch(model),
            operation_name="update_user",
        )
        return _normalise_user(updated)

    async def create_user(self, data: dict[str, Any]) -> dict[str, Any]:
        """Create a user. ``password_profile`` with a ``password`` is required by Graph."""
        required = {"display_name", "mail_nickname", "user_principal_name"}
        missing = sorted(required - set(data))
        if missing:
            raise ValidationError(f"Missing required field(s): {', '.join(missing)}")
        profile = data.get("password_profile")
        if not isinstance(profile, dict) or not profile.get("password"):
            raise ValidationError("password_profile with a 'password' is required to create a user")
        model = self._build_user_model(dict(data))
        created = await self._run(
            lambda: self.client.users.post(model),
            operation_name="create_user",
        )
        return _normalise_user(created)

    async def delete_user(self, user_id: str) -> None:
        """Delete a user. There is no undo; callers must confirm before calling."""
        user_id = require_non_empty(user_id, name="user_id")
        await self._run(
            lambda: self.client.users.by_user_id(user_id).delete(),
            operation_name="delete_user",
        )


__all__ = ["USER_FIELDS", "UsersClient"]
