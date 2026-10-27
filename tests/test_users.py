"""UsersClient tests against fake Graph clients (no SDK, no network)."""

from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

from m365_graph_toolkit.clients.users import UsersClient
from m365_graph_toolkit.errors import ValidationError
from tests.fakes import group_obj, page, user_obj


def _client_with_users(pages_list) -> MagicMock:
    client = MagicMock()
    client.users.get = AsyncMock(return_value=pages_list[0])
    if len(pages_list) > 1:
        client.users.with_url.return_value.get = AsyncMock(return_value=pages_list[1])
    return client


def _patch_config_builders(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(UsersClient, "_list_configuration", lambda self, **kwargs: "LIST_CONFIG")
    monkeypatch.setattr(UsersClient, "_search_configuration", lambda self, term: "SEARCH_CONFIG")


async def test_get_users_pages_and_normalises(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_config_builders(monkeypatch)
    client = _client_with_users(
        [page([user_obj()], next_link="https://example.com/p2"), page([user_obj(id="2")])]
    )
    users = await UsersClient(client).get_users()
    assert [user["id"] for user in users] == [
        "11111111-1111-1111-1111-111111111111",
        "2",
    ]
    assert users[0]["display_name"] == "Ada Example"
    # The configuration object is passed through to the SDK call.
    assert client.users.get.call_args.kwargs["request_configuration"] == "LIST_CONFIG"


async def test_get_users_rejects_bad_top() -> None:
    with pytest.raises(ValidationError):
        await UsersClient(MagicMock()).get_users(top=0)


async def test_get_user(monkeypatch: pytest.MonkeyPatch) -> None:
    client = MagicMock()
    client.users.by_user_id.return_value.get = AsyncMock(return_value=user_obj())
    user = await UsersClient(client).get_user("u1")
    assert user["user_principal_name"] == "ada.example@example.onmicrosoft.com"
    client.users.by_user_id.assert_called_with("u1")


async def test_find_user_returns_first_match(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_config_builders(monkeypatch)
    client = MagicMock()
    client.users.get = AsyncMock(return_value=page([user_obj()]))
    found = await UsersClient(client).find_user("Ada")
    assert found is not None and found["id"].startswith("1111")
    assert client.users.get.call_args.kwargs["request_configuration"] == "SEARCH_CONFIG"


async def test_find_user_returns_none_when_empty(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_config_builders(monkeypatch)
    client = MagicMock()
    client.users.get = AsyncMock(return_value=page([]))
    assert await UsersClient(client).find_user("Nobody") is None


async def test_get_user_memberships_keeps_groups_only() -> None:
    client = MagicMock()
    builder = client.users.by_user_id.return_value
    members = [
        group_obj(),
        SimpleNamespace(id="x", display_name="Not a group", odata_type="#microsoft.graph.user"),
        group_obj(id="g2", display_name="Second"),
    ]
    builder.transitive_member_of.get = AsyncMock(return_value=page(members))
    groups = await UsersClient(client).get_user_memberships("u1")
    assert [group["id"] for group in groups] == [
        "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
        "g2",
    ]
    # No per-member follow-up fetches (v1's N+1 pattern).
    assert client.users.by_user_id.return_value.get.call_count == 0
    assert client.groups.by_group_id.call_count == 0


async def test_get_user_memberships_pages() -> None:
    client = MagicMock()
    builder = client.users.by_user_id.return_value
    builder.transitive_member_of.get = AsyncMock(
        return_value=page([group_obj()], next_link="https://example.com/p2")
    )
    builder.transitive_member_of.with_url.return_value.get = AsyncMock(
        return_value=page([group_obj(id="g2")])
    )
    groups = await UsersClient(client).get_user_memberships("u1")
    assert len(groups) == 2


async def test_memberships_batch_preserves_order_and_survives_failure() -> None:
    client = MagicMock()

    # get_user_memberships is stubbed at the instance level: user "bad"
    # raises, the others succeed. Batch order must still be input order.
    users_client = UsersClient(client)
    users = [
        {"id": "u1", "display_name": "One", "user_principal_name": "one@example.com"},
        {"id": "bad", "display_name": "Bad", "user_principal_name": "bad@example.com"},
        {"id": "u2", "display_name": "Two", "user_principal_name": "two@example.com"},
    ]

    async def fake_memberships(user_id: str):
        if user_id == "bad":
            raise RuntimeError("boom")
        return [{"id": "g1", "display_name": "G", "group_types": ["Unified"]}]

    users_client.get_user_memberships = fake_memberships  # type: ignore[method-assign]
    rows = await users_client.get_user_memberships_batch(users, max_concurrent=2)
    assert [row["user_id"] for row in rows] == ["u1", "u2"]
    assert rows[0]["group_type"] == "Unified"


async def test_update_user_rejects_unknown_fields(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        UsersClient, "_build_user_model", lambda self, data: SimpleNamespace(**data)
    )
    with pytest.raises(ValidationError):
        # Unknown-field rejection lives in _build_user_model in the
        # real implementation; here the stub bypasses it, so assert
        # the client rejects an empty update instead.
        await UsersClient(MagicMock()).update_user("u1", {})
    # With the stubbed builder, a valid update reaches the SDK.
    client = MagicMock()
    client.users.by_user_id.return_value.patch = AsyncMock(return_value=user_obj())
    updated = await UsersClient(client).update_user("u1", {"display_name": "New"})
    assert updated["display_name"] == "Ada Example"


async def test_update_user_rejects_unknown_fields_real_builder() -> None:
    # Without stubbing, _build_user_model validates before any SDK
    # import: unknown fields raise even when msgraph is absent.
    with pytest.raises(ValidationError, match="Unknown user field"):
        UsersClient(MagicMock())._build_user_model({"displayName": "New"})


async def test_create_user_requires_password_profile() -> None:
    with pytest.raises(ValidationError, match="password_profile"):
        await UsersClient(MagicMock()).create_user(
            {
                "display_name": "New",
                "mail_nickname": "new",
                "user_principal_name": "new@example.com",
            }
        )


async def test_get_user_photo_returns_none_on_404() -> None:
    client = MagicMock()
    error = RuntimeError("Request_ResourceNotFound")
    error.response_status_code = 404  # type: ignore[attr-defined]
    client.users.by_user_id.return_value.photo.content.get = AsyncMock(side_effect=error)
    # BaseClient._run translates 404 to NotFoundError; get_user_photo
    # converts that to None.
    assert await UsersClient(client, max_retries=0).get_user_photo("u1") is None


async def test_get_user_photo_returns_bytes() -> None:
    client = MagicMock()
    client.users.by_user_id.return_value.photo.content.get = AsyncMock(return_value=b"\x89PNG")
    assert await UsersClient(client).get_user_photo("u1") == b"\x89PNG"


async def test_get_manager_returns_none_when_absent() -> None:
    client = MagicMock()
    error = RuntimeError("404 Not Found")
    client.users.by_user_id.return_value.manager.get = AsyncMock(side_effect=error)
    assert await UsersClient(client, max_retries=0).get_manager("u1") is None
