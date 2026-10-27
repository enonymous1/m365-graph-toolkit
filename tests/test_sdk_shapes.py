"""SDK shape tests: verify the builder chains the clients use exist in the
installed ``msgraph-sdk`` version, with the exact method and class names
the code calls.

These tests need the Graph dependencies (``pip install -e ".[dev]"``)
but no credentials and no network: the client is constructed with a
dummy credential and only attribute access is exercised. If a future
SDK release renames a builder, this file fails first, in CI, instead
of the failure surfacing on a maintainer's machine during a live run.

Note: constructing ``GraphServiceClient`` reads httpx's proxy
environment on some hosts; on the original build box that required
unsetting ``NO_PROXY`` (an httpx IPv6 parsing quirk), which is an
environment property, not a library behaviour.
"""

from __future__ import annotations

import pytest

msgraph = pytest.importorskip("msgraph")


class _DummyCredential:
    def get_token(self, *scopes, **kwargs):  # noqa: ANN002, ANN003, ANN202
        class _Token:
            expires_on = 1
            token = "dummy"

        return _Token()


@pytest.fixture()
def graph_client():
    from msgraph import GraphServiceClient

    return GraphServiceClient(
        credentials=_DummyCredential(), scopes=["https://graph.microsoft.com/.default"]
    )


def test_user_builder_chains(graph_client) -> None:
    users = graph_client.users.by_user_id("u1")
    assert hasattr(users.transitive_member_of, "get")
    assert hasattr(users.transitive_member_of, "with_url")
    assert hasattr(users.direct_reports, "get")
    assert hasattr(users.manager, "get")
    assert hasattr(users.photo.content, "get")
    assert hasattr(users.drive, "get")
    assert hasattr(graph_client.users, "with_url")


def test_group_builder_chains(graph_client) -> None:
    members = graph_client.groups.by_group_id("g1").members
    assert hasattr(members, "get")
    assert hasattr(members, "with_url")
    assert hasattr(graph_client.groups, "with_url")


def test_site_and_list_builder_chains(graph_client) -> None:
    site = graph_client.sites.by_site_id("s1")
    assert hasattr(site.lists, "get")
    assert hasattr(site.lists, "with_url")
    items = site.lists.by_list_id("l1").items
    assert hasattr(items, "get")
    assert hasattr(items, "post")
    assert hasattr(items, "with_url")
    assert hasattr(items, "by_list_item_id")
    assert hasattr(graph_client.sites, "with_url")


def test_audit_log_builder_chains(graph_client) -> None:
    assert hasattr(graph_client.audit_logs.sign_ins, "get")
    assert hasattr(graph_client.audit_logs.sign_ins, "with_url")
    assert hasattr(graph_client.audit_logs.directory_audits, "get")


def test_drive_builder_chains(graph_client) -> None:
    drive = graph_client.drives.by_drive_id("d1")
    item = drive.items.by_drive_item_id("root")
    assert hasattr(item.children, "get")
    assert hasattr(item.children, "post")
    assert hasattr(item.children, "with_url")
    assert hasattr(item.content, "get")
    assert hasattr(item.content, "put")
    assert hasattr(item.create_link, "post")
    assert hasattr(drive.search_with_q("q"), "get")
    assert hasattr(drive.root.content, "put")


def test_query_parameter_construction() -> None:
    from msgraph.generated.users.users_request_builder import UsersRequestBuilder

    params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
        search='"displayName:ada"', count=True, select=["id"], top=10
    )
    assert params.search == '"displayName:ada"'
    config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(query_parameters=params)
    config.headers.add("ConsistencyLevel", "eventual")
    assert config.query_parameters is params


def test_model_construction() -> None:
    from msgraph.generated.models.drive_item import DriveItem
    from msgraph.generated.models.field_value_set import FieldValueSet
    from msgraph.generated.models.folder import Folder
    from msgraph.generated.models.list_item import ListItem
    from msgraph.generated.models.user import User

    user = User(display_name="Ada", user_principal_name="ada@example.com")
    assert user.display_name == "Ada"
    assert DriveItem(name="folder", folder=Folder()).folder is not None
    item = ListItem(fields=FieldValueSet(additional_data={"Title": "x"}))
    assert item.fields.additional_data == {"Title": "x"}

    from msgraph.generated.drives.item.items.item.create_link.create_link_post_request_body import (
        CreateLinkPostRequestBody,
    )

    body = CreateLinkPostRequestBody(type="view", scope="anonymous")
    assert body.type == "view"
    assert body.scope == "anonymous"
