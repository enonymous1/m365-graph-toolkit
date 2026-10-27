"""Groups, SharePoint, and OneDrive client tests (fake Graph clients)."""

from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

from m365_graph_toolkit.clients.groups import GroupsClient
from m365_graph_toolkit.clients.onedrive import OneDriveClient
from m365_graph_toolkit.clients.sharepoint import SharePointClient
from m365_graph_toolkit.errors import ValidationError
from tests.fakes import group_obj, page

# --- Groups / Teams -------------------------------------------------------


async def test_get_teams_uses_team_filter(monkeypatch: pytest.MonkeyPatch) -> None:
    captured: dict = {}
    monkeypatch.setattr(
        GroupsClient,
        "_groups_configuration",
        lambda self, **kwargs: captured.update(kwargs) or "CONFIG",
    )
    client = MagicMock()
    client.groups.get = AsyncMock(return_value=page([group_obj()]))
    teams = await GroupsClient(client).get_teams()
    assert len(teams) == 1
    assert "Team" in captured["filter_query"]
    assert client.groups.get.call_args.kwargs["request_configuration"] == "CONFIG"


async def test_get_group_members_pages() -> None:
    client = MagicMock()
    builder = client.groups.by_group_id.return_value
    member = SimpleNamespace(
        id="u1",
        display_name="Ada",
        user_principal_name="ada@example.com",
        mail="ada@example.com",
        user_type="Member",
        account_enabled=True,
        odata_type="#microsoft.graph.user",
    )
    builder.members.get = AsyncMock(return_value=page([member], next_link="p2"))
    builder.members.with_url.return_value.get = AsyncMock(return_value=page([member]))
    members = await GroupsClient(client).get_group_members("g1")
    assert len(members) == 2
    assert members[0]["display_name"] == "Ada"


async def test_get_team_members_is_get_group_members() -> None:
    client = MagicMock()
    builder = client.groups.by_group_id.return_value
    builder.members.get = AsyncMock(return_value=page([]))
    assert await GroupsClient(client).get_team_members("g1") == []


# --- SharePoint -----------------------------------------------------------


def _sharepoint_client() -> tuple[MagicMock, SharePointClient]:
    client = MagicMock()
    return client, SharePointClient(client)


async def test_get_root_site() -> None:
    client, sp = _sharepoint_client()
    site = SimpleNamespace(
        id="root",
        display_name="Root",
        name="root",
        description="",
        web_url="https://contoso.example",
        created_date_time=None,
        last_modified_date_time=None,
    )
    client.sites.by_site_id.return_value.get = AsyncMock(return_value=site)
    result = await sp.get_root_site()
    assert result["display_name"] == "Root"
    client.sites.by_site_id.assert_called_with("root")


async def test_get_list_items_expands_fields_by_default(monkeypatch: pytest.MonkeyPatch) -> None:
    captured: dict = {}
    monkeypatch.setattr(
        SharePointClient,
        "_items_configuration",
        lambda self, **kwargs: captured.update(kwargs) or "CONFIG",
    )
    client, sp = _sharepoint_client()
    item = SimpleNamespace(
        id="i1",
        created_date_time=None,
        last_modified_date_time=None,
        fields=SimpleNamespace(
            additional_data={"Title": "Hello", "file": {"name": "a.txt", "url": "u"}}
        ),
    )
    items_builder = client.sites.by_site_id.return_value.lists.by_list_id.return_value.items
    items_builder.get = AsyncMock(return_value=page([item]))
    items = await sp.get_list_items("s1", "l1")
    assert captured["expand"] == ["fields"]
    assert items[0]["fields"]["Title"] == "Hello"


async def test_get_documents_returns_only_files() -> None:
    client, sp = _sharepoint_client()

    # get_list_items is stubbed to return normalised rows directly.
    async def fake_items(site_id, list_id, **kwargs):  # noqa: ANN001, ANN202
        return [
            {"id": "1", "fields": {"file": {"name": "a.txt", "url": "u1"}}},
            {"id": "2", "fields": {"Title": "not a file"}},
            {"id": "3", "fields": {}},
        ]

    sp.get_list_items = fake_items  # type: ignore[method-assign]
    docs = await sp.get_documents("s1", "l1")
    assert docs == [{"id": "1", "name": "a.txt", "url": "u1"}]


async def test_get_site_backup_empty_site_has_empty_lists() -> None:
    client, sp = _sharepoint_client()
    site = SimpleNamespace(
        id="s1",
        display_name="Empty",
        name="empty",
        description="",
        web_url="https://contoso.example/sites/empty",
        created_date_time=None,
        last_modified_date_time=None,
    )
    client.sites.by_site_id.return_value.get = AsyncMock(return_value=site)
    client.sites.by_site_id.return_value.lists.get = AsyncMock(return_value=page([]))
    backup = await sp.get_site_backup("s1")
    # v1 crashed here with IndexError on lists_data[0].
    assert backup["lists"] == []
    assert backup["site"]["display_name"] == "Empty"


async def test_build_list_item_rejects_empty_fields() -> None:
    with pytest.raises(ValidationError):
        SharePointClient(MagicMock())._build_list_item({})


# --- OneDrive -------------------------------------------------------------


async def test_onedrive_get_drive_info_includes_quota() -> None:
    client = MagicMock()
    quota = SimpleNamespace(total=100, used=25, remaining=75, deleted=0, state="normal")
    drive = SimpleNamespace(id="d1", drive_type="business", name="Drive", owner=None, quota=quota)
    client.me.drive.get = AsyncMock(return_value=drive)
    info = await OneDriveClient(client).get_drive_info()
    assert info["quota"]["used"] == 25
    assert info["quota"]["total"] == 100


async def test_onedrive_for_user_uses_user_drive() -> None:
    client = MagicMock()
    drive = SimpleNamespace(id="d1", drive_type="business", name="D", owner=None, quota=None)
    client.users.by_user_id.return_value.drive.get = AsyncMock(return_value=drive)
    info = await OneDriveClient(client).for_user("u1").get_drive_info()
    assert info["id"] == "d1"
    client.users.by_user_id.assert_called_with("u1")


async def test_onedrive_upload_refuses_large_files(
    tmp_path, monkeypatch: pytest.MonkeyPatch
) -> None:
    small = tmp_path / "small.txt"
    small.write_bytes(b"hello")
    monkeypatch.setattr("os.path.getsize", lambda path: 5 * 1024 * 1024)
    with pytest.raises(ValidationError, match="resumable"):
        await OneDriveClient(MagicMock()).upload_file(str(small))


async def test_onedrive_upload_small_file(tmp_path) -> None:
    small = tmp_path / "small.txt"
    small.write_bytes(b"hello")
    client = MagicMock()
    drive_meta = SimpleNamespace(id="d1")
    client.me.drive.get = AsyncMock(return_value=drive_meta)
    item = SimpleNamespace(
        id="i1",
        name="small.txt",
        size=5,
        web_url="u",
        created_date_time=None,
        last_modified_date_time=None,
        folder=None,
    )
    content_builder = (
        client.drives.by_drive_id.return_value.items.by_drive_item_id.return_value.content
    )
    content_builder.with_url.return_value.put = AsyncMock(return_value=item)
    uploaded = await OneDriveClient(client).upload_file(str(small))
    assert uploaded["name"] == "small.txt"
    assert uploaded["is_folder"] is False
    client.drives.by_drive_id.assert_called_with("d1")


async def test_onedrive_create_sharing_link_validates() -> None:
    with pytest.raises(ValidationError):
        await OneDriveClient(MagicMock()).create_sharing_link("i1", link_type="bogus")
    with pytest.raises(ValidationError):
        await OneDriveClient(MagicMock()).create_sharing_link("i1", scope="everyone")


async def test_onedrive_count_items() -> None:
    client = MagicMock()
    drive_client = OneDriveClient(client)
    items = [
        {"id": "1", "is_folder": True},
        {"id": "2", "is_folder": False},
        {"id": "3", "is_folder": False},
    ]

    async def fake_list(folder_id="root", **kwargs):  # noqa: ANN001, ANN202
        return items

    drive_client.list_items = fake_list  # type: ignore[method-assign]
    assert await drive_client.count_items() == {"total_items": 3, "folders": 1, "files": 2}
