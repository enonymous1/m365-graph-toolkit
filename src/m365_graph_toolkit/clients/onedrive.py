"""OneDrive client, using the Graph SDK throughout.

Drive access in ``msgraph-sdk`` 1.x is two-step (verified against the
installed 1.34.0 source, not assumed): first fetch the drive itself
(``users/{id}/drive`` or ``me/drive``) to learn its id, then work
through ``drives/{drive-id}`` for items, children, content, search,
and sharing links. v1 mixed two transports: SDK calls for some
methods and hand-rolled ``aiohttp`` calls (via ``_make_request`` /
``_get_headers``, neither of which existed) for the rest, built on a
``self.endpoint`` that was never defined. Every method here goes
through SDK builders.

Uploads and downloads use the drive-item ``content`` builder; files
of 4 MB or more need a resumable upload session, which this client
does not implement — :meth:`upload_file` refuses oversized files
loudly instead of silently sending a request Graph will reject.
"""

from __future__ import annotations

import asyncio
import os
from collections.abc import Iterable
from typing import Any
from urllib.parse import quote

from ..errors import ValidationError
from ..normalize import model_to_dict, require_non_empty
from .base import BaseClient

MAX_SIMPLE_UPLOAD_BYTES = 4 * 1024 * 1024
_GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

DRIVE_FIELDS = ("id", "drive_type", "name", "owner", "quota")


class OneDriveClient(BaseClient):
    """Microsoft Graph OneDrive / drive operations for one user (or ``me``)."""

    def __init__(self, graph_client: Any, *, user_id: str | None = None, **kwargs: Any) -> None:
        super().__init__(graph_client, **kwargs)
        self.user_id = user_id
        self._drive_id_cache: str | None = None

    def for_user(self, user_id: str) -> OneDriveClient:
        """Return a client bound to a different user, sharing retry settings."""
        return OneDriveClient(
            self.client,
            user_id=require_non_empty(user_id, name="user_id"),
            max_retries=self.max_retries,
            base_delay=self.base_delay,
        )

    def _drive_metadata_builder(self) -> Any:
        """Builder for the drive object itself (``/users/{id}/drive`` or ``/me/drive``)."""
        if self.user_id:
            return self.client.users.by_user_id(self.user_id).drive
        return self.client.me.drive

    async def _resolve_drive_id(self) -> str:
        """Resolve and cache the drive id (one extra request, once per client)."""
        if self._drive_id_cache is not None:
            return self._drive_id_cache
        drive = await self._run(
            lambda: self._drive_metadata_builder().get(),
            operation_name="get_drive_info",
        )
        drive_id = getattr(drive, "id", None)
        if isinstance(drive, dict):
            drive_id = drive.get("id")
        if not drive_id:
            raise ValidationError("Graph returned a drive without an id")
        self._drive_id_cache = str(drive_id)
        return self._drive_id_cache

    def _drives_builder(self, drive_id: str) -> Any:
        return self.client.drives.by_drive_id(drive_id)

    def _item_builder(self, drive_id: str, item_id: str) -> Any:
        return self._drives_builder(drive_id).items.by_drive_item_id(item_id)

    async def get_drive_info(self) -> dict[str, Any]:
        """Return drive metadata, including quota when Graph returns it."""
        drive = await self._run(
            lambda: self._drive_metadata_builder().get(),
            operation_name="get_drive_info",
        )
        row = model_to_dict(drive, DRIVE_FIELDS)
        quota = getattr(drive, "quota", None)
        if isinstance(drive, dict):
            quota = drive.get("quota")
        if quota is not None:
            row["quota"] = model_to_dict(quota, ("total", "used", "remaining", "deleted", "state"))
        if row.get("id"):
            self._drive_id_cache = str(row["id"])
        return row

    async def get_drive_quota(self) -> dict[str, Any]:
        """Return the drive's quota dict, or ``{}`` if Graph returns none."""
        info = await self.get_drive_info()
        quota = info.get("quota")
        return dict(quota) if isinstance(quota, dict) else {}

    async def list_items(
        self,
        folder_id: str = "root",
        *,
        select: Iterable[str] | None = None,
    ) -> list[dict[str, Any]]:
        """List the children of a folder (``root`` by default), paged."""
        folder_id = require_non_empty(folder_id, name="folder_id")
        drive_id = await self._resolve_drive_id()
        configuration = self._children_configuration(select=list(select) if select else None)

        async def first_page() -> Any:
            return await self._item_builder(drive_id, folder_id).children.get(
                request_configuration=configuration
            )

        async def next_page(url: str) -> Any:
            return await self._item_builder(drive_id, folder_id).children.with_url(url).get()

        items = await self._collect(first_page, next_page, operation_name="list_items")
        return [self._normalise_item(item) for item in items]

    def _children_configuration(self, *, select: list[str] | None) -> Any:
        from msgraph.generated.drives.item.items.item.children.children_request_builder import (
            ChildrenRequestBuilder,
        )

        query_params = ChildrenRequestBuilder.ChildrenRequestBuilderGetQueryParameters(
            select=select
        )
        return ChildrenRequestBuilder.ChildrenRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

    @staticmethod
    def _normalise_item(item: Any) -> dict[str, Any]:
        row = model_to_dict(
            item,
            ("id", "name", "size", "web_url", "created_date_time", "last_modified_date_time"),
        )
        row["is_folder"] = getattr(item, "folder", None) is not None or (
            isinstance(item, dict) and "folder" in item
        )
        return row

    async def get_item_by_path(self, path: str) -> dict[str, Any]:
        """Return an item addressed by drive-relative path.

        Path addressing has no dedicated builder in SDK 1.34, so the
        item is fetched via ``with_url`` on the drive root builder
        with the documented ``/root:/{path}`` URL form.
        """
        path = require_non_empty(path, name="path")
        drive_id = await self._resolve_drive_id()
        encoded = quote(path.strip("/"), safe="/")
        url = f"{_GRAPH_BASE_URL}/drives/{drive_id}/root:/{encoded}"
        item = await self._run(
            lambda: self._drives_builder(drive_id).root.with_url(url).get(),
            operation_name="get_item_by_path",
        )
        return self._normalise_item(item)

    async def create_folder(self, name: str, parent_id: str = "root") -> dict[str, Any]:
        """Create a folder under ``parent_id``."""
        name = require_non_empty(name, name="name")
        parent_id = require_non_empty(parent_id, name="parent_id")
        drive_id = await self._resolve_drive_id()
        model = self._build_folder_model(name)
        created = await self._run(
            lambda: self._item_builder(drive_id, parent_id).children.post(model),
            operation_name="create_folder",
        )
        return self._normalise_item(created)

    def _build_folder_model(self, name: str) -> Any:
        from msgraph.generated.models.drive_item import DriveItem
        from msgraph.generated.models.folder import Folder

        return DriveItem(name=name, folder=Folder())

    async def upload_file(self, file_path: str, parent_id: str = "root") -> dict[str, Any]:
        """Upload a file smaller than 4 MB by simple upload.

        Larger files raise :class:`~m365_graph_toolkit.errors.ValidationError`:
        Graph requires a resumable upload session for them, and silently
        attempting a simple upload would fail or truncate.
        """
        parent_id = require_non_empty(parent_id, name="parent_id")
        size = os.path.getsize(file_path)
        if size >= MAX_SIMPLE_UPLOAD_BYTES:
            raise ValidationError(
                f"File is {size} bytes; simple upload supports files under "
                f"{MAX_SIMPLE_UPLOAD_BYTES} bytes. Use a resumable upload session."
            )
        with open(file_path, "rb") as handle:
            data = handle.read()
        file_name = os.path.basename(file_path)
        drive_id = await self._resolve_drive_id()
        encoded = quote(file_name, safe="")
        url = f"{_GRAPH_BASE_URL}/drives/{drive_id}/items/{parent_id}:/{encoded}:/content"
        uploaded = await self._run(
            lambda: self._item_builder(drive_id, parent_id).content.with_url(url).put(data),
            operation_name="upload_file",
        )
        return self._normalise_item(uploaded)

    async def download_file(self, item_id: str) -> bytes:
        """Download a file's bytes by item id."""
        item_id = require_non_empty(item_id, name="item_id")
        drive_id = await self._resolve_drive_id()
        content = await self._run(
            lambda: self._item_builder(drive_id, item_id).content.get(),
            operation_name="download_file",
        )
        if content is None:
            return b""
        if isinstance(content, (bytes, bytearray)):
            return bytes(content)
        read = getattr(content, "read", None)
        if callable(read):
            data = read()
            if asyncio.iscoroutine(data):
                data = await data
            return bytes(data)
        return bytes(content)

    async def create_sharing_link(
        self, item_id: str, *, link_type: str = "view", scope: str = "anonymous"
    ) -> dict[str, Any]:
        """Create a sharing link. ``link_type`` is ``view``/``edit``/``embed``."""
        item_id = require_non_empty(item_id, name="item_id")
        if link_type not in {"view", "edit", "embed"}:
            raise ValidationError("link_type must be 'view', 'edit', or 'embed'")
        if scope not in {"anonymous", "organization"}:
            raise ValidationError("scope must be 'anonymous' or 'organization'")
        drive_id = await self._resolve_drive_id()
        body = self._build_sharing_link_body(link_type, scope)
        result = await self._run(
            lambda: self._item_builder(drive_id, item_id).create_link.post(body),
            operation_name="create_sharing_link",
        )
        row = model_to_dict(result, ("id", "type", "scope"))
        link = getattr(result, "link", None)
        if isinstance(result, dict):
            link = result.get("link")
        if link is not None:
            web_url = getattr(link, "web_url", None)
            if isinstance(link, dict):
                web_url = link.get("webUrl", link.get("web_url"))
            row["web_url"] = web_url
        return row

    def _build_sharing_link_body(self, link_type: str, scope: str) -> Any:
        from msgraph.generated.drives.item.items.item.create_link.create_link_post_request_body import (
            CreateLinkPostRequestBody,
        )

        return CreateLinkPostRequestBody(type=link_type, scope=scope)

    async def delete_item(self, item_id: str) -> None:
        """Delete a file or folder. There is no undo; callers must confirm first."""
        item_id = require_non_empty(item_id, name="item_id")
        drive_id = await self._resolve_drive_id()
        await self._run(
            lambda: self._item_builder(drive_id, item_id).delete(),
            operation_name="delete_item",
        )

    async def search(self, query: str) -> list[dict[str, Any]]:
        """Search the drive for ``query`` (paged)."""
        query = require_non_empty(query, name="query")
        drive_id = await self._resolve_drive_id()

        async def first_page() -> Any:
            return await self._drives_builder(drive_id).search_with_q(query).get()

        async def next_page(url: str) -> Any:
            return await self._drives_builder(drive_id).search_with_q(query).with_url(url).get()

        items = await self._collect(first_page, next_page, operation_name="search")
        return [self._normalise_item(item) for item in items]

    async def count_items(self) -> dict[str, int]:
        """Count files/folders in the drive root (paged listing)."""
        items = await self.list_items("root", select=["id", "name"])
        folders = sum(1 for item in items if item.get("is_folder"))
        return {"total_items": len(items), "folders": folders, "files": len(items) - folders}


__all__ = ["MAX_SIMPLE_UPLOAD_BYTES", "OneDriveClient"]
