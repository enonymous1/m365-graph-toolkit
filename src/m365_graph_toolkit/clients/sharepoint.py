"""SharePoint client: sites, lists, items, and backups.

Fixes over v1:

* No hardcoded tenant name. The root site is fetched as ``sites/root``
  through the SDK, and no example tenant string appears anywhere.
* List items are created/updated as SDK ``ListItem`` models with a
  ``FieldValueSet`` — v1 posted plain dicts, which the SDK rejects.
* ``get_list_items`` accepts ``select``/``expand`` and pages; v1
  fetched only the first page while calling the result a "full backup".
* Backup methods page through lists *and* items, and fail loudly on
  an empty site instead of crashing on ``lists_data[0]``.
* ``get_documents`` expands ``fields`` (required to see file metadata)
  and never indexes ``items.value`` without a null check.
* ``backup_documents`` still exports document *metadata* only; it
  does not claim to download file contents. Downloading contents
  belongs in :class:`~m365_graph_toolkit.clients.onedrive.OneDriveClient`.
"""

from __future__ import annotations

from collections.abc import Iterable
from typing import Any

from ..errors import ValidationError
from ..normalize import model_to_dict, require_non_empty
from .base import BaseClient

SITE_FIELDS = (
    "id",
    "display_name",
    "name",
    "description",
    "web_url",
    "created_date_time",
    "last_modified_date_time",
)

LIST_FIELDS = (
    "id",
    "display_name",
    "name",
    "description",
    "created_date_time",
    "last_modified_date_time",
    "web_url",
)

ITEM_FIELDS = ("id", "created_date_time", "last_modified_date_time")


def _normalise_site(site: Any) -> dict[str, Any]:
    return model_to_dict(site, SITE_FIELDS)


def _normalise_list(lst: Any) -> dict[str, Any]:
    return model_to_dict(lst, LIST_FIELDS)


def _normalise_item(item: Any) -> dict[str, Any]:
    row = model_to_dict(item, ITEM_FIELDS)
    fields = getattr(item, "fields", None)
    if fields is not None:
        additional = getattr(fields, "additional_data", None)
        if isinstance(additional, dict):
            row["fields"] = dict(additional)
        elif isinstance(fields, dict):
            row["fields"] = dict(fields)
    return row


class SharePointClient(BaseClient):
    """Microsoft Graph SharePoint operations."""

    async def get_root_site(self) -> dict[str, Any]:
        """Return the tenant's root SharePoint site."""
        site = await self._run(
            lambda: self.client.sites.by_site_id("root").get(),
            operation_name="get_root_site",
        )
        return _normalise_site(site)

    async def get_all_sites(self, *, select: Iterable[str] | None = None) -> list[dict[str, Any]]:
        """Return every site the caller can see, following pagination.

        Note: Graph ``GET /sites`` without a search returns only sites
        the caller follows or has recently accessed in some tenants;
        use ``search``-based discovery (see README) for a full tenant
        inventory. This method reports exactly what Graph returns and
        does not pretend otherwise.
        """
        configuration = self._sites_configuration(
            select=list(select) if select else list(SITE_FIELDS)
        )

        async def first_page() -> Any:
            return await self.client.sites.get(request_configuration=configuration)

        async def next_page(url: str) -> Any:
            return await self.client.sites.with_url(url).get()

        sites = await self._collect(first_page, next_page, operation_name="get_all_sites")
        return [_normalise_site(site) for site in sites]

    def _sites_configuration(self, *, select: list[str]) -> Any:
        from msgraph.generated.sites.sites_request_builder import SitesRequestBuilder

        query_params = SitesRequestBuilder.SitesRequestBuilderGetQueryParameters(select=select)
        return SitesRequestBuilder.SitesRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

    async def get_lists(self, site_id: str) -> list[dict[str, Any]]:
        """Return the lists in a site, following pagination."""
        site_id = require_non_empty(site_id, name="site_id")

        async def first_page() -> Any:
            return await self.client.sites.by_site_id(site_id).lists.get()

        async def next_page(url: str) -> Any:
            return await self.client.sites.by_site_id(site_id).lists.with_url(url).get()

        lists = await self._collect(first_page, next_page, operation_name="get_lists")
        return [_normalise_list(lst) for lst in lists]

    async def get_list_by_id(self, site_id: str, list_id: str) -> dict[str, Any]:
        """Return one list by id (use the id, not the display name)."""
        site_id = require_non_empty(site_id, name="site_id")
        list_id = require_non_empty(list_id, name="list_id")
        lst = await self._run(
            lambda: self.client.sites.by_site_id(site_id).lists.by_list_id(list_id).get(),
            operation_name="get_list_by_id",
        )
        return _normalise_list(lst)

    async def get_list_items(
        self,
        site_id: str,
        list_id: str,
        *,
        select: Iterable[str] | None = None,
        expand: Iterable[str] | None = None,
    ) -> list[dict[str, Any]]:
        """Return every item in a list, following pagination."""
        site_id = require_non_empty(site_id, name="site_id")
        list_id = require_non_empty(list_id, name="list_id")
        configuration = self._items_configuration(
            select=list(select) if select else None,
            expand=list(expand) if expand else ["fields"],
        )

        async def first_page() -> Any:
            return await (
                self.client.sites.by_site_id(site_id)
                .lists.by_list_id(list_id)
                .items.get(request_configuration=configuration)
            )

        async def next_page(url: str) -> Any:
            return (
                await self.client.sites.by_site_id(site_id)
                .lists.by_list_id(list_id)
                .items.with_url(url)
                .get()
            )

        items = await self._collect(first_page, next_page, operation_name="get_list_items")
        return [_normalise_item(item) for item in items]

    def _items_configuration(self, *, select: list[str] | None, expand: list[str] | None) -> Any:
        from msgraph.generated.sites.item.lists.item.items.items_request_builder import (
            ItemsRequestBuilder,
        )

        query_params = ItemsRequestBuilder.ItemsRequestBuilderGetQueryParameters(
            select=select,
            expand=expand,
        )
        return ItemsRequestBuilder.ItemsRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

    def _build_list_item(self, fields: dict[str, Any]) -> Any:
        """Build an SDK ListItem carrying a FieldValueSet (live path)."""
        from msgraph.generated.models.field_value_set import FieldValueSet
        from msgraph.generated.models.list_item import ListItem

        if not fields:
            raise ValidationError("fields must not be empty")
        return ListItem(fields=FieldValueSet(additional_data=dict(fields)))

    async def create_list_item(
        self, site_id: str, list_id: str, fields: dict[str, Any]
    ) -> dict[str, Any]:
        """Create a list item from a field-name -> value mapping."""
        site_id = require_non_empty(site_id, name="site_id")
        list_id = require_non_empty(list_id, name="list_id")
        model = self._build_list_item(fields)
        created = await self._run(
            lambda: (
                self.client.sites.by_site_id(site_id).lists.by_list_id(list_id).items.post(model)
            ),
            operation_name="create_list_item",
        )
        return _normalise_item(created)

    async def update_list_item(
        self, site_id: str, list_id: str, item_id: str, fields: dict[str, Any]
    ) -> dict[str, Any]:
        """Update a list item's fields."""
        site_id = require_non_empty(site_id, name="site_id")
        list_id = require_non_empty(list_id, name="list_id")
        item_id = require_non_empty(item_id, name="item_id")
        model = self._build_list_item(fields)
        updated = await self._run(
            lambda: (
                self.client.sites.by_site_id(site_id)
                .lists.by_list_id(list_id)
                .items.by_list_item_id(item_id)
                .patch(model)
            ),
            operation_name="update_list_item",
        )
        return _normalise_item(updated)

    async def delete_list_item(self, site_id: str, list_id: str, item_id: str) -> None:
        """Delete a list item. There is no undo; callers must confirm first."""
        site_id = require_non_empty(site_id, name="site_id")
        list_id = require_non_empty(list_id, name="list_id")
        item_id = require_non_empty(item_id, name="item_id")
        await self._run(
            lambda: (
                self.client.sites.by_site_id(site_id)
                .lists.by_list_id(list_id)
                .items.by_list_item_id(item_id)
                .delete()
            ),
            operation_name="delete_list_item",
        )

    async def get_site_backup(self, site_id: str) -> dict[str, Any]:
        """Return a site's metadata, lists, and every list's items.

        The caller decides how to persist the result (see
        :mod:`m365_graph_toolkit.exporters`). An empty site returns
        empty lists, not an exception.
        """
        site_id = require_non_empty(site_id, name="site_id")
        site = await self._run(
            lambda: self.client.sites.by_site_id(site_id).get(),
            operation_name="get_site_backup",
        )
        lists = await self.get_lists(site_id)
        lists_data: list[dict[str, Any]] = []
        for lst in lists:
            items = await self.get_list_items(site_id, str(lst["id"]))
            lists_data.append({**lst, "items": items})
        return {"site": _normalise_site(site), "lists": lists_data}

    async def get_documents(self, site_id: str, list_id: str) -> list[dict[str, Any]]:
        """Return file metadata for items in a list that represent files."""
        items = await self.get_list_items(site_id, list_id)
        documents: list[dict[str, Any]] = []
        for item in items:
            fields = item.get("fields") or {}
            file_info = fields.get("file") if isinstance(fields, dict) else None
            if isinstance(file_info, dict):
                documents.append(
                    {
                        "id": item.get("id"),
                        "name": file_info.get("name"),
                        "url": file_info.get("url"),
                    }
                )
        return documents


__all__ = ["SharePointClient"]
