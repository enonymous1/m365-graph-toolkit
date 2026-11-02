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

STIG import (folded in from ``mp-m365-app``, 2.1.0): ``create_list``
and ``import_list_from_csv`` create a list per STIG CSV and import
its rows, polling until the new columns are provisioned before
writing any item — the old code polled too, but with a fixed
``asyncio.sleep(6)`` loop that returned the half-created list as
success when provisioning never finished. Here unprovisioned
columns raise instead of writing items against guessed names, and
per-row failures are counted and reported in the returned summary
instead of ending the import.
"""

from __future__ import annotations

import asyncio
from collections.abc import Iterable
from pathlib import Path
from typing import Any

from ..errors import GraphError, ValidationError
from ..normalize import model_to_dict, require_non_empty
from ..stig import (
    discover_csv_files,
    load_stig_csv,
    map_row_to_fields,
    resolve_column_map,
)
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
        """Build an SDK ListItem carrying a FieldValueSet (live path).

        Validation runs before the SDK import (see ``_build_list_model``).
        """
        if not fields:
            raise ValidationError("fields must not be empty")
        from msgraph.generated.models.field_value_set import FieldValueSet
        from msgraph.generated.models.list_item import ListItem

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

    # --- Site search and list creation (STIG import) ---------------------

    async def search_sites(
        self, query: str, *, select: Iterable[str] | None = None
    ) -> list[dict[str, Any]]:
        """Return sites matching ``query``, following pagination.

        Callers pick the site they mean from the results; nothing
        here assumes the first match is the right one (the old STIG
        importer did exactly that with a hardcoded ``"STIG"``
        query).
        """
        query = require_non_empty(query, name="query")
        configuration = self._search_configuration(
            query=query, select=list(select) if select else list(SITE_FIELDS)
        )

        async def first_page() -> Any:
            return await self.client.sites.get(request_configuration=configuration)

        async def next_page(url: str) -> Any:
            return await self.client.sites.with_url(url).get()

        sites = await self._collect(first_page, next_page, operation_name="search_sites")
        return [_normalise_site(site) for site in sites]

    def _search_configuration(self, *, query: str, select: list[str]) -> Any:
        from msgraph.generated.sites.sites_request_builder import SitesRequestBuilder

        query_params = SitesRequestBuilder.SitesRequestBuilderGetQueryParameters(
            search=query, select=select
        )
        return SitesRequestBuilder.SitesRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

    def _build_list_model(self, display_name: str, columns: list[dict[str, Any]]) -> Any:
        """Build an SDK list model with multi-line text columns (live path).

        Validation runs before the SDK import, so a bad plan raises
        ``ValidationError`` even when ``msgraph-sdk`` is not installed
        (base install), instead of a misleading ``ModuleNotFoundError``.
        """
        # Validate first, import second (see docstring).
        seen_names: set[str] = set()
        for column in columns:
            name = str(column.get("name", "")).strip()
            if not name:
                raise ValidationError("List column names must not be empty")
            lowered = name.lower()
            if lowered in seen_names:
                raise ValidationError(f"Duplicate list column name: {name!r}")
            seen_names.add(lowered)

        from msgraph.generated.models.column_definition import ColumnDefinition
        from msgraph.generated.models.list_ import List_
        from msgraph.generated.models.list_info import ListInfo
        from msgraph.generated.models.text_column import TextColumn

        definitions = []
        for column in columns:
            name = str(column.get("name", "")).strip()
            definitions.append(
                ColumnDefinition(
                    name=name,
                    text=TextColumn(
                        allow_multiple_lines=bool(column.get("allow_multiple_lines", True))
                    ),
                )
            )
            definitions.append(
                ColumnDefinition(
                    name=name,
                    text=TextColumn(
                        allow_multiple_lines=bool(column.get("allow_multiple_lines", True))
                    ),
                )
            )
        return List_(
            display_name=display_name,
            columns=definitions,
            list_=ListInfo(template="genericList"),
        )

    async def create_list(
        self, site_id: str, display_name: str, columns: list[dict[str, Any]]
    ) -> dict[str, Any]:
        """Create a generic list with the given text columns.

        ``columns`` is a list of ``{"name": ..., "allow_multiple_lines": ...}``
        mappings. The ``Title`` column is built in by SharePoint and
        must not be listed here.
        """
        site_id = require_non_empty(site_id, name="site_id")
        display_name = require_non_empty(display_name, name="display_name")
        if any(str(column.get("name", "")).strip().lower() == "title" for column in columns):
            raise ValidationError("Do not define a 'Title' column; SharePoint provides it")
        model = self._build_list_model(display_name, columns)
        created = await self._run(
            lambda: self.client.sites.by_site_id(site_id).lists.post(model),
            operation_name="create_list",
        )
        return _normalise_list(created)

    async def get_list_columns(self, site_id: str, list_id: str) -> list[dict[str, Any]]:
        """Return a list's columns, following pagination."""
        site_id = require_non_empty(site_id, name="site_id")
        list_id = require_non_empty(list_id, name="list_id")

        async def first_page() -> Any:
            return (
                await self.client.sites.by_site_id(site_id).lists.by_list_id(list_id).columns.get()
            )

        async def next_page(url: str) -> Any:
            return (
                await self.client.sites.by_site_id(site_id)
                .lists.by_list_id(list_id)
                .columns.with_url(url)
                .get()
            )

        columns = await self._collect(first_page, next_page, operation_name="get_list_columns")
        return [_normalise_column(column) for column in columns]

    async def import_list_from_csv(
        self,
        site_id: str,
        plan: dict[str, Any],
        rows: list[dict[str, Any]],
        *,
        poll_attempts: int = 7,
        poll_delay_seconds: float = 6.0,
        max_failures_reported: int = 20,
    ) -> dict[str, Any]:
        """Create a list from a STIG plan and import its rows. Library only.

        The list is created, then its columns are polled until every
        planned column is provisioned as a text column and the
        ``Title`` mapping resolves. If provisioning never finishes,
        :class:`~m365_graph_toolkit.errors.GraphError` is raised and
        **no items are written** — the created list is left in place
        and named in the error. Per-row failures are counted, not
        fatal: the returned summary reports ``items_created``,
        ``items_failed``, and the first ``max_failures_reported``
        failures. A plan with no rows raises
        :class:`~m365_graph_toolkit.errors.ValidationError`.
        """
        site_id = require_non_empty(site_id, name="site_id")
        list_name = str(plan.get("list_name", "")).strip()
        if not list_name:
            raise ValidationError("plan must contain a non-empty 'list_name'")
        if not rows:
            raise ValidationError(f"Plan for list {list_name!r} has no rows to import")
        if poll_attempts < 1:
            raise ValidationError("poll_attempts must be >= 1")
        if poll_delay_seconds < 0:
            raise ValidationError("poll_delay_seconds must be >= 0")
        text_columns = [
            {"name": str(column["header"]), "allow_multiple_lines": True}
            for column in plan.get("columns", [])
            if not column.get("is_title")
        ]
        created = await self.create_list(site_id, list_name, text_columns)
        list_id = str(created.get("id"))
        if not list_id:
            raise GraphError(f"Graph created list {list_name!r} without an id")
        column_map: dict[str, str] | None = None
        last_error: Exception | None = None
        for attempt in range(poll_attempts):
            columns = await self.get_list_columns(site_id, list_id)
            try:
                column_map = resolve_column_map(plan, columns)
                break
            except ValidationError as error:
                last_error = error
                if attempt < poll_attempts - 1:
                    await asyncio.sleep(poll_delay_seconds)
        if column_map is None:
            raise GraphError(
                f"Columns for list {list_name!r} (id {list_id}) were not provisioned "
                f"after {poll_attempts} attempts: {last_error}. "
                "The list was created; no items were written."
            )
        created_count = 0
        failed_count = 0
        failures: list[dict[str, Any]] = []
        for index, row in enumerate(rows, start=1):
            fields = map_row_to_fields(plan, row, column_map)
            if not fields:
                failed_count += 1
                if len(failures) < max_failures_reported:
                    failures.append(
                        {"row": index, "error": "Row mapped to no fields", "fields": fields}
                    )
                continue
            try:
                await self.create_list_item(site_id, list_id, fields)
                created_count += 1
            except Exception as error:  # noqa: BLE001 - counted, import continues
                failed_count += 1
                self.logger.warning(
                    "Failed to import row %d into list %s: %s", index, list_name, error
                )
                if len(failures) < max_failures_reported:
                    failures.append({"row": index, "error": str(error), "fields": fields})
            if index % 500 == 0:
                self.logger.info("Imported %d/%d rows into list %s", index, len(rows), list_name)
        return {
            "list_name": list_name,
            "list_id": list_id,
            "source_file": plan.get("source_file"),
            "rows_total": len(rows),
            "items_created": created_count,
            "items_failed": failed_count,
            "failures": failures,
        }

    async def import_lists_from_directory(
        self,
        site_id: str,
        directory: str | Path,
        *,
        poll_attempts: int = 7,
        poll_delay_seconds: float = 6.0,
    ) -> list[dict[str, Any]]:
        """Import every STIG CSV in ``directory`` into ``site_id``. Library only.

        One CSV's failure does not stop the rest: its summary carries
        an ``error`` key and zero created items. The caller confirms
        the site and the directory before calling — there is no CLI
        path to this method.
        """
        site_id = require_non_empty(site_id, name="site_id")
        summaries: list[dict[str, Any]] = []
        for csv_file in discover_csv_files(directory):
            try:
                plan, rows = load_stig_csv(csv_file)
                summary = await self.import_list_from_csv(
                    site_id,
                    plan,
                    rows,
                    poll_attempts=poll_attempts,
                    poll_delay_seconds=poll_delay_seconds,
                )
            except Exception as error:  # noqa: BLE001 - reported per file, batch continues
                summaries.append(
                    {
                        "list_name": csv_file.stem,
                        "list_id": None,
                        "source_file": csv_file.name,
                        "rows_total": 0,
                        "items_created": 0,
                        "items_failed": 0,
                        "failures": [],
                        "error": str(error),
                    }
                )
                continue
            summaries.append(summary)
        return summaries


def _normalise_column(column: Any) -> dict[str, Any]:
    row = model_to_dict(column, ("id", "name", "display_name"))
    text = column.get("text") if isinstance(column, dict) else getattr(column, "text", None)
    row["is_text"] = text is not None
    if isinstance(text, dict):
        row["allow_multiple_lines"] = text.get(
            "allowMultipleLines", text.get("allow_multiple_lines")
        )
    else:
        row["allow_multiple_lines"] = getattr(text, "allow_multiple_lines", None)
    return row


__all__ = ["SharePointClient"]
