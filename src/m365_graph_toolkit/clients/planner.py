"""Planner client: plans, tasks (with details and buckets), and per-user tasks.

Folded in from the February/May 2026 ``mp-m365-app`` tree. Fixes over
that code:

* No hardcoded paths, no ``sys.path`` hacks, no per-module
  ``load_dotenv()``. The client takes an already-built Graph client
  (see :mod:`m365_graph_toolkit.clients.base`).
* Pagination is real. The old ``get_user_planner_tasks`` returned
  only the first page, ``get_all_planner_plans`` never paged the
  groups list, and both retried with an unbounded ``max_retries + 1``
  loop that retried 403s. Every list method here pages through
  ``odata_next_link`` and retries only 429/502/503/504 via the shared
  :func:`~m365_graph_toolkit.retry.with_retry`.
* One ``$select``/``$expand`` shape per call instead of fetching
  ``select=["*"]`` (the old plans call) or nothing at all.
* Per-user task fetching is bounded. The old code gathered one
  unbounded task per user and aborted the whole run on the first
  failure; :meth:`PlannerClient.get_tasks_for_users` caps
  concurrency with a semaphore and logs-and-continues per user.
* Tasks are returned as plain ``dict`` rows with the bucket name,
  assignee ids, applied-category labels, and checklist counts
  resolved. The old code wrote those through ``pandas`` straight
  into per-group directories under a hardcoded home path; callers
  here decide where rows go (see :mod:`m365_graph_toolkit.exporters`).
"""

from __future__ import annotations

import asyncio
from collections.abc import Iterable
from typing import Any

from ..errors import ValidationError
from ..normalize import model_to_dict, require_non_empty, scalar_value
from .base import BaseClient

PLAN_FIELDS = (
    "id",
    "title",
    "created_date_time",
)

TASK_FIELDS = (
    "id",
    "title",
    "bucket_id",
    "percent_complete",
    "priority",
    "created_date_time",
    "start_date_time",
    "due_date_time",
    "completed_date_time",
)

_TASK_SELECT = [
    "id",
    "title",
    "percentComplete",
    "createdDateTime",
    "dueDateTime",
    "assignments",
    "bucketId",
    "priority",
    "appliedCategories",
    "createdBy",
    "startDateTime",
    "completedDateTime",
    "completedBy",
]


def _mapping(obj: Any) -> dict[str, Any]:
    """Read an SDK ``additional_data`` mapping off an object or dict."""
    if isinstance(obj, dict):
        return obj
    additional = getattr(obj, "additional_data", None)
    return dict(additional) if isinstance(additional, dict) else {}


def _field(obj: Any, name: str) -> Any:
    """Read ``name`` off an SDK object or a plain dict."""
    if isinstance(obj, dict):
        return obj.get(name)
    return getattr(obj, name, None)


def _identity_id(identity_set: Any) -> str | None:
    """Return the ``user.id`` of an IdentitySet, or ``None``."""
    if identity_set is None:
        return None
    user = _field(identity_set, "user")
    if user is None:
        return None
    user_id = _field(user, "id")
    return str(user_id) if user_id is not None else None


def _normalise_plan(plan: Any, *, group_name: str | None = None) -> dict[str, Any]:
    row = model_to_dict(plan, PLAN_FIELDS)
    row["created_by_user_id"] = _identity_id(_field(plan, "created_by"))
    container = _field(plan, "container")
    row["container_id"] = _field(container, "container_id") if container is not None else None
    row["container_type"] = _field(container, "type") if container is not None else None
    row["container_url"] = _field(container, "url") if container is not None else None
    row["group_name"] = group_name
    return row


def _normalise_task(
    task: Any, *, bucket_name: str | None = None, plan_id: str | None = None
) -> dict[str, Any]:
    row = model_to_dict(task, TASK_FIELDS)
    row["plan_id"] = plan_id
    row["bucket_name"] = bucket_name
    row["created_by_user_id"] = _identity_id(_field(task, "created_by"))
    row["completed_by_user_id"] = _identity_id(_field(task, "completed_by"))
    row["assigned_to"] = sorted(_mapping(_field(task, "assignments")).keys())
    categories = _mapping(_field(task, "applied_categories"))
    row["labels"] = sorted(key for key, value in categories.items() if value)
    details = _field(task, "details")
    row["description"] = _field(details, "description") if details is not None else None
    checklist = _mapping(_field(details, "checklist")) if details is not None else {}
    items: list[dict[str, Any]] = []
    for item in checklist.values():
        items.append(
            {
                "title": _field(item, "title"),
                "is_checked": bool(_field(item, "is_checked")),
            }
        )
    row["checklist_items"] = items
    row["checklist_total"] = len(items)
    row["checklist_completed"] = sum(1 for item in items if item["is_checked"])
    row["due_date_time"] = scalar_value(row["due_date_time"])
    return row


class PlannerClient(BaseClient):
    """Microsoft Graph Planner operations."""

    async def get_user_tasks(
        self,
        user_id: str,
        *,
        select: Iterable[str] | None = None,
        top: int | None = None,
    ) -> list[dict[str, Any]]:
        """Return a user's Planner tasks, following pagination."""
        user_id = require_non_empty(user_id, name="user_id")
        if top is not None and top < 1:
            raise ValidationError("top must be >= 1")
        configuration = self._tasks_configuration(
            select=list(select) if select else list(_TASK_SELECT), top=top
        )

        async def first_page() -> Any:
            return await self.client.users.by_user_id(user_id).planner.tasks.get(
                request_configuration=configuration
            )

        async def next_page(url: str) -> Any:
            return await self.client.users.by_user_id(user_id).planner.tasks.with_url(url).get()

        tasks = await self._collect(first_page, next_page, operation_name="get_user_tasks")
        return [_normalise_task(task) for task in tasks]

    def _tasks_configuration(self, *, select: list[str], top: int | None) -> Any:
        from msgraph.generated.users.item.planner.tasks.tasks_request_builder import (
            TasksRequestBuilder,
        )

        query_params = TasksRequestBuilder.TasksRequestBuilderGetQueryParameters(
            select=select, top=top
        )
        return TasksRequestBuilder.TasksRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

    async def get_tasks_for_users(
        self,
        users: list[dict[str, Any]],
        *,
        max_concurrent: int = 5,
    ) -> list[dict[str, Any]]:
        """Fetch Planner tasks for many users concurrently, in input order.

        A user whose fetch fails contributes zero rows and a warning;
        failures never abort the whole batch. This replaces the old
        code's unbounded gather that raised on the first 403.
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
                    tasks = await self.get_user_tasks(str(user_id))
                except Exception as error:  # noqa: BLE001 - logged, batch continues
                    self.logger.warning("Planner fetch failed for user %s: %s", user_id, error)
                    return []
            for task in tasks:
                task["user_id"] = user.get("id")
                task["user_display_name"] = user.get("display_name")
            return tasks

        results = await asyncio.gather(*(fetch_one(user) for user in users))
        return [row for batch in results for row in batch]

    async def get_group_plans(
        self, group_id: str, *, top: int | None = None
    ) -> list[dict[str, Any]]:
        """Return the Planner plans owned by a group, following pagination."""
        group_id = require_non_empty(group_id, name="group_id")
        if top is not None and top < 1:
            raise ValidationError("top must be >= 1")

        async def first_page() -> Any:
            return await self.client.groups.by_group_id(group_id).planner.plans.get()

        async def next_page(url: str) -> Any:
            return await self.client.groups.by_group_id(group_id).planner.plans.with_url(url).get()

        plans = await self._collect(first_page, next_page, operation_name="get_group_plans")
        return [_normalise_plan(plan) for plan in plans]

    async def get_plan_buckets(self, plan_id: str) -> dict[str, str]:
        """Return a plan's bucket id -> name mapping."""
        plan_id = require_non_empty(plan_id, name="plan_id")

        async def first_page() -> Any:
            return await self.client.planner.plans.by_planner_plan_id(plan_id).buckets.get()

        async def next_page(url: str) -> Any:
            return (
                await self.client.planner.plans.by_planner_plan_id(plan_id)
                .buckets.with_url(url)
                .get()
            )

        buckets = await self._collect(first_page, next_page, operation_name="get_plan_buckets")
        return {
            str(_field(bucket, "id")): str(_field(bucket, "name"))
            for bucket in buckets
            if _field(bucket, "id") is not None
        }

    async def get_plan_tasks(
        self, plan_id: str, *, with_details: bool = True
    ) -> list[dict[str, Any]]:
        """Return a plan's tasks, with bucket names and details resolved."""
        plan_id = require_non_empty(plan_id, name="plan_id")
        bucket_map = await self.get_plan_buckets(plan_id)
        configuration = self._plan_tasks_configuration(expand_details=with_details)

        async def first_page() -> Any:
            return await self.client.planner.plans.by_planner_plan_id(plan_id).tasks.get(
                request_configuration=configuration
            )

        async def next_page(url: str) -> Any:
            return (
                await self.client.planner.plans.by_planner_plan_id(plan_id)
                .tasks.with_url(url)
                .get()
            )

        tasks = await self._collect(first_page, next_page, operation_name="get_plan_tasks")
        return [
            _normalise_task(
                task,
                bucket_name=bucket_map.get(str(_field(task, "bucket_id"))),
                plan_id=plan_id,
            )
            for task in tasks
        ]

    def _plan_tasks_configuration(self, *, expand_details: bool) -> Any:
        from msgraph.generated.planner.plans.item.tasks.tasks_request_builder import (
            TasksRequestBuilder,
        )

        query_params = TasksRequestBuilder.TasksRequestBuilderGetQueryParameters(
            select=list(_TASK_SELECT),
            expand=["details"] if expand_details else None,
        )
        return TasksRequestBuilder.TasksRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

    async def get_all_plans(
        self,
        groups: list[dict[str, Any]],
        *,
        max_concurrent: int = 5,
    ) -> list[dict[str, Any]]:
        """Return every plan owned by ``groups`` (each row carries its group name).

        ``groups`` is the output of ``GroupsClient.get_groups``. A group
        whose plans cannot be read contributes zero rows and a warning.
        """
        if max_concurrent < 1:
            raise ValidationError("max_concurrent must be >= 1")
        semaphore = asyncio.Semaphore(max_concurrent)

        async def fetch_one(group: dict[str, Any]) -> list[dict[str, Any]]:
            group_id = group.get("id")
            if not group_id:
                return []
            async with semaphore:
                try:
                    plans = await self.get_group_plans(str(group_id))
                except Exception as error:  # noqa: BLE001 - logged, batch continues
                    self.logger.warning(
                        "Planner plans fetch failed for group %s: %s", group_id, error
                    )
                    return []
            for plan in plans:
                plan["group_name"] = group.get("display_name")
            return plans

        results = await asyncio.gather(*(fetch_one(group) for group in groups))
        return [row for batch in results for row in batch]


__all__ = ["PLAN_FIELDS", "TASK_FIELDS", "PlannerClient"]
