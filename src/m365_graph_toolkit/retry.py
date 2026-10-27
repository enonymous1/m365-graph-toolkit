"""Retry and pagination helpers shared by every client.

Retry semantics (fixed from v1, whose loop ran ``max_retries + 1``
retries and whose retry predicate mixed ``and``/``or`` so an
``InternalServerError`` only retried while attempts remained by
accident):

* Exactly ``max_retries`` retries, then the last error is raised.
* Retryable: HTTP 429 / 502 / 503 / 504, plus timeouts and connection
  errors. Anything else (400, 401, 403, 404) fails immediately, so a
  permission problem is never disguised as a slow tenant.
* Exponential backoff capped at 30s, with jitter, honouring a
  ``Retry-After`` value when the error carries one.
* A run that hit retries is not silently presented as clean: callers
  log the retries, and :func:`with_retry` exposes the attempt count
  through the logger only — exporters never label throttled output as
  a complete measurement (see README, "Rate limits").

Pagination: Graph collection responses carry ``value`` plus an
``odata_next_link``. :func:`collect_pages` follows it until it is
absent. A page with ``value`` set to ``None`` contributes zero rows
instead of crashing, and a page object with no ``value`` attribute at
all raises :class:`~m365_graph_toolkit.errors.GraphError` — a shape
surprise should fail loudly, not be read as an empty tenant.
"""

from __future__ import annotations

import asyncio
import logging
import random
from collections.abc import Awaitable, Callable
from typing import Any, TypeVar

from .errors import GraphError

logger = logging.getLogger(__name__)

T = TypeVar("T")

_RETRYABLE_STATUSES = frozenset({429, 502, 503, 504})
_MAX_DELAY_SECONDS = 30.0


def _status_code(error: BaseException) -> int | None:
    """Best-effort HTTP status from an SDK / transport exception."""
    status = getattr(error, "response_status_code", None)
    if isinstance(status, int):
        return status
    response = getattr(error, "response", None)
    status = getattr(response, "status_code", None)
    if isinstance(status, int):
        return status
    status = getattr(response, "status", None)
    return status if isinstance(status, int) else None


def _retry_after_seconds(error: BaseException) -> float | None:
    """Best-effort ``Retry-After`` (seconds) from an exception, if any."""
    headers = getattr(error, "response_headers", None) or {}
    raw = headers.get("Retry-After") if isinstance(headers, dict) else None
    if raw is None:
        response = getattr(error, "response", None)
        headers = getattr(response, "headers", None) or {}
        raw = headers.get("Retry-After") if hasattr(headers, "get") else None
    try:
        value = float(raw) if raw is not None else None
    except (TypeError, ValueError):
        return None
    return value if value is not None and value >= 0 else None


def is_retryable(error: BaseException) -> bool:
    """Whether ``error`` is worth retrying (see module docstring)."""
    status = _status_code(error)
    if status is not None:
        return status in _RETRYABLE_STATUSES
    name = type(error).__name__
    if name in {
        "TimeoutError",
        "ConnectionError",
        "ClientConnectionError",
        "ServerDisconnectedError",
    }:
        return True
    message = str(error)
    return any(
        token in message
        for token in ("TooManyRequests", "ServiceUnavailable", "BadGateway", "GatewayTimeout")
    )


async def with_retry(
    operation: Callable[[], Awaitable[T]],
    *,
    max_retries: int = 3,
    base_delay: float = 1.0,
    operation_name: str = "Graph request",
) -> T:
    """Run ``operation`` with bounded retries (see module docstring)."""
    if max_retries < 0:
        raise ValueError("max_retries must be >= 0")
    delay = base_delay
    last_error: BaseException | None = None
    for attempt in range(max_retries + 1):
        try:
            return await operation()
        except Exception as error:  # noqa: BLE001 - predicate decides
            last_error = error
            if attempt >= max_retries or not is_retryable(error):
                raise
            retry_after = _retry_after_seconds(error)
            wait = retry_after if retry_after is not None else delay
            wait = min(wait, _MAX_DELAY_SECONDS) + random.uniform(0, 0.5)
            logger.warning(
                "%s failed (%s); retrying in %.1fs (%d/%d)",
                operation_name,
                error,
                wait,
                attempt + 1,
                max_retries,
            )
            await asyncio.sleep(wait)
            delay = min(delay * 2, _MAX_DELAY_SECONDS)
    # Unreachable: the loop either returns or raises.
    assert last_error is not None  # noqa: S101
    raise last_error


async def collect_pages(
    first_page: Callable[[], Awaitable[Any]],
    next_page: Callable[[str], Awaitable[Any]],
) -> list[Any]:
    """Collect every item from a paged Graph collection response."""
    items: list[Any] = []
    page = await first_page()
    while page is not None:
        if not hasattr(page, "value"):
            raise GraphError("Graph collection response has no 'value' attribute")
        values = page.value
        if values:
            items.extend(values)
        next_link = getattr(page, "odata_next_link", None)
        if not next_link:
            break
        page = await next_page(next_link)
    return items


__all__ = ["collect_pages", "is_retryable", "with_retry"]
