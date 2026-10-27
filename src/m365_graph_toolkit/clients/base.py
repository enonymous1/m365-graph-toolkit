"""Base client: one GraphServiceClient, shared pagination and retry.

Every v2 client is constructed with an already-built Graph client
(real or a test fake) plus optional retry settings. Construction never
touches the network and never reads the environment — the CLI builds
an :class:`~m365_graph_toolkit.auth.AuthManager` and passes its
client in via :meth:`BaseClient.from_auth`. That separation is what
makes the clients unit-testable without credentials.
"""

from __future__ import annotations

import logging
from collections.abc import Awaitable, Callable
from typing import Any

from ..auth import AuthManager
from ..errors import GraphError, NotFoundError
from ..retry import collect_pages, with_retry


class BaseClient:
    """Shared plumbing for the service clients."""

    def __init__(
        self,
        graph_client: Any,
        *,
        max_retries: int = 3,
        base_delay: float = 1.0,
    ) -> None:
        self.client = graph_client
        self.max_retries = max_retries
        self.base_delay = base_delay
        self.logger = logging.getLogger(self.__class__.__name__)

    @classmethod
    def from_auth(cls, auth: AuthManager, **kwargs: Any) -> BaseClient:
        """Build a client from an :class:`AuthManager` (live path only)."""
        try:
            from msgraph import GraphServiceClient
        except ImportError as exc:  # pragma: no cover - dependency missing
            from ..errors import AuthError

            raise AuthError(
                "msgraph-sdk is required for live Graph access; "
                "install it with: pip install 'm365-graph-toolkit[graph]'"
            ) from exc
        graph_client = GraphServiceClient(
            credentials=auth.get_credential(),  # type: ignore[arg-type]
            scopes=auth.scopes,
        )
        return cls(graph_client, **kwargs)

    async def _run(self, operation: Callable[[], Awaitable[Any]], *, operation_name: str) -> Any:
        """Run one SDK call with retry, translating 404s to NotFoundError."""
        try:
            return await with_retry(
                operation,
                max_retries=self.max_retries,
                base_delay=self.base_delay,
                operation_name=operation_name,
            )
        except Exception as error:
            status = getattr(error, "response_status_code", None)
            if status == 404 or "Request_ResourceNotFound" in str(error):
                raise NotFoundError(str(error), status_code=404) from error
            if isinstance(error, GraphError):
                raise
            raise GraphError(f"{operation_name} failed: {error}") from error

    async def _collect(
        self,
        first_page: Callable[[], Awaitable[Any]],
        next_page: Callable[[str], Awaitable[Any]],
        *,
        operation_name: str,
    ) -> list[Any]:
        """Collect a paged collection, retrying each page fetch."""
        return await collect_pages(
            lambda: self._run(first_page, operation_name=operation_name),
            lambda url: self._run(lambda: next_page(url), operation_name=operation_name),
        )


__all__ = ["BaseClient"]
