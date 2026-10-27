"""Typed errors for the M365 Graph Toolkit.

Callers should catch :class:`M365Error` for anything this library raises
on purpose. Everything else is a programming error or an unexpected
failure from a dependency, and it propagates.
"""

from __future__ import annotations


class M365Error(Exception):
    """Base class for errors raised by this library."""


class AuthError(M365Error):
    """Authentication configuration or credential failure."""


class GraphError(M365Error):
    """A Microsoft Graph request failed after retries (if any)."""

    def __init__(self, message: str, *, status_code: int | None = None) -> None:
        super().__init__(message)
        self.status_code = status_code


class NotFoundError(GraphError):
    """The requested Graph object does not exist (HTTP 404)."""


class ValidationError(M365Error):
    """A caller-supplied argument is missing or invalid."""


class ExportError(M365Error):
    """Writing an export file failed."""


__all__ = [
    "AuthError",
    "ExportError",
    "GraphError",
    "M365Error",
    "NotFoundError",
    "ValidationError",
]
