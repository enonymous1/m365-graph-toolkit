"""Clients for the service packages."""

from __future__ import annotations

from .audit_logs import AuditLogsClient
from .base import BaseClient
from .groups import GroupsClient
from .onedrive import OneDriveClient
from .sharepoint import SharePointClient
from .users import UsersClient

__all__ = [
    "AuditLogsClient",
    "BaseClient",
    "GroupsClient",
    "OneDriveClient",
    "SharePointClient",
    "UsersClient",
]
