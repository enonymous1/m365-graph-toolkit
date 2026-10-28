"""Clients for the service packages."""

from __future__ import annotations

from .audit_logs import AuditLogsClient
from .base import BaseClient
from .calendar import CalendarClient
from .groups import GroupsClient
from .mail import MailClient
from .onedrive import OneDriveClient
from .planner import PlannerClient
from .sharepoint import SharePointClient
from .teams_messages import TeamsMessagesClient
from .users import UsersClient

__all__ = [
    "AuditLogsClient",
    "BaseClient",
    "CalendarClient",
    "GroupsClient",
    "MailClient",
    "OneDriveClient",
    "PlannerClient",
    "SharePointClient",
    "TeamsMessagesClient",
    "UsersClient",
]
