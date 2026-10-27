"""Authentication: one path, explicit environment variables, no secrets in logs.

v1 had two incompatible auth designs: ``AuthManager`` (client-secret
credential, no ``get_credential()``) and a never-existing
``GraphAuthClient`` that every API client imported. v2 has one
``AuthManager``. It reads exactly three variables, fails fast naming
only the missing *names* (never values), and never logs or prints a
token — the v1 ``__main__`` block printed the first 10 characters of
an access token, which this version does not do.

Credential imports are lazy (inside the factory methods) so importing
this package — and running the offline tests and the ``demo`` CLI
command — does not require ``azure-identity`` to be installed. The
live commands import it on first use and surface a clear error if it
is missing.

``load_dotenv()`` is deliberately *not* called here. The CLI calls it
once at startup; library callers who want ``.env`` support call it
themselves. Importing a library must not silently read files from the
working directory.
"""

from __future__ import annotations

import logging
import os
from collections.abc import Mapping, Sequence
from typing import Any

from .errors import AuthError

logger = logging.getLogger(__name__)

DEFAULT_SCOPES: tuple[str, ...] = ("https://graph.microsoft.com/.default",)

_ENV_VARS = ("MS_TENANT_ID", "MS_CLIENT_ID", "MS_CLIENT_SECRET")


class AuthManager:
    """Client-credentials auth for a single Microsoft Entra app registration."""

    def __init__(
        self,
        *,
        tenant_id: str | None = None,
        client_id: str | None = None,
        client_secret: str | None = None,
        scopes: Sequence[str] = DEFAULT_SCOPES,
        env: Mapping[str, str] | None = None,
    ) -> None:
        source: Mapping[str, str] = os.environ if env is None else env
        self.tenant_id = tenant_id or source.get("MS_TENANT_ID", "")
        self.client_id = client_id or source.get("MS_CLIENT_ID", "")
        self.client_secret = client_secret or source.get("MS_CLIENT_SECRET", "")
        self.scopes = list(scopes)

        missing = [
            name
            for name, value in (
                ("MS_TENANT_ID", self.tenant_id),
                ("MS_CLIENT_ID", self.client_id),
                ("MS_CLIENT_SECRET", self.client_secret),
            )
            if not value
        ]
        if missing:
            raise AuthError("Missing required environment variables: " + ", ".join(missing))
        if not self.scopes:
            raise AuthError("At least one scope is required")

        self._credential: object | None = None
        self._auth_provider: object | None = None

    def get_credential(self) -> object:
        """Return (and cache) the ``azure-identity`` credential."""
        if self._credential is None:
            try:
                from azure.identity import ClientSecretCredential
            except ImportError as exc:  # pragma: no cover - dependency missing
                raise AuthError(
                    "azure-identity is required for live Graph access; "
                    "install it with: pip install 'm365-graph-toolkit[graph]'"
                ) from exc
            self._credential = ClientSecretCredential(
                tenant_id=self.tenant_id,
                client_id=self.client_id,
                client_secret=self.client_secret,
            )
        return self._credential

    def get_auth_provider(self) -> object:
        """Return the ``msgraph-core`` auth provider for the SDK client."""
        if self._auth_provider is None:
            try:
                from msgraph_core.authentication import AzureIdentityAuthenticationProvider
            except ImportError as exc:  # pragma: no cover - dependency missing
                raise AuthError(
                    "msgraph-core is required for live Graph access; "
                    "install it with: pip install 'm365-graph-toolkit[graph]'"
                ) from exc
            self._auth_provider = AzureIdentityAuthenticationProvider(
                self.get_credential(),
                scopes=self.scopes,
            )
        return self._auth_provider

    def get_access_token(self) -> object:
        """Acquire a token and return the ``AccessToken`` object.

        Only the expiry is ever logged. The token value itself is never
        logged, printed, or returned as a string by this library.
        """
        credential: Any = self.get_credential()
        get_token = credential.get_token
        token = get_token(*self.scopes)
        expires_on = getattr(token, "expires_on", None)
        logger.info("Access token acquired (expires_on=%s)", expires_on)
        return token


__all__ = ["DEFAULT_SCOPES", "AuthManager"]
