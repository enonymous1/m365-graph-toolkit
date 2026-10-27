from __future__ import annotations

import pytest

from m365_graph_toolkit.auth import AuthManager
from m365_graph_toolkit.errors import AuthError


def test_missing_env_vars_raise_naming_only_the_names() -> None:
    with pytest.raises(AuthError) as exc_info:
        AuthManager(env={})
    message = str(exc_info.value)
    assert "MS_TENANT_ID" in message
    assert "MS_CLIENT_ID" in message
    assert "MS_CLIENT_SECRET" in message
    # No values are ever echoed into the error.
    assert "hunter2" not in message


def test_partial_env_names_only_the_missing_ones() -> None:
    with pytest.raises(AuthError) as exc_info:
        AuthManager(env={"MS_TENANT_ID": "t", "MS_CLIENT_ID": "c"})
    assert "MS_CLIENT_SECRET" in str(exc_info.value)
    assert "MS_TENANT_ID" not in str(exc_info.value)


def test_explicit_arguments_override_env() -> None:
    auth = AuthManager(tenant_id="t", client_id="c", client_secret="s", env={})
    assert auth.tenant_id == "t"
    assert auth.scopes == ["https://graph.microsoft.com/.default"]


def test_empty_scopes_rejected() -> None:
    with pytest.raises(AuthError):
        AuthManager(tenant_id="t", client_id="c", client_secret="s", scopes=[])


def test_get_credential_caches_one_instance(monkeypatch: pytest.MonkeyPatch) -> None:
    import sys
    import types

    created: list[dict] = []

    class FakeCredential:
        def __init__(self, **kwargs):  # noqa: ANN003
            created.append(kwargs)

    azure_pkg = types.ModuleType("azure")
    identity_mod = types.ModuleType("azure.identity")
    identity_mod.ClientSecretCredential = FakeCredential  # type: ignore[attr-defined]
    monkeypatch.setitem(sys.modules, "azure", azure_pkg)
    monkeypatch.setitem(sys.modules, "azure.identity", identity_mod)

    auth = AuthManager(tenant_id="t", client_id="c", client_secret="s", env={})
    first = auth.get_credential()
    second = auth.get_credential()
    assert first is second
    assert len(created) == 1
    assert created[0]["tenant_id"] == "t"


def test_get_access_token_logs_expiry_not_value(monkeypatch: pytest.MonkeyPatch, caplog) -> None:
    import sys
    import types

    class FakeToken:
        expires_on = 1234567890

    class FakeCredential:
        def __init__(self, **kwargs):  # noqa: ANN003
            pass

        def get_token(self, *scopes):  # noqa: ANN002
            assert scopes == ("https://graph.microsoft.com/.default",)
            return FakeToken()

    azure_pkg = types.ModuleType("azure")
    identity_mod = types.ModuleType("azure.identity")
    identity_mod.ClientSecretCredential = FakeCredential  # type: ignore[attr-defined]
    monkeypatch.setitem(sys.modules, "azure", azure_pkg)
    monkeypatch.setitem(sys.modules, "azure.identity", identity_mod)

    auth = AuthManager(tenant_id="t", client_id="c", client_secret="s", env={})
    with caplog.at_level("INFO", logger="m365_graph_toolkit.auth"):
        token = auth.get_access_token()
    assert isinstance(token, FakeToken)
    assert "1234567890" in caplog.text
