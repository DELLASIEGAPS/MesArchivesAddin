import json
import os
from unittest.mock import Mock, patch

import azure.functions as func
import msal


class FakeMSAL:
    def acquire_token_on_behalf_of(self, *_args, **_kwargs):
        return {}


with patch.dict(
   os.environ,
   {
    "TENANT_ID": "00000000-0000-0000-0000-000000000001",
    "CLIENT_ID": "00000000-0000-0000-0000-000000000002",
    "CLIENT_SECRET": "unit-test-value",
   },
):
   with patch.object(msal, "ConfidentialClientApplication", return_value=FakeMSAL()):
    import function_app


def request(body):
    return func.HttpRequest(
        method="POST",
        url="/api/archive",
        headers={"Authorization": "Bearer test-token"},
        params={},
        route_params={},
        body=json.dumps(body).encode(),
    )


def test_unload_requires_explicit_confirmation(monkeypatch):
    monkeypatch.setattr(function_app, "_claims", lambda _token: {"preferred_username": "test"})
    assert function_app.archive(
        request({"annee": 2022, "action": "unload"})
    ).status_code == 400
    assert function_app.archive(
        request({"annee": 2022, "action": "unload", "confirmDelete": False})
    ).status_code == 400


def test_confirmed_unload_reaches_archive_service(monkeypatch):
    monkeypatch.setattr(function_app, "_claims", lambda _token: {"preferred_username": "test"})
    monkeypatch.setattr(
        function_app.MSAL,
        "acquire_token_on_behalf_of",
        lambda *_args, **_kwargs: {"access_token": "delegated-token"},
    )
    unload = Mock(return_value={"ok": True, "annee": 2022, "deleted": True})
    monkeypatch.setattr(
        function_app,
        "unload_archive",
        unload,
    )

    response = function_app.archive(
        request({"annee": 2022, "action": "unload", "confirmDelete": True})
    )

    assert response.status_code == 200
    assert json.loads(response.get_body())["deleted"] is True
    unload.assert_called_once_with("delegated-token", "2022")
