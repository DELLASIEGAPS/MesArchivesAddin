"""Azure Function : POST /api/archive {annee, action: load|unload}, jeton SSO Office en Bearer."""

import json
import os

import azure.functions as func
import jwt
import msal

from archive_service import ArchiveError, load_archive, unload_archive, YEARS

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

TENANT = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]  # paramètre d'application / Key Vault, jamais dans le dépôt
ORIGIN = "https://dellasiegaps.github.io"
SCOPES = ["https://graph.microsoft.com/Files.Read",
          "https://graph.microsoft.com/Mail.ReadWrite",
          "https://graph.microsoft.com/Mail-Advanced.ReadWrite"]
JWKS = jwt.PyJWKClient(f"https://login.microsoftonline.com/{TENANT}/discovery/v2.0/keys")
MSAL = msal.ConfidentialClientApplication(
    CLIENT_ID, client_credential=CLIENT_SECRET,
    authority=f"https://login.microsoftonline.com/{TENANT}")


def _reply(status: int, body: dict) -> func.HttpResponse:
    return func.HttpResponse(json.dumps(body), status_code=status, mimetype="application/json",
                             headers={"Access-Control-Allow-Origin": ORIGIN,
                                      "Access-Control-Allow-Headers": "Authorization, Content-Type"})


def _claims(token: str) -> dict:
    key = JWKS.get_signing_key_from_jwt(token).key
    claims = jwt.decode(token, key, algorithms=["RS256"],
                        audience=[f"api://{CLIENT_ID}", f"api://dellasiegaps.github.io/{CLIENT_ID}"],
                        issuer=f"https://login.microsoftonline.com/{TENANT}/v2.0")
    if "access_as_user" not in claims.get("scp", "").split():
        raise jwt.InvalidTokenError("scope manquant")
    return claims


@app.route(route="archive", methods=["POST", "OPTIONS"])
def archive(req: func.HttpRequest) -> func.HttpResponse:
    if req.method == "OPTIONS":
        return _reply(204, {})
    auth = req.headers.get("Authorization", "")
    if not auth.startswith("Bearer "):
        return _reply(401, {"ok": False, "error": "jeton manquant"})
    try:
        claims = _claims(auth[7:])
    except jwt.PyJWTError:
        return _reply(401, {"ok": False, "error": "jeton invalide"})
    try:
        body = req.get_json()
        year, action = str(body["annee"]), body["action"]
    except (ValueError, KeyError):
        return _reply(400, {"ok": False, "error": "corps invalide"})
    if year not in YEARS or action not in ("load", "unload"):
        return _reply(400, {"ok": False, "error": "annee ou action invalide"})
    obo = MSAL.acquire_token_on_behalf_of(auth[7:], SCOPES)
    if "access_token" not in obo:
        return _reply(403, {"ok": False, "error": "consentement ou OBO refusé"})
    try:
        mailbox = claims.get("preferred_username") or claims["upn"]
        result = (load_archive(obo["access_token"], mailbox, year) if action == "load"
                  else unload_archive(obo["access_token"], year))
        return _reply(200, result)
    except ArchiveError as error:
        return _reply(500, {"ok": False, "error": str(error)})
