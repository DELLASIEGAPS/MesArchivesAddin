"""Explicit Graph beta importMime call; callers must supply an authorized token."""

from __future__ import annotations

import base64
import json
import socket
import urllib.error
import urllib.parse
import urllib.request


GRAPH_BETA = "https://graph.microsoft.com/beta"


class GraphImportError(RuntimeError):
    """Raised when Microsoft Graph rejects a MIME import."""


def import_mime(
    access_token: str,
    mailbox: str,
    mail_folder_id: str,
    mime_content: bytes,
) -> dict[str, object]:
    if not access_token or not mailbox or not mail_folder_id or not mime_content:
        raise ValueError("Token, mailbox, folder, and MIME content are required.")

    encoded_mailbox = urllib.parse.quote(mailbox, safe="@.")
    encoded_folder = urllib.parse.quote(mail_folder_id, safe="")
    url = f"{GRAPH_BETA}/users/{encoded_mailbox}/mailFolders/{encoded_folder}/importMime"
    request = urllib.request.Request(
        url,
        data=base64.b64encode(mime_content),
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "text/plain",
            "Accept": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            payload = response.read()
            if response.status != 201:
                raise GraphImportError(f"Unexpected Graph status {response.status}.")
            try:
                result = json.loads(payload)
            except (json.JSONDecodeError, UnicodeDecodeError) as error:
                raise GraphImportError(
                    "Graph returned an invalid response after import; do not retry without reconciliation."
                ) from error
            if not isinstance(result, dict) or not isinstance(result.get("id"), str) or not result["id"]:
                raise GraphImportError(
                    "Graph did not return a message ID; do not retry without reconciliation."
                )
            return result
    except urllib.error.HTTPError as error:
        raise GraphImportError(f"Graph import failed with HTTP {error.code}.") from error
    except (urllib.error.URLError, TimeoutError, socket.timeout) as error:
        raise GraphImportError(
            "Graph import outcome is unknown; do not retry without reconciliation."
        ) from error
