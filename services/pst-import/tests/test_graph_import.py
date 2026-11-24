import base64
import io
import json
import urllib.error
from unittest.mock import patch

import pytest

from graph_import import GraphImportError, import_mime


class Response:
    status = 201

    def __init__(self, content):
        self.content = content

    def __enter__(self):
        return self

    def __exit__(self, *args):
        pass

    def read(self):
        return self.content


def test_import_mime_encodes_payload_and_target():
    response = Response(json.dumps({"id": "message-id"}).encode())
    with patch("graph_import.urllib.request.urlopen", return_value=response) as urlopen:
        result = import_mime("token", "user@example.com", "folder/id", b"From: a@example.com\r\n")
    request = urlopen.call_args.args[0]
    assert request.full_url == (
        "https://graph.microsoft.com/beta/users/user@example.com/mailFolders/folder%2Fid/importMime"
    )
    assert request.get_method() == "POST"
    assert request.get_header("Authorization") == "Bearer token"
    assert base64.b64decode(request.data) == b"From: a@example.com\r\n"
    assert result["id"] == "message-id"


def test_import_rejects_bad_response_and_ambiguous_network_failures():
    for content in (b"{}", b"not-json"):
        with patch("graph_import.urllib.request.urlopen", return_value=Response(content)):
            with pytest.raises(GraphImportError, match="do not retry"):
                import_mime("token", "user@example.com", "folder", b"message")
    with patch("graph_import.urllib.request.urlopen", side_effect=TimeoutError()):
        with pytest.raises(GraphImportError, match="outcome is unknown"):
            import_mime("token", "user@example.com", "folder", b"message")


def test_import_reports_http_error_without_exposing_response():
    error = urllib.error.HTTPError("https://example.com", 403, "Forbidden", {}, io.BytesIO(b"private"))
    with patch("graph_import.urllib.request.urlopen", side_effect=error):
        with pytest.raises(GraphImportError, match="HTTP 403") as exc:
            import_mime("token", "user@example.com", "folder", b"message")
    assert "private" not in str(exc.value)
