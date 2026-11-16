import io
import urllib.error
from unittest.mock import patch

import pytest

from archive_service import ArchiveError, download_pst, folder_name, pst_name


def test_pst_and_folder_names():
    assert pst_name("2024") == "ADR.pst"
    assert pst_name("2021") == "ADR_2021.pst"
    assert folder_name("2022") == "Archives 2022"


def test_rejects_unknown_year():
    with pytest.raises(ValueError):
        pst_name("2020")


def test_download_reads_while_response_is_open(tmp_path):
    content = b"x" * (1024 * 1024 + 17)
    response = io.BytesIO(content)
    destination = tmp_path / "archive.pst"
    with patch("archive_service.urllib.request.urlopen", return_value=response) as urlopen:
        download_pst("token", "2021", destination)
    assert destination.read_bytes() == content
    assert response.closed
    request = urlopen.call_args.args[0]
    assert request.full_url.endswith("/me/drive/root:/Personnel/ADR_2021.pst:/content")
    assert request.get_header("Authorization") == "Bearer token"


def test_download_reports_graph_http_error(tmp_path):
    error = urllib.error.HTTPError("https://example.com", 404, "Not Found", {}, None)
    with patch("archive_service.urllib.request.urlopen", side_effect=error):
        with pytest.raises(ArchiveError, match="HTTP 404"):
            download_pst("token", "2024", tmp_path / "archive.pst")
