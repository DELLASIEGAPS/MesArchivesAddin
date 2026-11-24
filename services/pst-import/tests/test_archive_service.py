import io
import urllib.error
from unittest.mock import patch

import pytest

from archive_service import (
    ARCHIVE_OWNER_PROPERTY,
    ArchiveError,
    GraphRequestError,
    _is_owned_archive_folder,
    download_pst,
    ensure_archive_folder,
    folder_name,
    pst_name,
    unload_archive,
)


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


def test_new_archive_folder_is_created_with_ownership_marker():
    with (
        patch("archive_service.find_folder", return_value=None),
        patch("archive_service._graph", return_value={"id": "created"}) as graph,
    ):
        assert ensure_archive_folder("token", "2022") == "created"

    _, method, url, body = graph.call_args.args
    assert method == "POST"
    assert url.endswith("/me/mailFolders")
    assert body["displayName"] == "Archives 2022"
    assert body["singleValueExtendedProperties"] == [
        {
            "id": ARCHIVE_OWNER_PROPERTY,
            "value": "MesArchives PST import v1:2022",
        }
    ]


def test_unmarked_archive_folder_cannot_be_deleted():
    with (
        patch("archive_service.find_folder", return_value="existing"),
        patch(
            "archive_service._is_owned_archive_folder",
            return_value=False,
        ),
        patch("archive_service._graph") as graph,
    ):
        with pytest.raises(ArchiveError, match="marqueur MesArchives"):
            unload_archive("token", "2022")

    graph.assert_not_called()


def test_existing_unmarked_folder_is_not_reused_for_import():
    with (
        patch("archive_service.find_folder", return_value="existing"),
        patch("archive_service._is_owned_archive_folder", return_value=False),
    ):
        with pytest.raises(ArchiveError, match="ne sera ni utilisé ni supprimé"):
            ensure_archive_folder("token", "2022")


def test_archive_folder_with_marker_for_another_year_is_not_owned():
    with patch(
        "archive_service._graph",
        return_value={"value": "MesArchives PST import v1:2021"},
    ):
        assert not _is_owned_archive_folder("token", "folder-id", "2022")


def test_missing_ownership_property_is_not_owned():
    with patch(
        "archive_service._graph",
        side_effect=GraphRequestError("GET", 404),
    ):
        assert not _is_owned_archive_folder("token", "folder-id", "2022")


def test_ownership_check_propagates_non_404_graph_errors():
    with patch(
        "archive_service._graph",
        side_effect=GraphRequestError("GET", 503),
    ):
        with pytest.raises(GraphRequestError, match="HTTP 503"):
            _is_owned_archive_folder("token", "folder-id", "2022")


def test_marked_archive_folder_can_be_deleted():
    with (
        patch("archive_service.find_folder", return_value="owned"),
        patch("archive_service._is_owned_archive_folder", return_value=True),
        patch("archive_service._graph", return_value={}) as graph,
    ):
        assert unload_archive("token", "2022") == {
            "ok": True,
            "annee": 2022,
            "deleted": True,
        }

    assert graph.call_args.args[:2] == ("token", "DELETE")
    assert graph.call_args.args[2].endswith("/me/mailFolders/owned")


def test_missing_archive_folder_is_a_no_op():
    with patch("archive_service.find_folder", return_value=None), patch(
        "archive_service._graph"
    ) as graph:
        assert unload_archive("token", "2022") == {
            "ok": True,
            "annee": 2022,
            "deleted": False,
        }

    graph.assert_not_called()
