"""Charger/décharger une archive PST dans la boîte de l'utilisateur via Microsoft Graph."""

from __future__ import annotations

import json
import tempfile
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Callable

from graph_import import GraphImportError, import_mime
from pst_to_mime import iter_pst_messages

GRAPH = "https://graph.microsoft.com/v1.0"
ONEDRIVE_FOLDER = "Personnel"
YEARS = ("2021", "2022", "2023", "2024")
ARCHIVE_OWNER_PROPERTY = (
    "String {014352da-084d-44ca-9589-65f15fe9af4e} Name MesArchivesArchiveYear"
)


class ArchiveError(RuntimeError):
    pass


class GraphRequestError(ArchiveError):
    def __init__(self, method: str, status_code: int):
        super().__init__(f"Graph {method} a échoué (HTTP {status_code}).")
        self.status_code = status_code


def pst_name(year: str) -> str:
    if year not in YEARS:
        raise ValueError("Année hors 2021-2024.")
    return "ADR.pst" if year == "2024" else f"ADR_{year}.pst"


def folder_name(year: str) -> str:
    pst_name(year)
    return f"Archives {year}"


def _graph(token: str, method: str, url: str, body: dict | None = None,
           download_to: Path | None = None):
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(
        url, data=data, method=method,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            if download_to is not None:
                with download_to.open("wb") as out:
                    while chunk := resp.read(1024 * 1024):
                        out.write(chunk)
                return {}
            payload = resp.read()
            return json.loads(payload) if payload else {}
    except urllib.error.HTTPError as error:
        raise GraphRequestError(method, error.code) from error


def find_folder(token: str, name: str, parent: str | None = None) -> str | None:
    base = f"{GRAPH}/me/mailFolders" + (f"/{parent}/childFolders" if parent else "")
    query = urllib.parse.urlencode({"$filter": f"displayName eq '{name}'", "$select": "id"})
    items = _graph(token, "GET", f"{base}?{query}").get("value", [])
    if len(items) > 1:
        raise ArchiveError(f"Plusieurs dossiers portent le nom {name}; opération refusée.")
    return items[0]["id"] if items else None


def _is_owned_archive_folder(token: str, folder: str, year: str) -> bool:
    property_id = urllib.parse.quote(ARCHIVE_OWNER_PROPERTY, safe="")
    try:
        marker = _graph(
            token,
            "GET",
            f"{GRAPH}/me/mailFolders/{folder}/singleValueExtendedProperties/{property_id}",
        )
    except GraphRequestError as error:
        if error.status_code == 404:
            return False
        raise
    return marker.get("value") == f"MesArchives PST import v1:{year}"


def ensure_archive_folder(token: str, year: str) -> str:
    name = folder_name(year)
    existing = find_folder(token, name)
    if existing:
        if not _is_owned_archive_folder(token, existing, year):
            raise ArchiveError(
                f"{name} existe déjà sans marqueur MesArchives; il ne sera ni utilisé ni supprimé."
            )
        return existing

    return _graph(
        token,
        "POST",
        f"{GRAPH}/me/mailFolders",
        {
            "displayName": name,
            "singleValueExtendedProperties": [
                {
                    "id": ARCHIVE_OWNER_PROPERTY,
                    "value": f"MesArchives PST import v1:{year}",
                }
            ],
        },
    )["id"]


def ensure_folder(token: str, name: str, parent: str | None = None) -> str:
    existing = find_folder(token, name, parent)
    if existing:
        return existing
    base = f"{GRAPH}/me/mailFolders" + (f"/{parent}/childFolders" if parent else "")
    return _graph(token, "POST", base, {"displayName": name})["id"]


def download_pst(token: str, year: str, dest: Path) -> None:
    path = urllib.parse.quote(f"{ONEDRIVE_FOLDER}/{pst_name(year)}")
    _graph(token, "GET", f"{GRAPH}/me/drive/root:/{path}:/content", download_to=dest)


def load_archive(token: str, mailbox: str, year: str,
                 progress: Callable[[int], None] | None = None) -> dict:
    root = ensure_archive_folder(token, year)
    if _graph(token, "GET", f"{GRAPH}/me/mailFolders/{root}?$select=totalItemCount").get("totalItemCount"):
        raise ArchiveError(f"{folder_name(year)} contient déjà des messages : déchargez avant de recharger.")
    cache: dict[str, str] = {"": root}
    imported = 0
    with tempfile.TemporaryDirectory() as tmp:
        pst = Path(tmp) / pst_name(year)
        download_pst(token, year, pst)
        for path, mime in iter_pst_messages(pst):
            parts = [p for p in path.split("/") if p][1:]  # ignore la racine du PST
            key, parent = "", root
            for part in parts:
                key = f"{key}/{part}"
                if key not in cache:
                    cache[key] = ensure_folder(token, part, parent)
                parent = cache[key]
            try:
                import_mime(token, mailbox, parent, mime)
            except GraphImportError as error:
                raise ArchiveError(f"Import interrompu après {imported} messages : {error}") from error
            imported += 1
            if progress:
                progress(imported)
    return {"ok": True, "annee": int(year), "imported": imported}


def unload_archive(token: str, year: str) -> dict:
    folder = find_folder(token, folder_name(year))
    if folder:
        if not _is_owned_archive_folder(token, folder, year):
            raise ArchiveError(
                f"Refus de supprimer {folder_name(year)} : son marqueur MesArchives est absent ou invalide."
            )
        _graph(token, "DELETE", f"{GRAPH}/me/mailFolders/{folder}")
    return {"ok": True, "annee": int(year), "deleted": bool(folder)}
