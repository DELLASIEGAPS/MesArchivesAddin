# PST import prototype

This is an experimental conversion layer, not a deployable Outlook import service.
It converts PST email items to MIME and exposes a narrow wrapper for Microsoft Graph
`mailFolder: importMime`. It does not import automatically, retrieve files from
OneDrive, authenticate callers, create folders, prevent duplicates, or change the
Outlook add-in.

## Important limitations

- Microsoft Graph `importMime` is a **beta** API and is not supported for production.
- `libpff`/`pypff` is alpha software. PST variants and individual message properties
  may not be parsed faithfully.
- A MIME `Date` header does not guarantee Exchange will preserve the original
  `receivedDateTime`; verify sorting and timestamps before relying on them.
- Conversion rejects messages lacking sender/recipient transport headers, RTF-only
  bodies, and non-UTF-8 bodies; it does not pretend such messages were imported.
  Attachments without filenames receive generated names. Inline images and other
  PST/MAPI features are not guaranteed to round-trip.
- This path does not import calendar items, contacts, tasks, rules, or all Outlook
  metadata. It is not equivalent to mounting a PST or Microsoft's Purview PST import.
- There is no live endpoint or deployment workflow here. No mailbox or PST is read
  or changed by these modules.
- A timeout or a malformed response after a Graph import can leave the outcome
  unknown. Never retry without reconciling against the destination mailbox.

## Local checks

From the repository root, use Python 3.11 or later in a virtual environment:

```sh
python -m pip install -r services/pst-import/requirements.txt pytest
PYTHONPATH=services/pst-import python -m pytest services/pst-import/tests
```

Do not run `import_mime` against the production mailbox. Before adding an endpoint,
the service needs Entra authentication, least-privilege delegated Microsoft Graph
permissions, protected PST retrieval, durable per-message idempotency, and an
explicit dry-run/confirmation flow. Those controls are intentionally not simulated
by this conversion module.

For the intended delegated approach, Microsoft documents `Files.Read` for the
signed-in user's OneDrive and `Mail-Advanced.ReadWrite` for `importMime`. These
are separate from the Azure deployment identity; the final authentication and
consent design is not implemented. See [Graph importMime](https://learn.microsoft.com/en-us/graph/api/mailfolder-importmime?view=graph-rest-beta)
and [OneDrive permissions](https://learn.microsoft.com/en-us/onedrive/developer/rest-api/concepts/permissions_reference).
