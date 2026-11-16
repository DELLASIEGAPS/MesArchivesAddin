# PST import prototype

This is an experimental prototype, not a production-ready Outlook import service.
It includes PST-to-MIME conversion, OneDrive retrieval, Graph import and an Azure
Function with delegated authentication. The Outlook add-in does not call this
Function, and the prototype has not been validated end-to-end against a live mailbox.

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
- Loading can partially import messages. There is no durable per-message
  idempotency or recovery, and a retry can create duplicates.
- Unloading deletes the marked MesArchives mailbox folder and all of its contents.
  The HTTP endpoint requires `confirmDelete: true`; folders without the
  MesArchives ownership marker are never reused or deleted. User-added content
  inside a marked archive folder is also deleted.
- The import runs synchronously in an HTTP request. Azure's HTTP response timeout
  can occur after 230 seconds even with a longer function timeout; long imports
  need a durable asynchronous design.
- The Power Automate marker reports a marker, not a successful mailbox import.
  The existing Web button does not mount a PST.
- A timeout or a malformed response after a Graph import can leave the outcome
  unknown. Never retry without reconciling against the destination mailbox.

## Local checks

From the repository root, use Python 3.11 or later in a virtual environment:

```sh
python -m pip install -r services/pst-import/requirements.txt pytest
PYTHONPATH=services/pst-import python -m pytest services/pst-import/tests
```

Do not deploy or run the endpoint against a production mailbox. Before deployment,
the service needs durable per-message idempotency and recovery, safe folder ownership
and deletion, an asynchronous job design, and an explicit dry-run/confirmation
flow. Validate authentication, least-privilege permissions and the complete flow
with a dedicated test mailbox first.

For the intended delegated approach, Microsoft documents `Files.Read` for the
signed-in user's OneDrive and `Mail-Advanced.ReadWrite` for `importMime`. These
are separate from the Azure deployment identity; the authentication and consent
design still requires end-to-end validation. See [Graph importMime](https://learn.microsoft.com/en-us/graph/api/mailfolder-importmime?view=graph-rest-beta)
and [OneDrive permissions](https://learn.microsoft.com/en-us/onedrive/developer/rest-api/concepts/permissions_reference).
