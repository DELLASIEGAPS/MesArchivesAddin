"""Experimental conversion of PST email messages to MIME for Graph importMime."""

from __future__ import annotations

from datetime import datetime, timezone
from email.message import EmailMessage
from email.parser import HeaderParser
from email.policy import SMTP, default
from email.utils import format_datetime
from pathlib import Path
from typing import Iterator

import pypff


class PstReadError(ValueError):
    """Raised when a PST cannot be read safely."""


def _read_attachment(attachment: object) -> bytes:
    size = int(attachment.size)
    if size < 0:
        raise PstReadError("Invalid attachment size.")
    data = attachment.read_buffer(size)
    if not isinstance(data, bytes) or len(data) != size:
        raise PstReadError("Attachment data was incomplete.")
    return data


def _message_to_mime(message: object) -> bytes:
    mime = EmailMessage(policy=SMTP)
    headers = getattr(message, "transport_headers", None)
    if not isinstance(headers, str) or not headers.strip():
        raise PstReadError("Message transport headers are missing; conversion cannot preserve addresses.")
    parsed_headers = HeaderParser(policy=default).parsestr(headers)
    if parsed_headers.defects:
        raise PstReadError("Message transport headers are malformed.")
    for name in ("From", "To", "Cc", "Bcc", "Subject", "Date", "Message-ID"):
        values = parsed_headers.get_all(name, [])
        if len(values) > 1:
            raise PstReadError(f"Multiple {name} headers cannot be converted safely.")
        if values:
            if values[0].defects:
                raise PstReadError(f"Malformed {name} header.")
            mime[name] = str(values[0])

    if "From" not in mime or not any(name in mime for name in ("To", "Cc", "Bcc")):
        raise PstReadError("Message sender or recipients are missing from transport headers.")
    subject = getattr(message, "subject", None)
    if subject and "Subject" not in mime:
        mime["Subject"] = subject

    delivery_time = getattr(message, "delivery_time", None)
    if isinstance(delivery_time, datetime) and "Date" not in mime:
        if delivery_time.tzinfo is None:
            delivery_time = delivery_time.replace(tzinfo=timezone.utc)
        mime["Date"] = format_datetime(delivery_time)

    plain_body = getattr(message, "plain_text_body", None)
    html_body = getattr(message, "html_body", None)
    if not plain_body and not html_body and getattr(message, "rtf_body", None):
        raise PstReadError("RTF-only message bodies are not supported.")
    try:
        if isinstance(plain_body, bytes):
            plain_body = plain_body.decode("utf-8")
        if isinstance(html_body, bytes):
            html_body = html_body.decode("utf-8")
    except UnicodeDecodeError as error:
        raise PstReadError("Message body is not valid UTF-8.") from error
    if plain_body:
        mime.set_content(plain_body)
        if html_body:
            mime.add_alternative(html_body, subtype="html")
    elif html_body:
        mime.set_content(html_body, subtype="html")
    else:
        mime.set_content("")

    for index in range(int(getattr(message, "number_of_attachments", 0))):
        attachment = message.get_attachment(index)
        filename = getattr(attachment, "long_filename", None) or f"piece-jointe-{index + 1}"
        content = _read_attachment(attachment)
        mime.add_attachment(
            content,
            maintype="application",
            subtype="octet-stream",
            filename=filename,
        )

    return mime.as_bytes(policy=SMTP)


def iter_pst_messages(pst_path: str | Path) -> Iterator[tuple[str, bytes]]:
    """Yield (folder path, RFC 5322 MIME bytes) for mail items in a PST."""
    pst = pypff.file()
    opened = False
    try:
        pst.open(str(pst_path))
        opened = True
        root = pst.get_root_folder()

        def walk(folder: object, path: str) -> Iterator[tuple[str, bytes]]:
            name = getattr(folder, "name", None) or ""
            current_path = f"{path}/{name}".strip("/")
            for index in range(int(folder.number_of_sub_messages)):
                message = folder.get_sub_message(index)
                try:
                    yield current_path, _message_to_mime(message)
                except PstReadError as error:
                    raise PstReadError(
                        f"Message {message.identifier} in {current_path!r}: {error}"
                    ) from error
            for index in range(int(folder.number_of_sub_folders)):
                yield from walk(folder.get_sub_folder(index), current_path)

        yield from walk(root, "")
    except (OSError, IOError) as error:
        raise PstReadError("The PST could not be read or converted.") from error
    finally:
        if opened:
            pst.close()
