from datetime import datetime, timezone
from email import policy
from email.parser import BytesParser
from types import SimpleNamespace
from unittest.mock import patch

from pst_to_mime import _message_to_mime


def test_converts_message_body_headers_and_attachment():
    attachment = SimpleNamespace(
        size=4,
        long_filename="note.txt",
        read_buffer=lambda size: b"test",
    )
    message = SimpleNamespace(
        subject="Archive message",
        sender_email_address="sender@example.com",
        display_to="recipient@example.com",
        transport_headers=(
            "From: sender@example.com\r\nTo: recipient@example.com\r\n"
            "Message-ID: <archive-1@example.com>\r\n"
        ),
        delivery_time=datetime(2024, 3, 5, 10, 30, tzinfo=timezone.utc),
        plain_text_body=b"Hello from the archive",
        html_body=b"<p>Hello from the archive</p>",
        number_of_attachments=1,
        get_attachment=lambda index: attachment,
    )

    with patch("pst_to_mime.pypff"):
        result = _message_to_mime(message)

    parsed = BytesParser(policy=policy.default).parsebytes(result)
    assert parsed["Subject"] == "Archive message"
    assert parsed["From"] == "sender@example.com"
    assert parsed["To"] == "recipient@example.com"
    assert parsed["Message-ID"] == "<archive-1@example.com>"
    assert parsed["Date"] == "Tue, 05 Mar 2024 10:30:00 +0000"
    assert parsed.get_body(preferencelist=("plain",)).get_content().strip() == "Hello from the archive"
    assert parsed.get_body(preferencelist=("html",)).get_content().strip() == "<p>Hello from the archive</p>"
    files = list(parsed.iter_attachments())
    assert len(files) == 1
    assert files[0].get_filename() == "note.txt"
    assert files[0].get_content() == b"test"


def test_rejects_truncated_attachment():
    attachment = SimpleNamespace(size=4, long_filename="note.txt", read_buffer=lambda size: b"bad")
    message = SimpleNamespace(
        subject="Archive message",
        sender_email_address=None,
        display_to=None,
        transport_headers="From: sender@example.com\r\nTo: recipient@example.com\r\n",
        delivery_time=None,
        plain_text_body="body",
        html_body=None,
        number_of_attachments=1,
        get_attachment=lambda index: attachment,
    )

    with patch("pst_to_mime.pypff"):
        try:
            _message_to_mime(message)
        except ValueError as error:
            assert "incomplete" in str(error)
        else:
            raise AssertionError("Truncated attachments must be rejected.")


def test_missing_attachment_name_uses_placeholder():
    attachment = SimpleNamespace(size=4, long_filename=None, read_buffer=lambda size: b"test")
    message = SimpleNamespace(
        transport_headers="From: sender@example.com\r\nTo: recipient@example.com\r\n",
        plain_text_body=b"body",
        number_of_attachments=1,
        get_attachment=lambda index: attachment,
    )
    parsed = BytesParser(policy=policy.default).parsebytes(_message_to_mime(message))
    result = list(parsed.iter_attachments())
    assert result[0].get_filename() == "piece-jointe-1"
    assert result[0].get_content() == b"test"


def test_rejects_missing_addresses_and_rtf_only_messages():
    message = SimpleNamespace(transport_headers=None)
    try:
        _message_to_mime(message)
    except ValueError as error:
        assert "headers are missing" in str(error)
    else:
        raise AssertionError("Missing transport headers must be rejected.")

    message = SimpleNamespace(
        transport_headers="From: sender@example.com\r\nTo: recipient@example.com\r\n",
        rtf_body=b"rtf data",
        number_of_attachments=0,
    )
    try:
        _message_to_mime(message)
    except ValueError as error:
        assert "RTF-only" in str(error)
    else:
        raise AssertionError("RTF-only messages must be rejected.")
