"""Email parser module for extracting data and attachments from .eml files."""

import email
import os
import tempfile
from dataclasses import dataclass
from datetime import datetime
from email import policy
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path
from typing import Optional


@dataclass
class EmailData:
    """Extracted email metadata and content."""
    sender_name: str
    sender_email: str
    subject: str
    date: Optional[datetime]
    body_text: str
    pdf_paths: list[Path]


def parse_eml_file(eml_path: str | Path) -> EmailData:
    """
    Parse an .eml file and extract metadata, body, and PDF attachments.

    Args:
        eml_path: Path to the .eml file

    Returns:
        EmailData object with extracted information

    Raises:
        FileNotFoundError: If the .eml file doesn't exist
        ValueError: If the file cannot be parsed as an email
    """
    eml_path = Path(eml_path)
    if not eml_path.exists():
        raise FileNotFoundError(f"Email file not found: {eml_path}")

    with open(eml_path, 'rb') as f:
        msg = email.message_from_binary_file(f, policy=policy.default)

    # Extract sender information
    sender_name, sender_email = parseaddr(msg.get('From', ''))

    # Extract date
    date_str = msg.get('Date')
    email_date = None
    if date_str:
        try:
            email_date = parsedate_to_datetime(date_str)
        except (ValueError, TypeError):
            pass

    # Extract subject
    subject = msg.get('Subject', '')

    # Extract body text
    body_text = _extract_body_text(msg)

    # Extract PDF attachments
    pdf_paths = _extract_pdf_attachments(msg, eml_path.stem)

    return EmailData(
        sender_name=sender_name,
        sender_email=sender_email,
        subject=subject,
        date=email_date,
        body_text=body_text,
        pdf_paths=pdf_paths
    )


def _extract_body_text(msg: email.message.Message) -> str:
    """Extract plain text body from email message."""
    body_parts = []

    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get('Content-Disposition', ''))

            # Skip attachments
            if 'attachment' in content_disposition:
                continue

            if content_type == 'text/plain':
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or 'utf-8'
                    try:
                        body_parts.append(payload.decode(charset, errors='replace'))
                    except (UnicodeDecodeError, LookupError):
                        body_parts.append(payload.decode('utf-8', errors='replace'))
    else:
        if msg.get_content_type() == 'text/plain':
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or 'utf-8'
                try:
                    body_parts.append(payload.decode(charset, errors='replace'))
                except (UnicodeDecodeError, LookupError):
                    body_parts.append(payload.decode('utf-8', errors='replace'))

    return '\n'.join(body_parts)


def _extract_pdf_attachments(msg: email.message.Message, prefix: str) -> list[Path]:
    """
    Extract PDF attachments from email and save to temp directory.

    Args:
        msg: Email message object
        prefix: Prefix for temp file names

    Returns:
        List of paths to extracted PDF files
    """
    pdf_paths = []
    temp_dir = Path(tempfile.gettempdir()) / 'pta_parser'
    temp_dir.mkdir(exist_ok=True)

    for part in msg.walk():
        content_disposition = str(part.get('Content-Disposition', ''))
        content_type = part.get_content_type()

        # Check if this is a PDF attachment
        is_pdf = (
            content_type == 'application/pdf' or
            content_type == 'application/octet-stream' and
            _get_filename(part).lower().endswith('.pdf')
        )

        if 'attachment' in content_disposition or is_pdf:
            filename = _get_filename(part)
            if filename and filename.lower().endswith('.pdf'):
                payload = part.get_payload(decode=True)
                if payload:
                    # Create unique filename
                    safe_filename = _sanitize_filename(filename)
                    pdf_path = temp_dir / f"{prefix}_{safe_filename}"

                    # Handle duplicate filenames
                    counter = 1
                    original_path = pdf_path
                    while pdf_path.exists():
                        stem = original_path.stem
                        pdf_path = temp_dir / f"{stem}_{counter}.pdf"
                        counter += 1

                    pdf_path.write_bytes(payload)
                    pdf_paths.append(pdf_path)

    return pdf_paths


def _get_filename(part: email.message.Message) -> str:
    """Get filename from email part."""
    filename = part.get_filename()
    if filename:
        return filename

    # Try Content-Type name parameter
    content_type = part.get('Content-Type', '')
    if 'name=' in content_type:
        import re
        match = re.search(r'name="?([^";\n]+)"?', content_type)
        if match:
            return match.group(1)

    return ''


def _sanitize_filename(filename: str) -> str:
    """Sanitize filename for safe filesystem use."""
    # Remove or replace problematic characters
    import re
    sanitized = re.sub(r'[<>:"/\\|?*]', '_', filename)
    return sanitized.strip()


def cleanup_temp_files(pdf_paths: list[Path]) -> None:
    """Remove temporary PDF files."""
    for path in pdf_paths:
        try:
            if path.exists():
                path.unlink()
        except OSError:
            pass
