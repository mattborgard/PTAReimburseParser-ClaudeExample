"""Gmail fetcher module for retrieving emails via Gmail API."""

import base64
import os
import pickle
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


# API scopes for Gmail and Drive
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/drive.file'  # For uploading to Drive
]


@dataclass
class GmailMessage:
    """Represents a Gmail message with metadata."""
    id: str
    subject: str
    sender_name: str
    sender_email: str
    date: Optional[datetime]
    snippet: str
    attachment_types: list[str]  # List of file extensions (e.g., ['.pdf', '.docx'])


@dataclass
class FetchedEmail:
    """Email data fetched from Gmail, ready for processing."""
    message_id: str
    sender_name: str
    sender_email: str
    subject: str
    date: Optional[datetime]
    body_text: str
    pdf_paths: list[Path]  # PDFs (need conversion before OCR)
    image_paths: list[Path]  # Images (can be OCR'd directly)
    doc_paths: list[Path]  # Word docs (text extraction, no OCR)
    attachment_paths: list[Path]  # All attachments (for Drive upload)

    def has_processable_files(self) -> bool:
        """Check if there are any files that can be processed."""
        return bool(self.pdf_paths or self.image_paths or self.doc_paths)


class GmailFetcher:
    """Fetch emails from Gmail using OAuth2."""

    def __init__(
        self,
        oauth_credentials_path: str | Path,
        token_path: Optional[str | Path] = None
    ):
        """
        Initialize Gmail fetcher.

        Args:
            oauth_credentials_path: Path to OAuth2 client credentials JSON
            token_path: Path to store/load token (default: credentials/gmail_token.pickle)
        """
        self.oauth_credentials_path = Path(oauth_credentials_path)
        if not self.oauth_credentials_path.exists():
            raise FileNotFoundError(
                f"OAuth credentials not found: {oauth_credentials_path}\n"
                "Download OAuth client credentials from Google Cloud Console."
            )

        if token_path:
            self.token_path = Path(token_path)
        else:
            self.token_path = self.oauth_credentials_path.parent / 'gmail_token.pickle'

        self._credentials = None
        self.service = self._authenticate()

    @property
    def credentials(self) -> Credentials:
        """Get the OAuth2 credentials for use with other Google APIs."""
        return self._credentials

    def _authenticate(self):
        """Authenticate with Gmail API using OAuth2."""
        creds = None

        # Load existing token if available
        if self.token_path.exists():
            with open(self.token_path, 'rb') as token:
                creds = pickle.load(token)

        # Check if credentials have all required scopes
        needs_reauth = False
        if creds and creds.valid:
            # Check if all scopes are present
            if hasattr(creds, 'scopes') and creds.scopes:
                for scope in SCOPES:
                    if scope not in creds.scopes:
                        needs_reauth = True
                        break

        # Refresh or get new credentials
        if not creds or not creds.valid or needs_reauth:
            if creds and creds.expired and creds.refresh_token and not needs_reauth:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    str(self.oauth_credentials_path),
                    SCOPES
                )
                creds = flow.run_local_server(port=0)

            # Save credentials for next run
            with open(self.token_path, 'wb') as token:
                pickle.dump(creds, token)

        self._credentials = creds
        return build('gmail', 'v1', credentials=creds)

    def list_messages(
        self,
        query: str = "has:attachment",
        max_results: int = 20
    ) -> list[GmailMessage]:
        """
        List Gmail messages matching a query.

        Args:
            query: Gmail search query (default: emails with attachments)
            max_results: Maximum number of messages to return

        Returns:
            List of GmailMessage objects
        """
        results = self.service.users().messages().list(
            userId='me',
            q=query,
            maxResults=max_results
        ).execute()

        messages = []
        for msg_info in results.get('messages', []):
            msg = self.service.users().messages().get(
                userId='me',
                id=msg_info['id'],
                format='full'
            ).execute()

            headers = {h['name']: h['value'] for h in msg.get('payload', {}).get('headers', [])}

            # Parse sender
            from_header = headers.get('From', '')
            sender_name, sender_email = self._parse_sender(from_header)

            # Parse date
            date_str = headers.get('Date', '')
            msg_date = self._parse_date(date_str)

            # Get attachment types
            attachment_types = self._get_attachment_types(msg.get('payload', {}))

            messages.append(GmailMessage(
                id=msg_info['id'],
                subject=headers.get('Subject', '(no subject)'),
                sender_name=sender_name,
                sender_email=sender_email,
                date=msg_date,
                snippet=msg.get('snippet', ''),
                attachment_types=attachment_types
            ))

        return messages

    def fetch_message(self, message_id: str) -> FetchedEmail:
        """
        Fetch a complete message with attachments.

        Args:
            message_id: Gmail message ID

        Returns:
            FetchedEmail with extracted data and PDF paths
        """
        msg = self.service.users().messages().get(
            userId='me',
            id=message_id,
            format='full'
        ).execute()

        headers = {h['name']: h['value'] for h in msg.get('payload', {}).get('headers', [])}

        # Parse sender
        from_header = headers.get('From', '')
        sender_name, sender_email = self._parse_sender(from_header)

        # Parse date
        date_str = headers.get('Date', '')
        msg_date = self._parse_date(date_str)

        # Extract body text
        body_text = self._extract_body(msg.get('payload', {}))

        # Extract all attachments
        all_attachments = self._extract_attachments(msg.get('payload', {}), message_id)

        # Separate files by type for processing
        # PDFs need conversion before OCR
        pdf_paths = [p for p in all_attachments if p.suffix.lower() == '.pdf']
        # Images can be OCR'd directly (Vision API supports these formats)
        image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.heic'}
        image_paths = [p for p in all_attachments if p.suffix.lower() in image_extensions]
        # Document files that contain text directly
        doc_extensions = {'.docx', '.doc'}
        doc_paths = [p for p in all_attachments if p.suffix.lower() in doc_extensions]

        return FetchedEmail(
            message_id=message_id,
            sender_name=sender_name,
            sender_email=sender_email,
            subject=headers.get('Subject', ''),
            date=msg_date,
            body_text=body_text,
            pdf_paths=pdf_paths,
            image_paths=image_paths,
            doc_paths=doc_paths,
            attachment_paths=all_attachments
        )

    def _parse_sender(self, from_header: str) -> tuple[str, str]:
        """Parse sender name and email from From header."""
        import re
        # Format: "Name <email@example.com>" or just "email@example.com"
        match = re.match(r'^"?([^"<]*)"?\s*<?([^>]+@[^>]+)>?$', from_header.strip())
        if match:
            name = match.group(1).strip()
            email = match.group(2).strip()
            return (name, email)
        return ('', from_header.strip())

    def _parse_date(self, date_str: str) -> Optional[datetime]:
        """Parse email date string."""
        from email.utils import parsedate_to_datetime
        try:
            return parsedate_to_datetime(date_str)
        except (ValueError, TypeError):
            return None

    def _get_attachment_types(self, payload: dict, types: set = None) -> list[str]:
        """Get list of attachment file extensions in payload."""
        if types is None:
            types = set()

        parts = payload.get('parts', [])
        if not parts:
            # Single part message
            filename = payload.get('filename', '')
            if filename:
                ext = Path(filename).suffix.lower()
                if ext:
                    types.add(ext)
            return sorted(types)

        for part in parts:
            filename = part.get('filename', '')
            if filename:
                ext = Path(filename).suffix.lower()
                if ext:
                    types.add(ext)
            # Check nested parts
            if 'parts' in part:
                self._get_attachment_types(part, types)

        return sorted(types)

    def _extract_body(self, payload: dict) -> str:
        """Extract plain text body from payload."""
        parts = payload.get('parts', [])

        if not parts:
            # Single part message
            if payload.get('mimeType') == 'text/plain':
                data = payload.get('body', {}).get('data', '')
                if data:
                    return base64.urlsafe_b64decode(data).decode('utf-8', errors='replace')
            return ''

        # Multi-part message
        for part in parts:
            mime_type = part.get('mimeType', '')
            if mime_type == 'text/plain':
                data = part.get('body', {}).get('data', '')
                if data:
                    return base64.urlsafe_b64decode(data).decode('utf-8', errors='replace')
            # Check nested parts
            if 'parts' in part:
                text = self._extract_body(part)
                if text:
                    return text

        return ''

    def _extract_attachments(self, payload: dict, message_id: str) -> list[Path]:
        """Extract all attachments and save to temp directory."""
        temp_dir = Path(tempfile.gettempdir()) / 'pta_parser' / 'gmail'
        temp_dir.mkdir(parents=True, exist_ok=True)

        attachment_paths = []
        self._find_and_save_attachments(payload, message_id, temp_dir, attachment_paths)
        return attachment_paths

    def _find_and_save_attachments(
        self,
        payload: dict,
        message_id: str,
        temp_dir: Path,
        attachment_paths: list[Path]
    ) -> None:
        """Recursively find and save all attachments."""
        parts = payload.get('parts', [])

        if not parts:
            # Check single part
            filename = payload.get('filename', '')
            if filename:  # Any file with a filename is an attachment
                self._save_attachment(payload, message_id, filename, temp_dir, attachment_paths)
            return

        for part in parts:
            filename = part.get('filename', '')
            if filename:  # Any file with a filename is an attachment
                self._save_attachment(part, message_id, filename, temp_dir, attachment_paths)

            # Check nested parts
            if 'parts' in part:
                self._find_and_save_attachments(part, message_id, temp_dir, attachment_paths)

    def _save_attachment(
        self,
        part: dict,
        message_id: str,
        filename: str,
        temp_dir: Path,
        file_paths: list[Path]
    ) -> None:
        """Save a single attachment to disk."""
        body = part.get('body', {})
        attachment_id = body.get('attachmentId')

        if attachment_id:
            # Fetch attachment data
            attachment = self.service.users().messages().attachments().get(
                userId='me',
                messageId=message_id,
                id=attachment_id
            ).execute()
            data = attachment.get('data', '')
        else:
            data = body.get('data', '')

        if data:
            file_data = base64.urlsafe_b64decode(data)

            # Sanitize filename but preserve extension
            name_part = Path(filename).stem
            ext_part = Path(filename).suffix
            safe_name = "".join(c for c in name_part if c.isalnum() or c in '._- ')
            safe_filename = f"{safe_name}{ext_part}"

            file_path = temp_dir / f"{message_id[:8]}_{safe_filename}"

            # Handle duplicates
            counter = 1
            original = file_path
            while file_path.exists():
                file_path = temp_dir / f"{original.stem}_{counter}{original.suffix}"
                counter += 1

            file_path.write_bytes(file_data)
            file_paths.append(file_path)


def cleanup_fetched_files(file_paths: list[Path]) -> None:
    """Remove temporary files."""
    for path in file_paths:
        try:
            if path.exists():
                path.unlink()
        except OSError:
            pass


def extract_text_from_docx(doc_path: Path) -> str:
    """Extract text from a .docx file."""
    try:
        from docx import Document
        doc = Document(str(doc_path))
        paragraphs = [para.text for para in doc.paragraphs]
        return '\n'.join(paragraphs)
    except ImportError:
        raise RuntimeError(
            "python-docx is required for .docx processing. "
            "Install it with: pip install python-docx"
        )
    except Exception as e:
        raise RuntimeError(f"Failed to extract text from {doc_path.name}: {e}")
