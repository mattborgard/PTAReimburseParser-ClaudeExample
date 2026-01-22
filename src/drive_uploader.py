"""Google Drive uploader module for archiving attachments."""

import os
import re
from pathlib import Path
from typing import Optional

from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload


class DriveUploader:
    """Upload and organize files in Google Drive."""

    SCOPES = ['https://www.googleapis.com/auth/drive.file']

    def __init__(
        self,
        credentials: Credentials,
        archive_folder_id: str
    ):
        """
        Initialize Drive uploader.

        Args:
            credentials: Google OAuth2 credentials
            archive_folder_id: ID of the archive folder in Drive
        """
        self.archive_folder_id = archive_folder_id
        self.service = build('drive', 'v3', credentials=credentials)
        self._month_folder_cache = {}

    def _list_subfolders(self, parent_id: str) -> dict[str, str]:
        """
        List subfolders in a parent folder.

        Args:
            parent_id: ID of the parent folder

        Returns:
            Dict mapping normalized folder name to folder ID
        """
        results = self.service.files().list(
            q=f"'{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
            fields="files(id, name)",
            pageSize=100
        ).execute()

        folders = {}
        for item in results.get('files', []):
            # Store with original name as key, but we'll normalize during lookup
            folders[item['name']] = item['id']

        return folders

    def _find_month_folder(self, month: str) -> Optional[str]:
        """
        Find the folder ID for a specific month.

        Args:
            month: Month name (e.g., "December" or "DECEMBER")

        Returns:
            Folder ID if found, None otherwise
        """
        # Always refresh the folder list to ensure we find existing folders
        self._month_folder_cache = self._list_subfolders(self.archive_folder_id)

        # Normalize month: uppercase, strip whitespace
        month_normalized = month.strip().upper()

        # Look for exact match first (e.g., "DECEMBER")
        for folder_name, folder_id in self._month_folder_cache.items():
            folder_normalized = folder_name.strip().upper()
            if folder_normalized == month_normalized:
                return folder_id

        # Fallback: partial match (e.g., "JANUARY" in "JANUARY 2025")
        for folder_name, folder_id in self._month_folder_cache.items():
            folder_normalized = folder_name.strip().upper()
            if month_normalized in folder_normalized or folder_normalized in month_normalized:
                return folder_id

        return None

    def _get_or_create_month_folder(self, month: str) -> str:
        """
        Get existing month folder or create one.

        Args:
            month: Month name

        Returns:
            Folder ID
        """
        folder_id = self._find_month_folder(month)
        if folder_id:
            return folder_id

        # Create new folder with uppercase month name only
        folder_name = month.upper()
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [self.archive_folder_id]
        }

        folder = self.service.files().create(
            body=file_metadata,
            fields='id'
        ).execute()

        folder_id = folder.get('id')
        self._month_folder_cache[folder_name] = folder_id
        return folder_id

    def upload_attachment(
        self,
        file_path: Path,
        entry_id: int,
        requestor_name: str,
        month: str,
        file_index: int = 1,
        file_type: str = "Reimbursement"
    ) -> str:
        """
        Upload an attachment to the appropriate month folder.

        Args:
            file_path: Path to the file to upload
            entry_id: Spreadsheet entry ID number
            requestor_name: Full name of the requestor
            month: Month name for folder organization (e.g., "January")
            file_index: Index if multiple files (1, 2, 3...)
            file_type: Type label (Invoice, Receipt, Reimbursement)

        Returns:
            ID of the uploaded file
        """
        # Get the month folder
        folder_id = self._get_or_create_month_folder(month)

        # Extract last name from requestor name
        last_name = self._extract_last_name(requestor_name)

        # Get file extension
        extension = file_path.suffix.lower()

        # Build new filename: "<ID> <LAST NAME> <TYPE> <INDEX>.ext"
        # e.g., "156 Kim Reimbursement 1.pdf"
        new_filename = f"{entry_id} {last_name} {file_type} {file_index}{extension}"

        # Determine MIME type
        mime_type = self._get_mime_type(extension)

        # Upload file
        file_metadata = {
            'name': new_filename,
            'parents': [folder_id]
        }

        media = MediaFileUpload(
            str(file_path),
            mimetype=mime_type,
            resumable=True
        )

        uploaded_file = self.service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, webViewLink'
        ).execute()

        return uploaded_file.get('id')

    def upload_attachments(
        self,
        file_paths: list[Path],
        entry_id: int,
        requestor_name: str,
        month: str
    ) -> list[str]:
        """
        Upload multiple attachments.

        Args:
            file_paths: List of file paths to upload
            entry_id: Spreadsheet entry ID
            requestor_name: Requestor's name
            month: Month name for folder (e.g., "January")

        Returns:
            List of uploaded file IDs
        """
        file_ids = []

        for i, file_path in enumerate(file_paths, 1):
            # Determine file type based on filename or default to Reimbursement
            file_type = self._detect_file_type(file_path.name)

            file_id = self.upload_attachment(
                file_path=file_path,
                entry_id=entry_id,
                requestor_name=requestor_name,
                month=month,
                file_index=i,
                file_type=file_type
            )
            file_ids.append(file_id)

        return file_ids

    def _extract_last_name(self, full_name: str) -> str:
        """Extract last name from full name."""
        if not full_name:
            return "Unknown"

        # Split by spaces and take the last part
        parts = full_name.strip().split()
        if len(parts) >= 2:
            return parts[-1]
        return parts[0] if parts else "Unknown"

    def _detect_file_type(self, filename: str) -> str:
        """Detect file type from filename."""
        filename_lower = filename.lower()

        if 'invoice' in filename_lower:
            return 'Invoice'
        elif 'receipt' in filename_lower:
            return 'Receipt'
        else:
            return 'Reimbursement'

    def _get_mime_type(self, extension: str) -> str:
        """Get MIME type for file extension."""
        mime_types = {
            '.pdf': 'application/pdf',
            '.png': 'image/png',
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.gif': 'image/gif',
            '.doc': 'application/msword',
            '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            '.xls': 'application/vnd.ms-excel',
            '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }
        return mime_types.get(extension, 'application/octet-stream')


def create_drive_uploader_from_gmail_token(
    token_path: Path,
    archive_folder_id: str
) -> DriveUploader:
    """
    Create a DriveUploader using existing Gmail OAuth token.

    Note: The token must have been created with Drive scope included.

    Args:
        token_path: Path to the Gmail OAuth token pickle file
        archive_folder_id: ID of the archive folder

    Returns:
        Configured DriveUploader
    """
    import pickle

    if not token_path.exists():
        raise FileNotFoundError(f"Token file not found: {token_path}")

    with open(token_path, 'rb') as f:
        creds = pickle.load(f)

    return DriveUploader(creds, archive_folder_id)
