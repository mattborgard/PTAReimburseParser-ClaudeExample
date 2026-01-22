"""Google Sheets writer module for appending reimbursement records."""

import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

from google.oauth2 import service_account
from googleapiclient.discovery import build


@dataclass
class SpreadsheetRow:
    """Data for a single spreadsheet row (columns A-T)."""
    id: int                          # A
    income_expense: str = "Expense"  # B
    year: int = 0                    # C
    month: str = ""                  # D
    date_received: str = ""          # E
    submitted_by: str = ""           # F
    grade: str = ""                  # G
    type: str = "Check"              # H
    budget_category: str = ""        # I
    budget_item: str = ""            # J
    amount_submitted: str = ""       # K
    amount_paid: str = ""            # L
    check_number: str = ""           # M
    myptez: str = ""                 # N
    bank: str = ""                   # O
    reconcile: str = ""              # P
    report: str = ""                 # Q
    all_mats_printed: str = ""       # R
    double_signed: str = ""          # S
    notes: str = ""                  # T


class SheetsWriter:
    """Handle Google Sheets operations."""

    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    def __init__(
        self,
        credentials_path: str | Path,
        spreadsheet_id: str,
        sheet_name: str = "Income and Expenses"
    ):
        """
        Initialize the Sheets writer.

        Args:
            credentials_path: Path to service account JSON file
            spreadsheet_id: The Google Sheets spreadsheet ID
            sheet_name: Name of the sheet to write to
        """
        self.spreadsheet_id = spreadsheet_id
        self.sheet_name = sheet_name
        self.service = self._authenticate(credentials_path)

    def _authenticate(self, credentials_path: str | Path) -> any:
        """Authenticate with Google Sheets API."""
        credentials_path = Path(credentials_path)
        if not credentials_path.exists():
            raise FileNotFoundError(f"Credentials file not found: {credentials_path}")

        credentials = service_account.Credentials.from_service_account_file(
            str(credentials_path),
            scopes=self.SCOPES
        )

        return build('sheets', 'v4', credentials=credentials)

    def get_next_id(self) -> int:
        """
        Get the next available ID from the spreadsheet.

        Returns:
            Next ID number (max existing ID + 1)
        """
        try:
            # Read the ID column (column A)
            range_name = f"'{self.sheet_name}'!A:A"
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=range_name
            ).execute()

            values = result.get('values', [])

            # Find max numeric ID
            max_id = 0
            for row in values:
                if row:
                    try:
                        id_val = int(row[0])
                        max_id = max(max_id, id_val)
                    except (ValueError, IndexError):
                        continue

            return max_id + 1

        except Exception as e:
            # If we can't read, start from 1
            print(f"Warning: Could not read existing IDs: {e}")
            return 1

    def append_row(self, row: SpreadsheetRow) -> int:
        """
        Append a new row to the spreadsheet.

        Args:
            row: SpreadsheetRow with data to append

        Returns:
            The row number where data was inserted
        """
        # Prepare the row values in column order (A-T)
        values = [[
            row.id,              # A
            row.income_expense,  # B
            row.year,            # C
            row.month,           # D
            row.date_received,   # E
            row.submitted_by,    # F
            row.grade,           # G
            row.type,            # H
            row.budget_category, # I
            row.budget_item,     # J
            row.amount_submitted,# K
            row.amount_paid,     # L
            row.check_number,    # M
            row.myptez,          # N
            row.bank,            # O
            row.reconcile,       # P
            row.report,          # Q
            row.all_mats_printed,# R
            row.double_signed,   # S
            row.notes,           # T
        ]]

        body = {'values': values}

        # Append to the sheet
        result = self.service.spreadsheets().values().append(
            spreadsheetId=self.spreadsheet_id,
            range=f"'{self.sheet_name}'!A:T",
            valueInputOption='USER_ENTERED',
            insertDataOption='INSERT_ROWS',
            body=body
        ).execute()

        # Parse the updated range to get row number
        updated_range = result.get('updates', {}).get('updatedRange', '')
        # Extract row number from range like "'Sheet'!A123:N123"
        if ':' in updated_range:
            row_part = updated_range.split(':')[0]
            row_num = ''.join(filter(str.isdigit, row_part.split('!')[-1]))
            return int(row_num) if row_num else -1

        return -1

    def get_column_headers(self) -> list[str]:
        """Get the column headers from the first row."""
        try:
            range_name = f"'{self.sheet_name}'!1:1"
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=range_name
            ).execute()

            values = result.get('values', [[]])
            return values[0] if values else []

        except Exception as e:
            print(f"Warning: Could not read headers: {e}")
            return []


def create_spreadsheet_row(
    form_data: dict[str, str],
    email_date: Optional[datetime],
    budget_category: str,
    budget_item: str,
    next_id: int,
    payment_type: str = "Check"
) -> SpreadsheetRow:
    """
    Create a SpreadsheetRow from form data.

    Args:
        form_data: Dictionary of extracted form fields
        email_date: Date the email was received
        budget_category: Selected budget category
        budget_item: Selected budget item
        next_id: ID number for this row
        payment_type: Payment type (Check, Debit, Amazon, etc.)

    Returns:
        SpreadsheetRow ready for insertion
    """
    # Parse the form date for year/month
    form_date = form_data.get('Date', '')
    year = 0
    month = ""

    if form_date:
        # Try to parse various date formats
        for fmt in ['%m-%d-%Y', '%m/%d/%Y', '%m-%d-%y', '%m/%d/%y']:
            try:
                parsed = datetime.strptime(form_date, fmt)
                year = parsed.year
                month = parsed.strftime('%B')  # Full month name
                break
            except ValueError:
                continue

    # Format email received date
    date_received = ""
    if email_date:
        date_received = email_date.strftime('%m/%d/%Y')

    # Extract grade from Teacher/Grade field
    grade = form_data.get('Teacher/Grade', '')
    # Try to extract just the grade portion if it contains teacher name
    if '/' in grade:
        parts = grade.split('/')
        grade = parts[-1].strip() if len(parts) > 1 else grade

    # Clean up amount (remove $ if present for spreadsheet)
    amount = form_data.get('Amount', '')
    if amount.startswith('$'):
        amount = amount[1:]

    # Build notes from various fields
    notes_parts = []

    # Add TODO based on payment type
    if payment_type.lower() in ('check', 'cheque'):
        notes_parts.append("TODO: WRITE CHECK")
    elif payment_type.lower() in ('amazon', 'debit'):
        notes_parts.append("TODO: ORDER ON AMAZON")

    # Add form details
    if form_data.get('Event'):
        notes_parts.append(f"Event: {form_data['Event']}")
    if form_data.get('Child'):
        notes_parts.append(f"Child: {form_data['Child']}")
    if form_data.get('Delivery'):
        notes_parts.append(f"Delivery: {form_data['Delivery']}")

    notes = "; ".join(notes_parts)

    return SpreadsheetRow(
        id=next_id,
        income_expense="Expense",
        year=year,
        month=month,
        date_received=date_received,
        submitted_by=form_data.get('Requestor', ''),
        grade=grade,
        type=payment_type,
        budget_category=budget_category,
        budget_item=budget_item,
        amount_submitted=amount,
        amount_paid="",
        check_number="",
        notes=notes
    )
