# PTA Reimbursement Parser

A CLI tool for processing PTA reimbursement request emails. Extracts data from scanned forms using OCR, allows review/editing, appends records to Google Sheets, and archives attachments to Google Drive.

While this is a useful tool for me specifically (and you could probably adapt it to your own needs), this was mostly a test of Claude Code's capabilities. This was made with *zero* coding help from me (and thus shouldn't be taken as evidence of my coding skills, or lack thereof!) Aside from some familiarity with cloud services and being able to ask the right sorts of prompts, this could have been created by a user with no software development experience.

## Features

- **Email Processing**: Fetch emails directly from Gmail or process local `.eml` files
- **Multi-format Support**: Handles PDF, image (.jpg, .png, .heic), and Word (.docx) attachments
- **OCR Extraction**: Uses Google Cloud Vision API to extract text from scanned forms
- **Smart Field Parsing**: Automatically extracts requestor, amount, date, teacher/grade, event, etc.
- **Interactive Review**: CLI interface to review and edit extracted data before submission
- **Google Sheets Integration**: Appends records to your expense tracking spreadsheet
- **Google Drive Archiving**: Uploads attachments to organized monthly folders
- **Printing**: Optional printing of attachments (Windows)

## Prerequisites

- Python 3.10+
- [Poppler](https://poppler.freedesktop.org/) (for PDF processing)
  - Windows: `choco install poppler` or [manual download](https://github.com/oschwartz10612/poppler-windows/releases)
- Google Cloud Project with:
  - Cloud Vision API enabled
  - Google Sheets API enabled
  - Google Drive API enabled
- OAuth2 credentials for Gmail access

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/PTAParser.git
   cd PTAParser
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Copy the example config:
   ```bash
   cp config/config.example.yaml config/config.yaml
   ```

4. Set up Google Cloud credentials (see [Configuration](#configuration))

## Configuration

### 1. Google Cloud Setup

1. Create a project in [Google Cloud Console](https://console.cloud.google.com/)
2. Enable the following APIs:
   - Cloud Vision API
   - Google Sheets API
   - Google Drive API
3. Create a **Service Account**:
   - Download the JSON key file
   - Save to `credentials/` folder
   - Share your Google Sheet with the service account email
4. Create **OAuth2 credentials** (for Gmail):
   - Create OAuth client ID (Desktop application type)
   - Download the JSON file to `credentials/`
   - Add your email as a test user in OAuth consent screen

### 2. Edit config.yaml

```yaml
# Path to Poppler (for PDF conversion)
poppler_path: "C:\\Program Files\\poppler-24.08.0\\Library\\bin"

google_cloud:
  credentials_file: "credentials/your-service-account.json"

gmail:
  oauth_credentials_file: "credentials/gmail_oauth_credentials.json"

google_drive:
  archive_folder_id: "your-drive-folder-id"

google_sheets:
  spreadsheet_id: "your-spreadsheet-id"
  sheet_name: "Income and Expenses"

field_mappings:
  payment_types:
    - "Check"
    - "Debit"
    - "Amazon"
  budget_categories:
    - "Classroom Events"
    - "Teacher Appreciation"
    # ... add your categories
  budget_items:
    - "HRP / Class Parties Fund"
    - "Staff Appreciation Expense"
    # ... add your items
```

## Usage

### List Gmail Messages

```bash
python -m src.main gmail-list
```

Shows recent emails with attachments:
```
1. Winter Party Reimbursement                [.pdf] [.png]
   From: Jane Doe <jane@email.com>
   Date: 2025-01-15 10:30
   ID: 19abc123def456
```

### Process a Gmail Message

```bash
python -m src.main gmail-process <message-id>
```

Example workflow:
```
Fetching message: 19abc123def456
  From: Jane Doe <jane@email.com>
  Subject: Winter Party Reimbursement
  Found 1 PDF attachment(s)

Extracting PDF: reimbursement_form.pdf
  Converted to 2 page(s)
  Running OCR on 2 page(s)...
  OCR complete.

=== Extracted Data ===
+-----------------+----------------------+
| Field           | Value                |
+-----------------+----------------------+
| Requestor       | Jane Doe             |
| Date            | 1/15/2025            |
| Amount          | $45.67               |
| Email           | jane@email.com       |
| Teacher/Grade   | Mrs. Smith - 3rd     |
| Type            | Home Room Parent     |
| Event           | Winter Party         |
+-----------------+----------------------+

Edit a field? ok

Select Payment Type:
  1. Check
  2. Debit
  3. Amazon
> 1

Add to spreadsheet? (y/n): y
✓ Added row #156 to "Income and Expenses"

Upload attachments to Google Drive? (y/n): y
✓ Uploaded 1 file(s) to Drive archive (JANUARY)
```

### Process a Local .eml File

```bash
python -m src.main process path/to/email.eml
```

### Dry Run Mode

Test without writing to spreadsheet:
```bash
python -m src.main gmail-process <message-id> --dry-run
```

## Project Structure

```
PTAParser/
├── src/
│   ├── main.py              # CLI entry point
│   ├── gmail_fetcher.py     # Gmail API integration
│   ├── email_parser.py      # .eml file parsing
│   ├── pdf_processor.py     # PDF to image conversion
│   ├── ocr_processor.py     # Google Vision OCR
│   ├── field_extractor.py   # Form field parsing
│   ├── cli_review.py        # Interactive review interface
│   ├── sheets_writer.py     # Google Sheets integration
│   ├── drive_uploader.py    # Google Drive archiving
│   └── printer.py           # Windows printing support
├── config/
│   ├── config.example.yaml  # Example configuration
│   └── config.yaml          # Your configuration (gitignored)
├── credentials/             # Google API credentials (gitignored)
├── requirements.txt
└── README.md
```

## Spreadsheet Column Mapping

| Form Field | Spreadsheet Column |
|------------|-------------------|
| (auto-generated) | ID |
| "Expense" | Income/Expense |
| Date (year) | Year |
| Date (month) | Month |
| Email received date | Date Received |
| Requestor | Submitted By |
| Teacher/Grade | Grade |
| Payment Type | Type |
| Budget Category | Budget Category |
| Budget Item | Budget Item |
| Amount | Amount Submitted |
| Notes + TODO | Notes |

## License

MIT License

## Contributing

Contributions are welcome! Please open an issue or submit a pull request.
