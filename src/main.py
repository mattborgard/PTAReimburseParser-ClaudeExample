"""Main CLI entry point for PTA Reimbursement Parser."""

import argparse
import sys
from pathlib import Path

import yaml

from . import email_parser
from . import pdf_processor
from . import ocr_processor
from . import field_extractor
from . import cli_review
from . import sheets_writer
from . import gmail_fetcher
from . import drive_uploader
from . import printer


def load_config(config_path: Path) -> dict:
    """Load configuration from YAML file."""
    if not config_path.exists():
        raise FileNotFoundError(
            f"Config file not found: {config_path}\n"
            "Copy config/config.example.yaml to config/config.yaml and fill in your values."
        )

    with open(config_path, 'r') as f:
        return yaml.safe_load(f)


def process_eml_file(
    eml_path: Path,
    config: dict,
    dry_run: bool = False
) -> bool:
    """
    Process a single .eml file.

    Args:
        eml_path: Path to the .eml file
        config: Configuration dictionary
        dry_run: If True, don't write to Google Sheets

    Returns:
        True if processing was successful
    """
    print(f"\nProcessing: {eml_path.name}")

    # Step 1: Parse email
    try:
        email_data = email_parser.parse_eml_file(eml_path)
        print(f"  From: {email_data.sender_name} <{email_data.sender_email}>")
        print(f"  Date: {email_data.date}")
    except Exception as e:
        cli_review.display_error(f"Failed to parse email: {e}")
        return False

    if not email_data.pdf_paths:
        cli_review.display_error("No PDF attachments found in email.")
        return False

    print(f"  Found {len(email_data.pdf_paths)} PDF attachment(s)")

    # Process each PDF
    all_ocr_text = []
    all_image_paths = []

    for pdf_path in email_data.pdf_paths:
        print(f"\nExtracting: {pdf_path.name}")

        # Step 2: Convert PDF to images
        try:
            poppler_path = config.get('poppler_path')
            image_paths = pdf_processor.convert_pdf_to_images(pdf_path, poppler_path=poppler_path)
            all_image_paths.extend(image_paths)
            print(f"  Converted to {len(image_paths)} page(s)")
        except RuntimeError as e:
            cli_review.display_error(str(e))
            return False
        except Exception as e:
            cli_review.display_error(f"Failed to convert PDF: {e}")
            return False

        # Step 3: Run OCR
        try:
            print(f"  Running OCR on {len(image_paths)} page(s)...")
            credentials_path = config['google_cloud']['credentials_file']
            vision_client = ocr_processor.initialize_vision_client(credentials_path)
            ocr_result = ocr_processor.process_images(vision_client, image_paths)
            all_ocr_text.append(ocr_result.full_text)
            print("  OCR complete.")
        except Exception as e:
            cli_review.display_error(f"OCR failed: {e}")
            # Cleanup
            pdf_processor.cleanup_images(all_image_paths)
            return False

    # Combine OCR text from all PDFs
    combined_ocr_text = "\n\n=== Next PDF ===\n\n".join(all_ocr_text)

    # Step 4: Extract form fields
    form_data = field_extractor.extract_fields(combined_ocr_text)
    data_dict = field_extractor.form_data_to_dict(form_data)

    # Add raw text for reference (hidden from main display)
    data_dict['_raw_text'] = combined_ocr_text

    # Step 5: CLI review
    try:
        reviewed_data = cli_review.review_and_edit(data_dict)
    except KeyboardInterrupt:
        print("\nCancelled by user.")
        # Cleanup
        pdf_processor.cleanup_images(all_image_paths)
        email_parser.cleanup_temp_files(email_data.pdf_paths)
        return False

    # Remove raw text from data
    reviewed_data.pop('_raw_text', None)

    # Step 6: Select payment type, budget category and item
    payment_types = config.get('field_mappings', {}).get('payment_types', ['Check', 'Debit', 'Amazon'])
    budget_categories = config.get('field_mappings', {}).get('budget_categories', [])
    budget_items = config.get('field_mappings', {}).get('budget_items', [])

    payment_type = cli_review.select_from_list(
        payment_types,
        "Select Payment Type"
    )

    if budget_categories:
        budget_category = cli_review.select_from_list(
            budget_categories,
            "Select Budget Category"
        )
    else:
        budget_category = input("\nEnter Budget Category: ").strip()

    if budget_items:
        budget_item = cli_review.select_from_list(
            budget_items,
            "Select Budget Item"
        )
    else:
        budget_item = input("\nEnter Budget Item: ").strip()

    # Step 7: Confirm and write to spreadsheet
    if dry_run:
        cli_review.display_info("Dry run mode - not writing to spreadsheet")
        print("\nData that would be written:")
        print(f"  Payment Type: {payment_type}")
        print(f"  Budget Category: {budget_category}")
        print(f"  Budget Item: {budget_item}")
        for key, value in reviewed_data.items():
            print(f"  {key}: {value}")
    else:
        if cli_review.confirm_action("Add to spreadsheet?"):
            try:
                writer = sheets_writer.SheetsWriter(
                    credentials_path=config['google_cloud']['credentials_file'],
                    spreadsheet_id=config['google_sheets']['spreadsheet_id'],
                    sheet_name=config['google_sheets'].get('sheet_name', 'Income and Expenses')
                )

                next_id = writer.get_next_id()

                row = sheets_writer.create_spreadsheet_row(
                    form_data=reviewed_data,
                    email_date=email_data.date,
                    budget_category=budget_category,
                    budget_item=budget_item,
                    next_id=next_id,
                    payment_type=payment_type
                )

                row_num = writer.append_row(row)
                cli_review.display_success(
                    f"Added row #{next_id} to \"{config['google_sheets'].get('sheet_name', 'Income and Expenses')}\""
                )

            except Exception as e:
                cli_review.display_error(f"Failed to write to spreadsheet: {e}")
                return False
        else:
            print("Skipped adding to spreadsheet.")

    # Cleanup temporary files
    pdf_processor.cleanup_images(all_image_paths)
    email_parser.cleanup_temp_files(email_data.pdf_paths)

    return True


def process_folder(
    folder_path: Path,
    config: dict,
    dry_run: bool = False
) -> tuple[int, int]:
    """
    Process all .eml files in a folder.

    Args:
        folder_path: Path to the folder
        config: Configuration dictionary
        dry_run: If True, don't write to Google Sheets

    Returns:
        Tuple of (successful_count, failed_count)
    """
    if not folder_path.is_dir():
        raise NotADirectoryError(f"Not a directory: {folder_path}")

    eml_files = list(folder_path.glob("*.eml"))

    if not eml_files:
        print(f"No .eml files found in {folder_path}")
        return (0, 0)

    print(f"Found {len(eml_files)} .eml file(s)")

    successful = 0
    failed = 0

    for eml_file in eml_files:
        try:
            if process_eml_file(eml_file, config, dry_run):
                successful += 1
            else:
                failed += 1
        except Exception as e:
            cli_review.display_error(f"Error processing {eml_file.name}: {e}")
            failed += 1

        # Ask to continue if there are more files
        if eml_file != eml_files[-1]:
            if not cli_review.confirm_action("\nContinue to next file?"):
                print("Stopping batch processing.")
                break

    print(f"\n=== Summary ===")
    print(f"Successful: {successful}")
    print(f"Failed: {failed}")

    return (successful, failed)


def list_gmail_messages(config: dict, query: str, max_results: int) -> None:
    """
    List Gmail messages matching a query.

    Args:
        config: Configuration dictionary
        query: Gmail search query
        max_results: Maximum number of messages to show
    """
    oauth_path = config.get('gmail', {}).get('oauth_credentials_file')
    if not oauth_path:
        cli_review.display_error(
            "Gmail OAuth credentials not configured.\n"
            "Add 'gmail.oauth_credentials_file' to config.yaml"
        )
        return

    try:
        fetcher = gmail_fetcher.GmailFetcher(oauth_path)
    except FileNotFoundError as e:
        cli_review.display_error(str(e))
        return

    print(f"\nSearching Gmail: {query}")
    print(f"Max results: {max_results}\n")

    messages = fetcher.list_messages(query=query, max_results=max_results)

    if not messages:
        print("No messages found.")
        return

    print(f"Found {len(messages)} message(s):\n")
    print("-" * 80)

    for i, msg in enumerate(messages, 1):
        date_str = msg.date.strftime('%Y-%m-%d %H:%M') if msg.date else 'Unknown'
        attachments_indicator = " ".join(f"[{ext}]" for ext in msg.attachment_types) if msg.attachment_types else ""

        print(f"{i}. {msg.subject[:50]:<50} {attachments_indicator}")
        print(f"   From: {msg.sender_name} <{msg.sender_email}>")
        print(f"   Date: {date_str}")
        print(f"   ID: {msg.id}")
        print()


def process_gmail_message(
    message_id: str,
    config: dict,
    dry_run: bool = False
) -> bool:
    """
    Fetch and process a Gmail message by ID.

    Args:
        message_id: Gmail message ID
        config: Configuration dictionary
        dry_run: If True, don't write to Google Sheets

    Returns:
        True if processing was successful
    """
    oauth_path = config.get('gmail', {}).get('oauth_credentials_file')
    if not oauth_path:
        cli_review.display_error(
            "Gmail OAuth credentials not configured.\n"
            "Add 'gmail.oauth_credentials_file' to config.yaml"
        )
        return False

    try:
        fetcher = gmail_fetcher.GmailFetcher(oauth_path)
    except FileNotFoundError as e:
        cli_review.display_error(str(e))
        return False

    print(f"\nFetching message: {message_id}")

    try:
        email_data = fetcher.fetch_message(message_id)
    except Exception as e:
        cli_review.display_error(f"Failed to fetch message: {e}")
        return False

    print(f"  From: {email_data.sender_name} <{email_data.sender_email}>")
    print(f"  Subject: {email_data.subject}")
    print(f"  Date: {email_data.date}")

    if not email_data.has_processable_files():
        cli_review.display_error("No PDF, image, or document attachments found in email.")
        return False

    if email_data.pdf_paths:
        print(f"  Found {len(email_data.pdf_paths)} PDF attachment(s)")
    if email_data.image_paths:
        print(f"  Found {len(email_data.image_paths)} image attachment(s)")
    if email_data.doc_paths:
        print(f"  Found {len(email_data.doc_paths)} document attachment(s)")

    # Process all OCR-able files
    all_ocr_text = []
    all_image_paths = []  # Track converted images for cleanup

    # Process PDFs - convert to images first
    for pdf_path in email_data.pdf_paths:
        print(f"\nExtracting PDF: {pdf_path.name}")

        try:
            poppler_path = config.get('poppler_path')
            image_paths = pdf_processor.convert_pdf_to_images(pdf_path, poppler_path=poppler_path)
            all_image_paths.extend(image_paths)
            print(f"  Converted to {len(image_paths)} page(s)")
        except RuntimeError as e:
            cli_review.display_error(str(e))
            return False
        except Exception as e:
            cli_review.display_error(f"Failed to convert PDF: {e}")
            return False

        try:
            print(f"  Running OCR on {len(image_paths)} page(s)...")
            credentials_path = config['google_cloud']['credentials_file']
            vision_client = ocr_processor.initialize_vision_client(credentials_path)
            ocr_result = ocr_processor.process_images(vision_client, image_paths)
            all_ocr_text.append(ocr_result.full_text)
            print("  OCR complete.")
        except Exception as e:
            cli_review.display_error(f"OCR failed: {e}")
            pdf_processor.cleanup_images(all_image_paths)
            return False

    # Process image attachments directly (no PDF conversion needed)
    if email_data.image_paths:
        print(f"\nProcessing {len(email_data.image_paths)} image attachment(s)...")
        try:
            credentials_path = config['google_cloud']['credentials_file']
            vision_client = ocr_processor.initialize_vision_client(credentials_path)
            ocr_result = ocr_processor.process_images(vision_client, email_data.image_paths)
            all_ocr_text.append(ocr_result.full_text)
            print("  OCR complete.")
        except Exception as e:
            cli_review.display_error(f"OCR failed: {e}")
            pdf_processor.cleanup_images(all_image_paths)
            return False

    # Process document attachments (extract text directly, no OCR)
    if email_data.doc_paths:
        print(f"\nExtracting text from {len(email_data.doc_paths)} document(s)...")
        for doc_path in email_data.doc_paths:
            try:
                doc_text = gmail_fetcher.extract_text_from_docx(doc_path)
                all_ocr_text.append(doc_text)
                print(f"  Extracted text from {doc_path.name}")
            except Exception as e:
                cli_review.display_error(f"Failed to extract text: {e}")
                pdf_processor.cleanup_images(all_image_paths)
                return False

    combined_ocr_text = "\n\n=== Next Attachment ===\n\n".join(all_ocr_text)

    form_data = field_extractor.extract_fields(combined_ocr_text)
    data_dict = field_extractor.form_data_to_dict(form_data)
    data_dict['_raw_text'] = combined_ocr_text

    try:
        reviewed_data = cli_review.review_and_edit(data_dict)
    except KeyboardInterrupt:
        print("\nCancelled by user.")
        pdf_processor.cleanup_images(all_image_paths)
        gmail_fetcher.cleanup_fetched_files(email_data.attachment_paths)
        return False

    reviewed_data.pop('_raw_text', None)

    payment_types = config.get('field_mappings', {}).get('payment_types', ['Check', 'Debit', 'Amazon'])
    budget_categories = config.get('field_mappings', {}).get('budget_categories', [])
    budget_items = config.get('field_mappings', {}).get('budget_items', [])

    payment_type = cli_review.select_from_list(
        payment_types,
        "Select Payment Type"
    )

    if budget_categories:
        budget_category = cli_review.select_from_list(
            budget_categories,
            "Select Budget Category"
        )
    else:
        budget_category = input("\nEnter Budget Category: ").strip()

    if budget_items:
        budget_item = cli_review.select_from_list(
            budget_items,
            "Select Budget Item"
        )
    else:
        budget_item = input("\nEnter Budget Item: ").strip()

    if dry_run:
        cli_review.display_info("Dry run mode - not writing to spreadsheet")
        print("\nData that would be written:")
        print(f"  Payment Type: {payment_type}")
        print(f"  Budget Category: {budget_category}")
        print(f"  Budget Item: {budget_item}")
        for key, value in reviewed_data.items():
            print(f"  {key}: {value}")
    else:
        if cli_review.confirm_action("Add to spreadsheet?"):
            try:
                writer = sheets_writer.SheetsWriter(
                    credentials_path=config['google_cloud']['credentials_file'],
                    spreadsheet_id=config['google_sheets']['spreadsheet_id'],
                    sheet_name=config['google_sheets'].get('sheet_name', 'Income and Expenses')
                )

                next_id = writer.get_next_id()

                row = sheets_writer.create_spreadsheet_row(
                    form_data=reviewed_data,
                    email_date=email_data.date,
                    budget_category=budget_category,
                    budget_item=budget_item,
                    next_id=next_id,
                    payment_type=payment_type
                )

                row_num = writer.append_row(row)
                cli_review.display_success(
                    f"Added row #{next_id} to \"{config['google_sheets'].get('sheet_name', 'Income and Expenses')}\""
                )

                # Upload attachments to Google Drive if configured
                archive_folder_id = config.get('google_drive', {}).get('archive_folder_id')
                if archive_folder_id and email_data.attachment_paths:
                    if cli_review.confirm_action("Upload attachments to Google Drive?"):
                        try:
                            uploader = drive_uploader.DriveUploader(
                                credentials=fetcher.credentials,
                                archive_folder_id=archive_folder_id
                            )

                            # Use email received date for folder organization
                            if email_data.date:
                                archive_month = email_data.date.strftime('%B')  # e.g., "January"
                            else:
                                archive_month = row.month  # Fallback to form date

                            file_ids = uploader.upload_attachments(
                                file_paths=email_data.attachment_paths,
                                entry_id=next_id,
                                requestor_name=reviewed_data.get('Requestor', ''),
                                month=archive_month
                            )

                            cli_review.display_success(
                                f"Uploaded {len(file_ids)} file(s) to Drive archive ({archive_month.upper()})"
                            )
                        except Exception as e:
                            cli_review.display_error(f"Failed to upload to Drive: {e}")
                            # Continue even if Drive upload fails

                # Offer to print attachments (PDFs and images)
                printable_files = email_data.pdf_paths + email_data.image_paths
                if printable_files:
                    if cli_review.confirm_action("Print attachments?", default=False):
                        try:
                            printers = printer.get_available_printers()
                            default_printer = printer.get_default_printer()

                            if printers:
                                selected = printer.select_printer(printers, default_printer)
                            else:
                                selected = None
                                print("Using system default printer")

                            print(f"\nPrinting {len(printable_files)} file(s)...")
                            success, failed = printer.print_pdfs(printable_files, selected)

                            if success > 0:
                                cli_review.display_success(f"Sent {success} file(s) to printer")
                            if failed > 0:
                                cli_review.display_error(f"Failed to print {failed} file(s)")
                        except Exception as e:
                            cli_review.display_error(f"Print error: {e}")

            except Exception as e:
                cli_review.display_error(f"Failed to write to spreadsheet: {e}")
                return False
        else:
            print("Skipped adding to spreadsheet.")

    pdf_processor.cleanup_images(all_image_paths)
    gmail_fetcher.cleanup_fetched_files(email_data.attachment_paths)

    return True


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="PTA Reimbursement Parser - Extract data from reimbursement request emails"
    )

    parser.add_argument(
        '--config', '-c',
        type=Path,
        default=Path('config/config.yaml'),
        help='Path to config file (default: config/config.yaml)'
    )

    parser.add_argument(
        '--dry-run', '-n',
        action='store_true',
        help="Don't write to Google Sheets, just show what would be done"
    )

    subparsers = parser.add_subparsers(dest='command', help='Available commands')

    # Process single file command
    process_parser = subparsers.add_parser(
        'process',
        help='Process a single .eml file'
    )
    process_parser.add_argument(
        'file',
        type=Path,
        help='Path to the .eml file'
    )

    # Process folder command
    folder_parser = subparsers.add_parser(
        'process-folder',
        help='Process all .eml files in a folder'
    )
    folder_parser.add_argument(
        'folder',
        type=Path,
        help='Path to the folder containing .eml files'
    )

    # Gmail list command
    gmail_list_parser = subparsers.add_parser(
        'gmail-list',
        help='List recent emails from Gmail with attachments'
    )
    gmail_list_parser.add_argument(
        '--query', '-q',
        default='has:attachment',
        help='Gmail search query (default: "has:attachment")'
    )
    gmail_list_parser.add_argument(
        '--max', '-m',
        type=int,
        default=20,
        help='Maximum number of messages to list (default: 20)'
    )

    # Gmail process command
    gmail_process_parser = subparsers.add_parser(
        'gmail-process',
        help='Fetch and process a Gmail message by ID'
    )
    gmail_process_parser.add_argument(
        'message_id',
        help='Gmail message ID (from gmail-list output)'
    )

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    # Load configuration
    try:
        config = load_config(args.config)
    except FileNotFoundError as e:
        cli_review.display_error(str(e))
        sys.exit(1)
    except yaml.YAMLError as e:
        cli_review.display_error(f"Invalid config file: {e}")
        sys.exit(1)

    # Execute command
    try:
        if args.command == 'process':
            if not args.file.exists():
                cli_review.display_error(f"File not found: {args.file}")
                sys.exit(1)

            success = process_eml_file(args.file, config, args.dry_run)
            sys.exit(0 if success else 1)

        elif args.command == 'process-folder':
            if not args.folder.exists():
                cli_review.display_error(f"Folder not found: {args.folder}")
                sys.exit(1)

            successful, failed = process_folder(args.folder, config, args.dry_run)
            sys.exit(0 if failed == 0 else 1)

        elif args.command == 'gmail-list':
            list_gmail_messages(config, args.query, args.max)
            sys.exit(0)

        elif args.command == 'gmail-process':
            success = process_gmail_message(args.message_id, config, args.dry_run)
            sys.exit(0 if success else 1)

    except KeyboardInterrupt:
        print("\nInterrupted by user.")
        sys.exit(130)


if __name__ == '__main__':
    main()
