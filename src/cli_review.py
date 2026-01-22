"""CLI review interface for reviewing and editing extracted form data."""

from typing import Optional


def _safe_print(text: str) -> None:
    """Print text, replacing characters that can't be encoded in the console."""
    try:
        print(text)
    except UnicodeEncodeError:
        # Replace problematic characters with ASCII equivalents
        safe_text = text.encode('ascii', errors='replace').decode('ascii')
        print(safe_text)


def print_table(data: dict[str, str], title: str = "Extracted Data") -> None:
    """
    Print data in a formatted table.

    Args:
        data: Dictionary of field names to values
        title: Title to display above the table
    """
    # Filter out internal fields (starting with _)
    display_data = {k: v for k, v in data.items() if not k.startswith('_')}

    if not display_data:
        print("No data to display.")
        return

    # Calculate column widths
    field_width = max(len(str(k)) for k in display_data.keys()) + 2
    value_width = max(len(str(v)) for v in display_data.values()) + 2
    value_width = min(max(value_width, 20), 50)  # Min 20, max 50

    # Print title
    print(f"\n=== {title} ===")

    # Print top border
    print(f"+-{'-' * field_width}-+-{'-' * value_width}-+")

    # Print header
    print(f"| {'Field':<{field_width}} | {'Value':<{value_width}} |")

    # Print separator
    print(f"+-{'-' * field_width}-+-{'-' * value_width}-+")

    # Print data rows
    for field, value in display_data.items():
        display_value = str(value) if value else "(empty)"
        # Truncate long values
        if len(display_value) > value_width:
            display_value = display_value[:value_width - 3] + "..."
        _safe_print(f"| {field:<{field_width}} | {display_value:<{value_width}} |")

    # Print bottom border
    print(f"+-{'-' * field_width}-+-{'-' * value_width}-+")


def edit_field(data: dict[str, str], field_name: str) -> dict[str, str]:
    """
    Edit a specific field value.

    Args:
        data: Dictionary of field data
        field_name: Name of field to edit

    Returns:
        Updated data dictionary
    """
    # Find matching field (case-insensitive)
    matching_field = None
    for key in data.keys():
        if key.lower() == field_name.lower():
            matching_field = key
            break

    if not matching_field:
        print(f"Field '{field_name}' not found. Available fields:")
        for key in data.keys():
            if not key.startswith('_'):
                print(f"  - {key}")
        return data

    current_value = data[matching_field]
    print(f"\nCurrent value for '{matching_field}': {current_value or '(empty)'}")
    new_value = input("Enter new value (or press Enter to keep current): ").strip()

    if new_value:
        data[matching_field] = new_value
        print(f"Updated '{matching_field}' to: {new_value}")
    else:
        print("Kept existing value.")

    return data


def review_and_edit(data: dict[str, str]) -> dict[str, str]:
    """
    Interactive review and edit loop.

    Args:
        data: Dictionary of extracted field data

    Returns:
        Updated data dictionary after user review
    """
    while True:
        print_table(data)
        print("\nOptions:")
        print("  - Enter a field name to edit it")
        print("  - Type 'ok' or press Enter to continue")
        print("  - Type 'raw' to see raw OCR text (if available)")
        print("  - Type 'quit' to cancel")

        choice = input("\nEdit a field? ").strip().lower()

        if choice in ('ok', ''):
            return data
        elif choice == 'quit':
            raise KeyboardInterrupt("User cancelled")
        elif choice == 'raw':
            if '_raw_text' in data:
                print("\n=== Raw OCR Text ===")
                print(data['_raw_text'])
                print("=== End Raw Text ===\n")
            else:
                print("Raw text not available.")
        else:
            data = edit_field(data, choice)

    return data


def select_from_list(
    options: list[str],
    prompt: str = "Select an option",
    allow_other: bool = True
) -> str:
    """
    Display a numbered list and let user select.

    Args:
        options: List of options to display
        prompt: Prompt text
        allow_other: Whether to allow custom input

    Returns:
        Selected option string
    """
    print(f"\n{prompt}:")
    for i, option in enumerate(options, 1):
        print(f"  {i}. {option}")

    if allow_other:
        print(f"  {len(options) + 1}. Other (enter custom value)")

    while True:
        try:
            choice = input("\n> ").strip()

            # Check if it's a number
            if choice.isdigit():
                idx = int(choice)
                if 1 <= idx <= len(options):
                    return options[idx - 1]
                elif allow_other and idx == len(options) + 1:
                    custom = input("Enter custom value: ").strip()
                    return custom
                else:
                    print(f"Please enter a number between 1 and {len(options) + (1 if allow_other else 0)}")
            else:
                # Allow typing the option directly
                for option in options:
                    if option.lower() == choice.lower():
                        return option
                if allow_other:
                    confirm = input(f"Use '{choice}' as custom value? (y/n): ").strip().lower()
                    if confirm == 'y':
                        return choice
                print("Invalid selection. Please enter a number or valid option.")

        except (ValueError, KeyboardInterrupt):
            if allow_other:
                return ""
            continue


def confirm_action(prompt: str = "Proceed?", default: bool = True) -> bool:
    """
    Ask for yes/no confirmation.

    Args:
        prompt: The confirmation prompt
        default: Default value if user just presses Enter

    Returns:
        True if confirmed, False otherwise
    """
    default_str = "Y/n" if default else "y/N"
    response = input(f"\n{prompt} ({default_str}): ").strip().lower()

    if not response:
        return default
    return response in ('y', 'yes')


def display_success(message: str) -> None:
    """Display a success message."""
    print(f"\n[OK] {message}")


def display_error(message: str) -> None:
    """Display an error message."""
    print(f"\n[ERROR] {message}")


def display_info(message: str) -> None:
    """Display an info message."""
    print(f"\n[INFO] {message}")
