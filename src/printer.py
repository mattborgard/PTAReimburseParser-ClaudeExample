"""Printer module for printing PDF attachments."""

import os
import subprocess
import sys
from pathlib import Path
from typing import Optional


def get_available_printers() -> list[str]:
    """
    Get list of available printers on Windows.

    Returns:
        List of printer names
    """
    if sys.platform != 'win32':
        return []

    try:
        import win32print
        printers = []
        for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
            printers.append(printer[2])  # printer[2] is the printer name
        return printers
    except ImportError:
        # Fallback: try to get printers via PowerShell
        try:
            result = subprocess.run(
                ['powershell', '-Command', 'Get-Printer | Select-Object -ExpandProperty Name'],
                capture_output=True,
                text=True,
                timeout=10
            )
            if result.returncode == 0:
                return [p.strip() for p in result.stdout.strip().split('\n') if p.strip()]
        except Exception:
            pass
        return []


def get_default_printer() -> Optional[str]:
    """
    Get the default printer name.

    Returns:
        Default printer name or None
    """
    if sys.platform != 'win32':
        return None

    try:
        import win32print
        return win32print.GetDefaultPrinter()
    except ImportError:
        # Fallback via PowerShell
        try:
            result = subprocess.run(
                ['powershell', '-Command',
                 'Get-CimInstance -ClassName Win32_Printer | Where-Object {$_.Default -eq $true} | Select-Object -ExpandProperty Name'],
                capture_output=True,
                text=True,
                timeout=10
            )
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout.strip()
        except Exception:
            pass
        return None


def print_pdf(file_path: Path, printer_name: Optional[str] = None) -> bool:
    """
    Print a PDF file.

    Args:
        file_path: Path to the PDF file
        printer_name: Name of printer (uses default if None)

    Returns:
        True if print job was submitted successfully
    """
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    if sys.platform != 'win32':
        # On non-Windows, try lpr
        try:
            cmd = ['lpr', str(file_path)]
            if printer_name:
                cmd = ['lpr', '-P', printer_name, str(file_path)]
            subprocess.run(cmd, check=True, timeout=30)
            return True
        except Exception:
            return False

    # Windows printing
    try:
        import win32api
        import win32print

        if printer_name is None:
            printer_name = win32print.GetDefaultPrinter()

        # Use ShellExecute to print via the default PDF handler
        win32api.ShellExecute(
            0,
            "print",
            str(file_path),
            f'/d:"{printer_name}"',
            ".",
            0  # SW_HIDE
        )
        return True

    except ImportError:
        # Fallback: use PowerShell Start-Process with -Verb Print
        try:
            if printer_name:
                # Set temporary default printer, print, then restore
                # This is a workaround without pywin32
                cmd = [
                    'powershell', '-Command',
                    f'Start-Process -FilePath "{file_path}" -Verb Print -PassThru | Wait-Process -Timeout 60'
                ]
            else:
                cmd = [
                    'powershell', '-Command',
                    f'Start-Process -FilePath "{file_path}" -Verb Print'
                ]

            subprocess.run(cmd, timeout=30)
            return True
        except Exception as e:
            print(f"Print error: {e}")
            return False


def print_pdfs(
    file_paths: list[Path],
    printer_name: Optional[str] = None,
    confirm_each: bool = False
) -> tuple[int, int]:
    """
    Print multiple PDF files.

    Args:
        file_paths: List of PDF file paths
        printer_name: Printer name (uses default if None)
        confirm_each: If True, confirm before each file

    Returns:
        Tuple of (successful_count, failed_count)
    """
    successful = 0
    failed = 0

    for file_path in file_paths:
        if confirm_each:
            response = input(f"Print {file_path.name}? (y/n): ").strip().lower()
            if response != 'y':
                continue

        try:
            if print_pdf(file_path, printer_name):
                successful += 1
                print(f"  Sent to printer: {file_path.name}")
            else:
                failed += 1
                print(f"  Failed to print: {file_path.name}")
        except Exception as e:
            failed += 1
            print(f"  Error printing {file_path.name}: {e}")

    return (successful, failed)


def select_printer(printers: list[str], default: Optional[str] = None) -> Optional[str]:
    """
    Let user select a printer from a list.

    Args:
        printers: List of printer names
        default: Default printer name

    Returns:
        Selected printer name or None to use default
    """
    if not printers:
        print("No printers found.")
        return None

    print("\nAvailable printers:")
    for i, printer in enumerate(printers, 1):
        default_marker = " (default)" if printer == default else ""
        print(f"  {i}. {printer}{default_marker}")

    print(f"  {len(printers) + 1}. Use system default")

    while True:
        try:
            choice = input("\nSelect printer: ").strip()
            if not choice:
                return default

            idx = int(choice)
            if 1 <= idx <= len(printers):
                return printers[idx - 1]
            elif idx == len(printers) + 1:
                return default
            else:
                print(f"Please enter 1-{len(printers) + 1}")
        except ValueError:
            # Maybe they typed the printer name
            for printer in printers:
                if choice.lower() in printer.lower():
                    return printer
            print("Invalid selection")
