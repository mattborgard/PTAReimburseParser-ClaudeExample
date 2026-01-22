"""PDF processor module for converting PDFs to images."""

import tempfile
from pathlib import Path
from typing import Optional

from pdf2image import convert_from_path
from PIL import Image


def convert_pdf_to_images(
    pdf_path: str | Path,
    dpi: int = 300,
    output_dir: Optional[Path] = None,
    poppler_path: Optional[str | Path] = None
) -> list[Path]:
    """
    Convert a PDF file to a list of images (one per page).

    Args:
        pdf_path: Path to the PDF file
        dpi: Resolution for the output images (default 300 for good OCR quality)
        output_dir: Directory to save images (uses temp dir if not specified)
        poppler_path: Path to Poppler bin directory (optional, uses PATH if not set)

    Returns:
        List of paths to the generated image files

    Raises:
        FileNotFoundError: If the PDF file doesn't exist
        RuntimeError: If PDF conversion fails (e.g., Poppler not installed)
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    # Set up output directory
    if output_dir is None:
        output_dir = Path(tempfile.gettempdir()) / 'pta_parser' / 'images'
    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        # Convert PDF to images
        convert_kwargs = {'dpi': dpi}
        if poppler_path:
            convert_kwargs['poppler_path'] = str(poppler_path)
        images = convert_from_path(pdf_path, **convert_kwargs)
    except Exception as e:
        if 'poppler' in str(e).lower() or 'pdftoppm' in str(e).lower():
            raise RuntimeError(
                "Poppler is not installed or not in PATH. "
                "On Windows, install with: choco install poppler\n"
                "On macOS: brew install poppler\n"
                "On Ubuntu/Debian: apt-get install poppler-utils"
            ) from e
        raise

    # Save images and collect paths
    image_paths = []
    base_name = pdf_path.stem

    for i, image in enumerate(images):
        image_filename = f"{base_name}_page_{i + 1}.png"
        image_path = output_dir / image_filename

        # Handle existing files
        counter = 1
        while image_path.exists():
            image_filename = f"{base_name}_page_{i + 1}_{counter}.png"
            image_path = output_dir / image_filename
            counter += 1

        image.save(image_path, 'PNG')
        image_paths.append(image_path)

    return image_paths


def get_page_count(pdf_path: str | Path) -> int:
    """
    Get the number of pages in a PDF file.

    Args:
        pdf_path: Path to the PDF file

    Returns:
        Number of pages in the PDF
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    # Use pdf2image's page count functionality
    from pdf2image.pdf2image import pdfinfo_from_path

    try:
        info = pdfinfo_from_path(pdf_path)
        return info.get('Pages', 0)
    except Exception:
        # Fallback: convert and count
        images = convert_from_path(pdf_path, dpi=72)  # Low DPI for speed
        return len(images)


def cleanup_images(image_paths: list[Path]) -> None:
    """Remove temporary image files."""
    for path in image_paths:
        try:
            if path.exists():
                path.unlink()
        except OSError:
            pass
