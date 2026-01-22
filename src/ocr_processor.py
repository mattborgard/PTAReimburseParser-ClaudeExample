"""OCR processor module using Google Cloud Vision API."""

import io
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from google.cloud import vision
from PIL import Image


# Vision API limits
MAX_IMAGE_SIZE_BYTES = 20 * 1024 * 1024  # 20MB file size limit
MAX_IMAGE_PIXELS = 40_000_000  # Stay under 75M pixel limit with margin


@dataclass
class OCRResult:
    """Result from OCR processing."""
    full_text: str
    pages: list[str]  # Text per page
    confidence: Optional[float] = None


def initialize_vision_client(credentials_path: Optional[str | Path] = None) -> vision.ImageAnnotatorClient:
    """
    Initialize the Google Cloud Vision client.

    Args:
        credentials_path: Path to service account JSON file.
                         If not provided, uses GOOGLE_APPLICATION_CREDENTIALS env var.

    Returns:
        Configured Vision API client
    """
    if credentials_path:
        credentials_path = Path(credentials_path)
        if not credentials_path.exists():
            raise FileNotFoundError(f"Credentials file not found: {credentials_path}")
        os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = str(credentials_path)

    return vision.ImageAnnotatorClient()


def _compress_image_for_api(
    image_path: Path,
    max_size: int = MAX_IMAGE_SIZE_BYTES,
    max_pixels: int = MAX_IMAGE_PIXELS
) -> bytes:
    """
    Compress an image to fit within Vision API limits.

    Args:
        image_path: Path to the image file
        max_size: Maximum size in bytes
        max_pixels: Maximum total pixels (width * height)

    Returns:
        Compressed image content as bytes
    """
    img = Image.open(image_path)

    # Convert to RGB if necessary (for JPEG compression)
    if img.mode in ('RGBA', 'P'):
        img = img.convert('RGB')

    # First, check if we need to reduce dimensions for pixel count
    current_pixels = img.width * img.height
    if current_pixels > max_pixels:
        # Calculate scale factor to fit within pixel limit
        scale = (max_pixels / current_pixels) ** 0.5
        new_width = int(img.width * scale)
        new_height = int(img.height * scale)
        img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)

    # Now compress to JPEG with progressively lower quality if needed
    quality = 85

    while quality >= 30:
        buffer = io.BytesIO()
        img.save(buffer, format='JPEG', quality=quality, optimize=True)
        content = buffer.getvalue()

        if len(content) <= max_size:
            return content

        quality -= 10

    # If still too large, scale down further
    scale = 0.8
    while scale >= 0.3:
        new_width = int(img.width * scale)
        new_height = int(img.height * scale)
        scaled = img.resize((new_width, new_height), Image.Resampling.LANCZOS)

        buffer = io.BytesIO()
        scaled.save(buffer, format='JPEG', quality=70, optimize=True)
        content = buffer.getvalue()

        if len(content) <= max_size:
            return content

        scale -= 0.1

    # Return whatever we have
    return content


def process_image(
    client: vision.ImageAnnotatorClient,
    image_path: str | Path
) -> str:
    """
    Run OCR on a single image using document text detection.

    Args:
        client: Vision API client
        image_path: Path to the image file

    Returns:
        Extracted text from the image
    """
    image_path = Path(image_path)
    if not image_path.exists():
        raise FileNotFoundError(f"Image file not found: {image_path}")

    # Load and compress image if needed
    content = _compress_image_for_api(image_path)

    image = vision.Image(content=content)

    # Use document text detection for better results on forms
    response = client.document_text_detection(image=image)

    if response.error.message:
        raise RuntimeError(f"Vision API error: {response.error.message}")

    # Extract full text
    if response.full_text_annotation:
        return response.full_text_annotation.text
    elif response.text_annotations:
        return response.text_annotations[0].description
    else:
        return ""


def process_images(
    client: vision.ImageAnnotatorClient,
    image_paths: list[Path]
) -> OCRResult:
    """
    Run OCR on multiple images and combine results.

    Args:
        client: Vision API client
        image_paths: List of paths to image files

    Returns:
        OCRResult with combined text and per-page text
    """
    page_texts = []

    for image_path in image_paths:
        text = process_image(client, image_path)
        page_texts.append(text)

    # Combine all pages
    full_text = '\n\n--- Page Break ---\n\n'.join(page_texts)

    return OCRResult(
        full_text=full_text,
        pages=page_texts
    )


def process_pdf_directly(
    client: vision.ImageAnnotatorClient,
    pdf_path: str | Path
) -> OCRResult:
    """
    Process a PDF directly using Vision API's async document processing.

    Note: This is an alternative to converting to images first.
    For multi-page PDFs, the image-based approach may give better results.

    Args:
        client: Vision API client
        pdf_path: Path to the PDF file

    Returns:
        OCRResult with extracted text
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    # Load PDF content
    with open(pdf_path, 'rb') as f:
        content = f.read()

    # Create input config for PDF
    input_config = vision.InputConfig(
        content=content,
        mime_type='application/pdf'
    )

    # Configure the request
    feature = vision.Feature(type_=vision.Feature.Type.DOCUMENT_TEXT_DETECTION)
    request = vision.AnnotateFileRequest(
        input_config=input_config,
        features=[feature]
    )

    # Process the PDF
    response = client.batch_annotate_files(requests=[request])

    # Extract text from all pages
    page_texts = []
    for file_response in response.responses:
        for page_response in file_response.responses:
            if page_response.error.message:
                raise RuntimeError(f"Vision API error: {page_response.error.message}")

            if page_response.full_text_annotation:
                page_texts.append(page_response.full_text_annotation.text)

    full_text = '\n\n--- Page Break ---\n\n'.join(page_texts)

    return OCRResult(
        full_text=full_text,
        pages=page_texts
    )
