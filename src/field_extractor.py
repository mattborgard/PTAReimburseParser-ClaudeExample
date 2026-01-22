"""Form field extractor module for parsing OCR text into structured data."""

import re
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class FormData:
    """Extracted form field data."""
    requestor: str = ""
    date: str = ""
    amount: str = ""
    email: str = ""
    phone: str = ""
    child_name: str = ""
    teacher_grade: str = ""
    reimbursement_type: str = ""  # Home Room, Teacher, PTA Program
    event: str = ""
    payable_to: str = ""
    delivery: str = ""  # mailbox, send home, pickup
    notes: str = ""
    raw_text: str = ""


def extract_fields(ocr_text: str) -> FormData:
    """
    Extract form fields from OCR text.

    Args:
        ocr_text: Raw text from OCR processing

    Returns:
        FormData with extracted fields
    """
    form_data = FormData(raw_text=ocr_text)

    # Normalize text for easier parsing
    text = ocr_text.replace('\r\n', '\n')

    # Extract each field
    form_data.requestor = _extract_requestor(text)
    form_data.date = _extract_date(text)
    form_data.amount = _extract_amount(text)
    form_data.email = _extract_email(text)
    form_data.phone = _extract_phone(text)
    form_data.child_name = _extract_child_name(text)
    form_data.teacher_grade = _extract_teacher_grade(text)
    form_data.reimbursement_type = _extract_reimbursement_type(text)
    form_data.event = _extract_event(text)
    form_data.payable_to = _extract_payable_to(text)
    form_data.delivery = _extract_delivery(text)

    return form_data


def _extract_requestor(text: str) -> str:
    """Extract the check requestor name."""
    patterns = [
        r'Check\s+Request(?:or|er)[\s:]*([A-Za-z\s]+?)(?:\n|$|Date)',
        r'Request(?:or|er)[\s:]+([A-Za-z\s]+?)(?:\n|$)',
        r'Name[\s:]+([A-Za-z\s]+?)(?:\n|$|Email|Phone)',
        r'Submitted\s+[Bb]y[\s:]+([A-Za-z\s]+?)(?:\n|$)',
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            name = match.group(1).strip()
            # Clean up common OCR artifacts
            name = re.sub(r'\s+', ' ', name)
            if name and len(name) > 1:
                return name

    return ""


def _extract_date(text: str) -> str:
    """Extract the date from the form."""
    patterns = [
        # MM-DD-YYYY or MM/DD/YYYY
        r'Date[\s:]*(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})',
        r'(\d{1,2}[-/]\d{1,2}[-/]\d{4})',
        r'(\d{1,2}[-/]\d{1,2}[-/]\d{2})',
        # Written months
        r'Date[\s:]*([A-Za-z]+\s+\d{1,2},?\s+\d{4})',
        r'([A-Za-z]+\s+\d{1,2},?\s+\d{4})',
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()

    return ""


def _extract_amount(text: str) -> str:
    """Extract the amount requested."""
    patterns = [
        r'Amount\s+Requested[\s:]*\$?([\d,]+\.?\d*)',
        r'Amount[\s:]*\$?([\d,]+\.?\d*)',
        r'Total[\s:]*\$?([\d,]+\.?\d*)',
        r'\$\s*([\d,]+\.\d{2})',
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            amount = match.group(1).strip()
            # Ensure proper formatting
            if amount:
                # Remove commas and format
                amount_clean = amount.replace(',', '')
                try:
                    value = float(amount_clean)
                    return f"${value:.2f}"
                except ValueError:
                    return f"${amount}"

    return ""


def _extract_email(text: str) -> str:
    """Extract email address."""
    # Standard email pattern
    pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    match = re.search(pattern, text)
    if match:
        return match.group(0).lower()
    return ""


def _extract_phone(text: str) -> str:
    """Extract phone number."""
    patterns = [
        # (XXX) XXX-XXXX
        r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}',
        # XXX-XXX-XXXX
        r'\d{3}[-.\s]\d{3}[-.\s]\d{4}',
        # XXXXXXXXXX
        r'(?<!\d)\d{10}(?!\d)',
    ]

    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(0)

    return ""


def _extract_child_name(text: str) -> str:
    """Extract child's name."""
    patterns = [
        r"Child(?:'s)?\s+Name[\s:]+([A-Za-z\s]+?)(?:\n|$|Teacher|Grade)",
        r"Student(?:'s)?\s+Name[\s:]+([A-Za-z\s]+?)(?:\n|$|Teacher|Grade)",
        r"Child[\s:]+([A-Za-z\s]+?)(?:\n|$)",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            name = match.group(1).strip()
            name = re.sub(r'\s+', ' ', name)
            if name and len(name) > 1:
                return name

    return ""


def _extract_teacher_grade(text: str) -> str:
    """Extract teacher name and/or grade.

    Captures the full field value for patterns like:
    - "Mrs. Lanford 5th"
    - "K Michaud" (K for kindergarten)
    - "McCord / 3rd"
    - "5th - Johnson"
    - "Teacher/Grade: Mrs, Lanford - Kindergarten"
    """
    # First, try to find the combined Teacher/Grade field - most reliable
    # This captures everything after "Teacher/Grade:" until a clear field boundary
    combined_pattern = r'Teacher\s*/\s*Grade[\s:]+([^\n]+?)(?=\n\s*(?:Email|Phone|Child|Event|Amount|Payable|Delivery|Reimbursement)|$|\n\n)'
    match = re.search(combined_pattern, text, re.IGNORECASE)
    if match:
        result = match.group(1).strip()
        result = re.sub(r'\s+', ' ', result)
        if result and len(result) > 1:
            return result

    # Try to find separate Teacher and Grade fields and combine them
    teacher_match = re.search(r'\bTeacher[\s:]+([A-Za-z][A-Za-z.,\s]+?)(?=\n|$)', text, re.IGNORECASE)
    grade_match = re.search(r'\bGrade[\s:]+([A-Za-z0-9][A-Za-z0-9\s]+?)(?=\n|$)', text, re.IGNORECASE)

    if teacher_match and grade_match:
        teacher = teacher_match.group(1).strip()
        grade = grade_match.group(1).strip()
        # Clean up each part
        teacher = re.sub(r'\s+', ' ', teacher)
        grade = re.sub(r'\s+', ' ', grade)
        if teacher and grade:
            return f"{teacher} - {grade}"

    # Try just Teacher field
    if teacher_match:
        result = teacher_match.group(1).strip()
        result = re.sub(r'\s+', ' ', result)
        if result and len(result) > 1:
            return result

    # Try just Grade field
    if grade_match:
        result = grade_match.group(1).strip()
        result = re.sub(r'\s+', ' ', result)
        if result and len(result) > 1:
            return result

    # Fallback: look for common teacher/grade patterns anywhere
    # Pattern for "Mrs./Mr./Ms. LastName" followed by grade
    teacher_grade_pattern = r'((?:Mrs?\.?|Ms\.?|Miss)\s+[A-Z][a-z]+)[\s,/-]*((?:Pre-?)?K(?:indergarten)?|[1-5](?:st|nd|rd|th)?(?:\s*grade)?)'
    match = re.search(teacher_grade_pattern, text, re.IGNORECASE)
    if match:
        teacher = match.group(1).strip()
        grade = match.group(2).strip()
        return f"{teacher} - {grade}"

    # Fallback: grade followed by teacher name
    grade_teacher_pattern = r'\b((?:Pre-?)?K(?:indergarten)?|[1-5](?:st|nd|rd|th)?)\b[\s,/-]+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)'
    match = re.search(grade_teacher_pattern, text)
    if match:
        grade = match.group(1)
        teacher = match.group(2)
        return f"{grade} {teacher}".strip()

    return ""


def _extract_reimbursement_type(text: str) -> str:
    """Extract reimbursement type (Home Room, Teacher, PTA Program)."""
    text_lower = text.lower()

    # Look for checkbox indicators or explicit mentions
    type_patterns = [
        (r'(?:☑|✓|✔|x|\[x\])\s*home\s*room', 'Home Room Parent'),
        (r'(?:☑|✓|✔|x|\[x\])\s*teacher', 'Teacher'),
        (r'(?:☑|✓|✔|x|\[x\])\s*pta\s*program', 'PTA Program'),
        (r'home\s*room\s*parent\s*reimbursement', 'Home Room Parent'),
        (r'teacher\s*reimbursement', 'Teacher'),
        (r'pta\s*program\s*reimbursement', 'PTA Program'),
        (r'reimbursement\s*type[\s:]*home\s*room', 'Home Room Parent'),
        (r'reimbursement\s*type[\s:]*teacher', 'Teacher'),
        (r'reimbursement\s*type[\s:]*pta', 'PTA Program'),
    ]

    for pattern, type_name in type_patterns:
        if re.search(pattern, text_lower):
            return type_name

    return ""


def _extract_event(text: str) -> str:
    """Extract event name."""
    patterns = [
        r'Event[\s:]+([A-Za-z\s]+?)(?:\n|$|Amount)',
        r'For[\s:]+([A-Za-z\s]+?(?:Party|Event|Activity))(?:\n|$)',
        r'Purpose[\s:]+([A-Za-z\s]+?)(?:\n|$)',
        # Common event names
        r'(Winter\s+Party)',
        r'(Fall\s+Party)',
        r'(Spring\s+Party)',
        r'(Valentine(?:\'?s)?\s+(?:Day\s+)?Party)',
        r'(Halloween\s+Party)',
        r'(End\s+of\s+Year\s+Party)',
        r'(Field\s+Day)',
        r'(Teacher\s+Appreciation)',
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            event = match.group(1).strip() if match.lastindex else match.group(0).strip()
            event = re.sub(r'^Event[\s:]+', '', event, flags=re.IGNORECASE)
            event = re.sub(r'^For[\s:]+', '', event, flags=re.IGNORECASE)
            if event and len(event) > 2:
                return event

    return ""


def _extract_payable_to(text: str) -> str:
    """Extract 'Make Check Payable To' name."""
    patterns = [
        r'Make\s+Check\s+Payable\s+To[\s:]+([A-Za-z\s]+?)(?:\n|$)',
        r'Payable\s+To[\s:]+([A-Za-z\s]+?)(?:\n|$)',
        r'Pay\s+To[\s:]+([A-Za-z\s]+?)(?:\n|$)',
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            name = match.group(1).strip()
            name = re.sub(r'\s+', ' ', name)
            if name and len(name) > 1:
                return name

    return ""


def _extract_delivery(text: str) -> str:
    """Extract delivery preference."""
    text_lower = text.lower()

    delivery_patterns = [
        (r'(?:☑|✓|✔|x|\[x\])\s*(?:teacher\'?s?\s*)?mailbox', 'Teacher mailbox'),
        (r'(?:☑|✓|✔|x|\[x\])\s*send\s*home\s*with\s*child', 'Send home with child'),
        (r'(?:☑|✓|✔|x|\[x\])\s*(?:i\'?ll\s*)?pick\s*(?:it\s*)?up', 'Pickup'),
        (r'mailbox', 'Teacher mailbox'),
        (r'send\s*home', 'Send home with child'),
        (r'pick\s*up', 'Pickup'),
    ]

    for pattern, delivery_type in delivery_patterns:
        if re.search(pattern, text_lower):
            return delivery_type

    return ""


def form_data_to_dict(form_data: FormData) -> dict:
    """Convert FormData to dictionary (excluding raw_text)."""
    return {
        'Requestor': form_data.requestor,
        'Date': form_data.date,
        'Amount': form_data.amount,
        'Email': form_data.email,
        'Phone': form_data.phone,
        'Child': form_data.child_name,
        'Teacher/Grade': form_data.teacher_grade,
        'Type': form_data.reimbursement_type,
        'Event': form_data.event,
        'Payable To': form_data.payable_to,
        'Delivery': form_data.delivery,
    }
