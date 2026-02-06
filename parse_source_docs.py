#!/usr/bin/env python3
"""
Source Document Parser
======================
Parses tax source documents (W-2, 1099-INT, 1099-DIV, 1099-R, SSA-1099, 1098)
directly from PDFs and outputs JSON compatible with populate_checksheet.py.

This fills the SOURCE column. Use parse_input_listing.py for the CCH column.

Usage:
    python parse_source_docs.py "L:\\Client Folder\\2025"
    python parse_source_docs.py "L:\\Client Folder\\2025" --output data.json
    python parse_source_docs.py "L:\\Client Folder\\2025" --checksheet "C:\\Tax\\checksheet.xlsx"
"""

import argparse
import json
import os
import re
import sys
from pathlib import Path
from datetime import datetime

try:
    import pdfplumber
except ImportError:
    print("ERROR: pdfplumber not found. Install it:")
    print("  pip install pdfplumber")
    sys.exit(1)

# OCR support (optional)
OCR_AVAILABLE = False
POPPLER_PATH = None
try:
    from pdf2image import convert_from_path
    import pytesseract

    # Set Tesseract path for Windows - check common locations
    tesseract_paths = [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Users\ngroom\AppData\Local\Programs\Tesseract-OCR\tesseract.exe',
    ]
    for tp in tesseract_paths:
        if os.path.exists(tp):
            pytesseract.pytesseract.tesseract_cmd = tp
            break

    # Set Poppler path for Windows
    poppler_paths = [
        r'C:\Users\ngroom\Downloads\Release-25.12.0-0\poppler-25.12.0\Library\bin',
        r'C:\Program Files\poppler\Library\bin',
    ]
    for pp in poppler_paths:
        if os.path.exists(pp):
            POPPLER_PATH = pp
            break

    OCR_AVAILABLE = True
except ImportError:
    pass


def extract_text_from_pdf(pdf_path, use_ocr=True):
    """Extract all text from a PDF, using OCR if needed."""
    text = ""
    is_scanned = False

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                text += page_text + "\n"
    except Exception as e:
        print(f"  Warning: Could not read {pdf_path}: {e}")
        return "", False

    # Check if we got meaningful text
    if len(text.strip()) < 50:
        is_scanned = True
        if use_ocr and OCR_AVAILABLE:
            print(f"  Attempting OCR...")
            text = extract_text_with_ocr(pdf_path)

    return text, is_scanned


def extract_text_with_ocr(pdf_path):
    """Extract text from a scanned PDF using OCR."""
    if not OCR_AVAILABLE:
        return ""

    text = ""
    try:
        # Convert PDF pages to images
        if POPPLER_PATH:
            images = convert_from_path(pdf_path, dpi=300, poppler_path=POPPLER_PATH)
        else:
            images = convert_from_path(pdf_path, dpi=300)

        for i, image in enumerate(images):
            # Run OCR on each page
            page_text = pytesseract.image_to_string(image)
            text += page_text + "\n"

    except Exception as e:
        print(f"  OCR error: {e}")

    return text


def extract_amount(text, patterns, default=0):
    """Extract a dollar amount using multiple regex patterns."""
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            try:
                amt_str = match.group(1).replace(',', '').replace('$', '').strip()
                return float(amt_str)
            except (ValueError, IndexError):
                continue
    return default


def extract_amount_with_quality(text, patterns, field_name, default=0):
    """
    Extract a dollar amount with quality tracking.

    Returns:
        tuple: (value, quality_info)
        quality_info = {
            'field': field_name,
            'found': bool,
            'pattern_attempts': int (1=first pattern matched, higher=less confident),
            'confidence': int (0-100),
            'raw_match': str or None,
            'issues': list of strings
        }
    """
    quality = {
        'field': field_name,
        'found': False,
        'pattern_attempts': 0,
        'confidence': 0,
        'raw_match': None,
        'issues': []
    }

    for attempt, pattern in enumerate(patterns, 1):
        quality['pattern_attempts'] = attempt
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            try:
                raw = match.group(1)
                quality['raw_match'] = raw
                amt_str = raw.replace(',', '').replace('$', '').strip()
                value = float(amt_str)
                quality['found'] = True

                # Calculate confidence based on pattern position and value sanity
                base_confidence = max(100 - (attempt - 1) * 15, 40)  # First pattern = 100%, decreases

                # Validate the value
                if value < 0:
                    quality['issues'].append('Negative value')
                    base_confidence -= 30
                if value > 10000000:  # Over 10 million seems suspicious
                    quality['issues'].append('Unusually high value')
                    base_confidence -= 20
                if '.' in amt_str and len(amt_str.split('.')[-1]) != 2:
                    quality['issues'].append('Unusual decimal places')
                    base_confidence -= 10

                quality['confidence'] = max(base_confidence, 10)
                return value, quality

            except (ValueError, IndexError):
                continue

    # Not found
    quality['confidence'] = 0
    return default, quality


def add_quality_metadata(data, form_type, is_ocr=False, required_fields=None):
    """
    Add quality metadata to a parsed data dict.

    Args:
        data: Dict of parsed form data
        form_type: Type of form (W-2, 1099-INT, etc.)
        is_ocr: Whether the source was OCR
        required_fields: List of required field names

    Returns:
        data dict with _quality metadata added
    """
    if required_fields is None:
        required_fields = []

    issues = []
    missing_required = []
    field_quality = {}

    # Check each field
    for field, value in data.items():
        if field in ('source_file', '_quality'):
            continue

        is_found = value is not None and value != '' and value != 0
        confidence = 85 if is_found else 0

        # Reduce confidence for OCR
        if is_ocr and is_found:
            confidence = min(confidence, 70)

        # Check for suspicious values
        field_issues = []
        if isinstance(value, (int, float)):
            if value < 0:
                field_issues.append('Negative value')
                confidence -= 30
            if value > 10000000:
                field_issues.append('Unusually high')
                confidence -= 20

        field_quality[field] = {
            'found': is_found,
            'confidence': max(confidence, 0),
            'issues': field_issues
        }

        if field_issues:
            issues.extend([f"{field}: {i}" for i in field_issues])

        # Check required
        if field in required_fields and not is_found:
            missing_required.append(field)
            issues.append(f"Missing required: {field}")

    # Calculate overall confidence
    confidences = [q['confidence'] for q in field_quality.values() if q['found']]
    overall = sum(confidences) // len(confidences) if confidences else 0

    if is_ocr:
        overall = min(overall, 75)
    if missing_required:
        overall -= len(missing_required) * 10

    data['_quality'] = {
        'is_ocr': is_ocr,
        'overall_confidence': max(overall, 0),
        'issues': issues,
        'missing_required': missing_required,
        'field_quality': field_quality
    }

    return data


class ExtractionResult:
    """Container for extraction results with quality tracking."""

    def __init__(self, form_type, source_file, is_ocr=False):
        self.form_type = form_type
        self.source_file = source_file
        self.is_ocr = is_ocr
        self.data = {}
        self.quality = {}
        self.overall_confidence = 100
        self.issues = []
        self.missing_required = []
        self.math_errors = []

    def add_field(self, field_name, value, quality_info=None):
        """Add a field with optional quality info."""
        self.data[field_name] = value
        if quality_info:
            self.quality[field_name] = quality_info
            if quality_info.get('issues'):
                self.issues.extend([f"{field_name}: {i}" for i in quality_info['issues']])

    def add_text_field(self, field_name, value, confidence=80):
        """Add a text field (non-numeric)."""
        self.data[field_name] = value
        self.quality[field_name] = {
            'field': field_name,
            'found': bool(value),
            'confidence': confidence if value else 0,
            'issues': [] if value else ['Not found']
        }

    def check_required(self, required_fields):
        """Check for missing required fields."""
        for field in required_fields:
            if field not in self.data or not self.data[field]:
                self.missing_required.append(field)
                self.issues.append(f"Missing required: {field}")

    def check_math(self, formula_name, expected, actual, tolerance=0.01):
        """Validate math relationships between fields."""
        if abs(expected - actual) > tolerance:
            self.math_errors.append({
                'check': formula_name,
                'expected': expected,
                'actual': actual,
                'difference': actual - expected
            })
            self.issues.append(f"Math error in {formula_name}: expected {expected:.2f}, got {actual:.2f}")

    def calculate_overall_confidence(self):
        """Calculate overall extraction confidence."""
        if not self.quality:
            self.overall_confidence = 50
            return

        confidences = [q.get('confidence', 50) for q in self.quality.values() if q.get('found')]
        if confidences:
            self.overall_confidence = sum(confidences) // len(confidences)
        else:
            self.overall_confidence = 0

        # Reduce confidence for issues
        if self.is_ocr:
            self.overall_confidence = min(self.overall_confidence, 75)
        if self.missing_required:
            self.overall_confidence -= len(self.missing_required) * 10
        if self.math_errors:
            self.overall_confidence -= len(self.math_errors) * 15

        self.overall_confidence = max(self.overall_confidence, 0)

    def to_dict(self):
        """Convert to dictionary for JSON output."""
        self.calculate_overall_confidence()
        result = dict(self.data)
        result['source_file'] = self.source_file
        result['_quality'] = {
            'is_ocr': self.is_ocr,
            'overall_confidence': self.overall_confidence,
            'field_quality': self.quality,
            'issues': self.issues,
            'missing_required': self.missing_required,
            'math_errors': self.math_errors
        }
        return result


def identify_document_type(text, filename):
    """Identify what type of tax document this is."""
    text_upper = text.upper()
    fname_upper = filename.upper()

    # Check filename first - ORDER MATTERS: check specific forms before generic ones
    if 'W2' in fname_upper or 'W-2' in fname_upper:
        return 'W-2'
    if '1099-INT' in fname_upper or '1099INT' in fname_upper:
        return '1099-INT'
    if '1099-DIV' in fname_upper or '1099DIV' in fname_upper:
        return '1099-DIV'
    if '1099-R' in fname_upper or '1099R' in fname_upper:
        return '1099-R'
    if '1099-B' in fname_upper or '1099B' in fname_upper:
        return '1099-B'
    if '1099-Q' in fname_upper or '1099Q' in fname_upper:
        return '1099-Q'
    if 'SSA-1099' in fname_upper or 'SSA1099' in fname_upper or 'SSI' in fname_upper:
        return 'SSA-1099'
    # Check 1098-T BEFORE generic 1098
    if '1098-T' in fname_upper or '1098T' in fname_upper:
        return '1098-T'
    if '1098' in fname_upper:
        return '1098'
    # K-1 by filename
    if 'K-1' in fname_upper or 'K1' in fname_upper:
        return 'K-1'

    # Check content - check for consolidated forms FIRST (before individual types)
    # because consolidated docs contain multiple 1099 types
    if 'CONSOLIDATED' in text_upper and '1099' in text_upper:
        return 'CONSOLIDATED-1099'
    # Vanguard/Fidelity consolidated 1099 - contains multiple form types
    if 'TAX INFORMATION STATEMENT' in text_upper and ('1099-DIV' in text_upper or '1099-INT' in text_upper):
        return 'CONSOLIDATED-1099'
    if 'VANGUARD' in text_upper and ('1099-DIV' in text_upper or 'DIVIDENDS AND DISTRIBUTIONS' in text_upper):
        return 'CONSOLIDATED-1099'
    if 'FIDELITY' in text_upper and ('1099-DIV' in text_upper or '1099-INT' in text_upper or '1099-B' in text_upper):
        return 'CONSOLIDATED-1099'
    # Multi-form documents with both DIV and INT sections
    if '1099-DIV' in text_upper and '1099-INT' in text_upper:
        return 'CONSOLIDATED-1099'

    # Now check individual form types
    if 'WAGE AND TAX STATEMENT' in text_upper or 'FORM W-2' in text_upper:
        return 'W-2'
    if 'INTEREST INCOME' in text_upper and '1099-INT' in text_upper:
        return '1099-INT'
    if 'DIVIDENDS AND DISTRIBUTIONS' in text_upper or '1099-DIV' in text_upper:
        return '1099-DIV'
    if 'DISTRIBUTIONS FROM PENSIONS' in text_upper or '1099-R' in text_upper:
        return '1099-R'
    if 'PROCEEDS FROM BROKER' in text_upper or '1099-B' in text_upper:
        return '1099-B'
    if 'SOCIAL SECURITY BENEFIT STATEMENT' in text_upper or 'SSA-1099' in text_upper:
        return 'SSA-1099'
    if 'MORTGAGE INTEREST STATEMENT' in text_upper or ('1098' in text_upper and 'MORTGAGE' in text_upper):
        return '1098'
    # 1098-T Tuition Statement (content-based)
    if '1098-T' in text_upper or 'TUITION STATEMENT' in text_upper:
        return '1098-T'
    if 'FILER IS AN ELIGIBLE EDUCATIONAL INSTITUTION' in text_upper:
        return '1098-T'

    # 1099-Q Qualified Education Program (content-based)
    if '1099-Q' in text_upper or 'QUALIFIED EDUCATION PROGRAM' in text_upper:
        return '1099-Q'
    if 'PAYMENTS FROM QUALIFIED EDUCATION' in text_upper:
        return '1099-Q'

    # Schedule K-1 (content-based) - check BEFORE PROPERTY-TAX since K-1s may contain "real estate"
    if 'SCHEDULE K-1' in text_upper:
        return 'K-1'
    if "PARTNER'S SHARE OF INCOME" in text_upper or "SHAREHOLDER'S SHARE OF INCOME" in text_upper:
        return 'K-1'
    if "BENEFICIARY'S SHARE OF INCOME" in text_upper:
        return 'K-1'
    if 'FORM 1065' in text_upper or 'FORM 1120S' in text_upper or 'FORM 1120-S' in text_upper:
        return 'K-1'

    # Property Tax - check AFTER K-1 to avoid false positives
    if 'AD VALOREM' in text_upper or 'PROPERTY TAX' in text_upper or 'TAX COLLECTOR' in text_upper:
        return 'PROPERTY-TAX'
    if 'REAL ESTATE' in text_upper and 'TAX' in text_upper and 'ASSESSMENT' in text_upper:
        return 'PROPERTY-TAX'

    return 'UNKNOWN'


def parse_w2(text, filename, is_ocr=False):
    """Parse W-2 wage statement with quality tracking."""
    result = ExtractionResult('W-2', filename, is_ocr)

    # Initialize data dict for backward compatibility
    data = {
        'employer_name': '',
        'box1_wages': 0,
        'box2_fed_withholding': 0,
        'box3_ss_wages': 0,
        'box4_ss_tax': 0,
        'box5_medicare_wages': 0,
        'box6_medicare_tax': 0,
        'box16_state_wages': 0,
        'box17_state_withholding': 0,
        'source_file': filename
    }

    # W-2 formats vary. Common patterns:
    # Format A (Justworks): Labels on one line, values on next line
    #   "1 Wages, tips, other compensation 2 Federal income tax withheld"
    #   "47676.89 2459.16"
    # Format B: Labels followed by EIN then values
    #   "1 Wages... 2 Federal..."
    #   "92-3246447 28665.38 251.85"

    # Box 1 & 2: Try multiple patterns
    # Pattern 1: Values on line immediately after labels (no EIN)
    box12_match = re.search(
        r'1\s+[Ww]ages.*?2\s+[Ff]ederal.*?withheld\s*[\n\r]+([\d,]+\.?\d{2})\s+([\d,]+\.?\d{2})',
        text, re.DOTALL
    )
    if not box12_match:
        # Pattern 2: EIN followed by values
        box12_match = re.search(
            r'1\s+[Ww]ages.*?2\s+[Ff]ederal.*?[\n\r]+[\d-]+\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)',
            text, re.DOTALL
        )
    if box12_match:
        data['box1_wages'] = float(box12_match.group(1).replace(',', ''))
        data['box2_fed_withholding'] = float(box12_match.group(2).replace(',', ''))

    # Box 3 & 4: SS wages and SS tax
    # Pattern 1: Values on line immediately after labels
    box34_match = re.search(
        r'3\s+[Ss]ocial\s+[Ss]ecurity\s+[Ww]ages.*?4\s+[Ss]ocial\s+[Ss]ecurity\s+[Tt]ax.*?withheld\s*[\n\r]+([\d,]+\.?\d{2})\s+([\d,]+\.?\d{2})',
        text, re.DOTALL
    )
    if not box34_match:
        # Pattern 2: With employer name or other text between
        box34_match = re.search(
            r'3\s+[Ss]ocial\s+[Ss]ecurity\s+[Ww]ages.*?4\s+[Ss]ocial\s+[Ss]ecurity\s+[Tt]ax.*?[\n\r]+([A-Z][A-Za-z\s]+)?[\n\r]*([\d,]+\.?\d*)\s+([\d,]+\.?\d*)',
            text, re.DOTALL
        )
        if box34_match:
            data['box3_ss_wages'] = float(box34_match.group(2).replace(',', ''))
            data['box4_ss_tax'] = float(box34_match.group(3).replace(',', ''))
    if box34_match and data['box3_ss_wages'] == 0:
        data['box3_ss_wages'] = float(box34_match.group(1).replace(',', ''))
        data['box4_ss_tax'] = float(box34_match.group(2).replace(',', ''))

    # Box 5 & 6: Medicare wages and tax
    # Pattern 1: Values on line immediately after labels (Justworks format - no EIN prefix)
    box56_match = re.search(
        r'5\s+[Mm]edicare\s+[Ww]ages.*?6\s+[Mm]edicare\s+[Tt]ax.*?withheld\s*[\n\r]+([\d,]+\.?\d{2})\s+([\d,]+\.?\d{2})',
        text, re.DOTALL
    )
    if not box56_match:
        # Pattern 2: EIN followed by values (Rippling format)
        box56_match = re.search(
            r'5\s+[Mm]edicare\s+[Ww]ages.*?6\s+[Mm]edicare\s+[Tt]ax.*?[\n\r]+[\d-]+\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)',
            text, re.DOTALL
        )
    if not box56_match:
        # Pattern 3: Generic fallback - values on next line
        box56_match = re.search(
            r'5\s+[Mm]edicare\s+[Ww]ages.*?6\s+[Mm]edicare\s+[Tt]ax.*?[\n\r]+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)',
            text, re.DOTALL
        )
    if box56_match:
        data['box5_medicare_wages'] = float(box56_match.group(1).replace(',', ''))
        data['box6_medicare_tax'] = float(box56_match.group(2).replace(',', ''))

    # Validate and correct potential Box 3/4 vs 5/6 mix-ups using tax rate math
    # SS tax rate is 6.2%, Medicare rate is 1.45%
    if data['box3_ss_wages'] > 0 and data['box4_ss_tax'] > 0:
        ss_rate = data['box4_ss_tax'] / data['box3_ss_wages']
        # If Box 3/4 looks like Medicare rate (~1.45%), they're probably swapped
        if 0.012 < ss_rate < 0.018:  # Medicare rate range (1.2% - 1.8%)
            # Box 3/4 has Medicare values, Box 5/6 might have wrong values
            actual_medicare_wages = data['box3_ss_wages']
            actual_medicare_tax = data['box4_ss_tax']

            # Check if Box 5/6 looks like SS rate
            if data['box5_medicare_wages'] > 0 and data['box6_medicare_tax'] > 0:
                med_rate = data['box6_medicare_tax'] / data['box5_medicare_wages']
                if 0.055 < med_rate < 0.070:  # SS rate range (5.5% - 7%)
                    # Box 5/6 actually has SS values - full swap
                    data['box3_ss_wages'] = data['box5_medicare_wages']
                    data['box4_ss_tax'] = data['box6_medicare_tax']
                    data['box5_medicare_wages'] = actual_medicare_wages
                    data['box6_medicare_tax'] = actual_medicare_tax
                else:
                    # Box 5/6 doesn't have SS values - just fix Medicare
                    # SS wages often equal Medicare wages when under wage base
                    data['box5_medicare_wages'] = actual_medicare_wages
                    data['box6_medicare_tax'] = actual_medicare_tax
                    # Calculate expected SS tax if wages are reasonable
                    data['box4_ss_tax'] = round(actual_medicare_wages * 0.062, 2)
            else:
                # No Box 5/6 values found - copy Medicare values there
                data['box5_medicare_wages'] = actual_medicare_wages
                data['box6_medicare_tax'] = actual_medicare_tax
                data['box4_ss_tax'] = round(actual_medicare_wages * 0.062, 2)

    # Additional check: if Box 5/6 equals Box 1/2, the Medicare pattern failed
    # In this case, try to use Box 3 values for Medicare (if they look correct)
    if (abs(data['box5_medicare_wages'] - data['box1_wages']) < 0.01 and
        abs(data['box6_medicare_tax'] - data['box2_fed_withholding']) < 0.01):
        # Box 5/6 pattern failed and matched Box 1/2
        if data['box3_ss_wages'] > 0 and data['box4_ss_tax'] > 0:
            rate34 = data['box4_ss_tax'] / data['box3_ss_wages']
            if 0.012 < rate34 < 0.018:  # Medicare rate in Box 3/4
                # Use Box 3/4 for Medicare, calculate SS
                data['box5_medicare_wages'] = data['box3_ss_wages']
                data['box6_medicare_tax'] = data['box4_ss_tax']
                # SS wages typically same as Medicare wages, calc SS tax
                data['box4_ss_tax'] = round(data['box3_ss_wages'] * 0.062, 2)
            elif 0.055 < rate34 < 0.070:  # SS rate in Box 3/4 (correct)
                # Box 3/4 is correct, but Box 5/6 pattern failed
                # Medicare wages typically equal SS wages when under wage base
                data['box5_medicare_wages'] = data['box3_ss_wages']
                data['box6_medicare_tax'] = round(data['box3_ss_wages'] * 0.0145, 2)

    # Employer name - look for line after "c Employer's name" or company name patterns
    employer_patterns = [
        r"[Cc]\s+[Ee]mployer.*?name.*?[\n\r]+([A-Z][A-Z0-9\s\.,&-]+(?:LLC|INC|CORP|CO)?)",
        r"[Ee]mployer.*?name.*?ZIP.*?[\n\r]+([A-Z][A-Z0-9\s\.,&-]+(?:LLC|INC|CORP|CO)?)",
    ]
    for pattern in employer_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            name = match.group(1).strip()
            # Clean up - stop at newline or address indicator
            name = re.split(r'[\n\r]|\d{4,}', name)[0].strip()
            if len(name) > 3:
                data['employer_name'] = name[:60]
                break

    # Fallback: look for LLC/INC/CORP in text
    if not data['employer_name']:
        company_match = re.search(r'([A-Z][A-Z0-9\s&-]+(?:LLC|INC|CORP|COMPANY|CO\.))', text)
        if company_match:
            data['employer_name'] = company_match.group(1).strip()[:60]

    # Fallback: filename
    if not data['employer_name']:
        fname_match = re.search(r'W-?2[-_]?([A-Za-z]+)', filename, re.IGNORECASE)
        if fname_match:
            data['employer_name'] = fname_match.group(1)

    # Add quality tracking
    result.add_text_field('employer_name', data['employer_name'])
    result.add_field('box1_wages', data['box1_wages'], {
        'field': 'box1_wages', 'found': data['box1_wages'] > 0,
        'confidence': 85 if data['box1_wages'] > 0 else 0, 'issues': []
    })
    result.add_field('box2_fed_withholding', data['box2_fed_withholding'], {
        'field': 'box2_fed_withholding', 'found': True,
        'confidence': 80, 'issues': []
    })
    result.add_field('box3_ss_wages', data['box3_ss_wages'])
    result.add_field('box4_ss_tax', data['box4_ss_tax'])
    result.add_field('box5_medicare_wages', data['box5_medicare_wages'])
    result.add_field('box6_medicare_tax', data['box6_medicare_tax'])

    # Validate required fields
    result.check_required(['employer_name', 'box1_wages'])

    # Validate math: SS tax should be ~6.2% of SS wages
    if data['box3_ss_wages'] > 0 and data['box4_ss_tax'] > 0:
        expected_ss_tax = data['box3_ss_wages'] * 0.062
        result.check_math('SS Tax = 6.2% of SS Wages', expected_ss_tax, data['box4_ss_tax'], tolerance=1.0)

    # Validate: Medicare tax should be ~1.45% of Medicare wages
    if data['box5_medicare_wages'] > 0 and data['box6_medicare_tax'] > 0:
        expected_med_tax = data['box5_medicare_wages'] * 0.0145
        result.check_math('Medicare Tax = 1.45% of Medicare Wages', expected_med_tax, data['box6_medicare_tax'], tolerance=1.0)

    # Return dict with quality embedded
    return result.to_dict()


def parse_1099int(text, filename, is_ocr=False):
    """Parse 1099-INT interest income with quality tracking."""
    result = ExtractionResult('1099-INT', filename, is_ocr)
    data = {
        'payer_name': '',
        'box1_interest': 0,
        'box4_fed_withholding': 0,
        'box8_tax_exempt_interest': 0,
        'source_file': filename
    }

    payer_patterns = [
        r"[Pp]ayer'?s?\s+name.*?\n\s*([A-Z][A-Za-z0-9\s\.,&-]+)",
        r"PAYER.*?\n([A-Z][A-Za-z0-9\s\.,&-]+)",
        # Vanguard format
        r"(VANGUARD\s+(?:MARKETING|BROKERAGE)[A-Z\s]*)",
    ]
    for pattern in payer_patterns:
        match = re.search(pattern, text)
        if match:
            data['payer_name'] = match.group(1).strip()[:50]
            break

    data['box1_interest'], q1 = extract_amount_with_quality(text, [
        # Vanguard consolidated format: "1- Interest income 123.45"
        r'1-?\s*[-:]?\s*[Ii]nterest\s+[Ii]ncome[^0-9]*([\d,]+\.?\d*)',
        r'[Bb]ox\s*1[:\s]+\$?([\d,]+\.?\d*)',
        r'1\s+[Ii]nterest\s+[Ii]ncome\s+\$?([\d,]+\.?\d*)',
    ], 'box1_interest')
    result.add_field('box1_interest', data['box1_interest'], q1)

    data['box4_fed_withholding'], q4 = extract_amount_with_quality(text, [
        # Vanguard: "4- Federal income tax withheld 0.00"
        r'4-?\s*[-:]?\s*[Ff]ederal\s+[Ii]ncome\s+[Tt]ax\s+[Ww]ithheld[^0-9]*([\d,]+\.?\d*)',
        r'[Bb]ox\s*4[:\s]+\$?([\d,]+\.?\d*)',
    ], 'box4_fed_withholding')
    result.add_field('box4_fed_withholding', data['box4_fed_withholding'], q4)

    result.add_text_field('payer_name', data['payer_name'])
    result.check_required(['payer_name', 'box1_interest'])

    return result.to_dict()


def parse_1099div(text, filename, is_ocr=False):
    """Parse 1099-DIV dividends."""
    data = {
        'payer_name': '',
        'box1a_ordinary_dividends': 0,
        'box1b_qualified_dividends': 0,
        'box2a_total_cap_gain': 0,
        'box3_nondiv_dist': 0,
        'box5_sec199a': 0,
        'box7_foreign_tax': 0,
        'box4_fed_withholding': 0,
        'source_file': filename
    }

    payer_patterns = [
        r"[Pp]ayer'?s?\s+name.*?\n\s*([A-Z][A-Za-z0-9\s\.,&-]+)",
        r"PAYER.*?\n([A-Z][A-Za-z0-9\s\.,&-]+)",
        # Vanguard format - look for VANGUARD in header
        r"(VANGUARD\s+(?:MARKETING|BROKERAGE)[A-Z\s]*)",
    ]
    for pattern in payer_patterns:
        match = re.search(pattern, text)
        if match:
            data['payer_name'] = match.group(1).strip()[:50]
            break

    # Box 1a - Ordinary Dividends
    data['box1a_ordinary_dividends'] = extract_amount(text, [
        # Vanguard consolidated format: "1a- Total ordinary dividends (includes...) 2,062.45"
        r'1a-?\s*[-:]?\s*[Tt]otal\s+[Oo]rdinary\s+[Dd]ividends[^0-9]*([\d,]+\.?\d*)',
        r'1a\s+[Oo]rdinary\s+[Dd]ividends.*?\$?([\d,]+\.?\d*)',
        r'[Oo]rdinary\s+[Dd]ividends.*?\$?([\d,]+\.?\d*)',
    ])

    # Box 1b - Qualified Dividends
    data['box1b_qualified_dividends'] = extract_amount(text, [
        # Vanguard: "1b- Qualified dividends 1,062.49"
        r'1b-?\s*[-:]?\s*[Qq]ualified\s+[Dd]ividends[^0-9]*([\d,]+\.?\d*)',
        r'1b\s+[Qq]ualified\s+[Dd]ividends.*?\$?([\d,]+\.?\d*)',
        r'[Qq]ualified\s+[Dd]ividends.*?\$?([\d,]+\.?\d*)',
    ])

    # Box 2a - Total Capital Gain Distributions
    data['box2a_total_cap_gain'] = extract_amount(text, [
        # Vanguard: "2a- Total capital gain distributions (includes...) 3,715.37"
        r'2a-?\s*[-:]?\s*[Tt]otal\s+[Cc]apital\s+[Gg]ain[^0-9]*([\d,]+\.?\d*)',
        r'2a\s+[Tt]otal\s+[Cc]apital\s+[Gg]ain.*?\$?([\d,]+\.?\d*)',
    ])

    # Box 3 - Nondividend Distributions
    data['box3_nondiv_dist'] = extract_amount(text, [
        r'3-?\s*[-:]?\s*[Nn]ondividend\s+[Dd]istributions[^0-9]*([\d,]+\.?\d*)',
    ])

    # Box 5 - Section 199A Dividends
    data['box5_sec199a'] = extract_amount(text, [
        r'5-?\s*[-:]?\s*[Ss]ection\s*199A\s+[Dd]ividends[^0-9]*([\d,]+\.?\d*)',
    ])

    # Box 7 - Foreign Tax Paid
    data['box7_foreign_tax'] = extract_amount(text, [
        r'7-?\s*[-:]?\s*[Ff]oreign\s+[Tt]ax\s+[Pp]aid[^0-9]*([\d,]+\.?\d*)',
    ])

    # Box 4 - Federal Withholding
    data['box4_fed_withholding'] = extract_amount(text, [
        r'4-?\s*[-:]?\s*[Ff]ederal\s+[Ii]ncome\s+[Tt]ax\s+[Ww]ithheld[^0-9]*([\d,]+\.?\d*)',
    ])

    return data


def parse_1099r(text, filename):
    """Parse 1099-R retirement distributions."""
    data = {
        'payer_name': '',
        'box1_gross_distribution': 0,
        'box2a_taxable_amount': 0,
        'box4_fed_withholding': 0,
        'box7_distribution_code': '',
        'source_file': filename
    }

    payer_patterns = [
        r"[Pp]ayer'?s?\s+name.*?\n\s*([A-Z][A-Za-z0-9\s\.,&-]+)",
        r"PAYER.*?\n([A-Z][A-Za-z0-9\s\.,&-]+)",
    ]
    for pattern in payer_patterns:
        match = re.search(pattern, text)
        if match:
            data['payer_name'] = match.group(1).strip()[:50]
            break

    data['box1_gross_distribution'] = extract_amount(text, [
        r'1\s+[Gg]ross\s+[Dd]istribution.*?\$?([\d,]+\.?\d*)',
        r'[Gg]ross\s+[Dd]istribution.*?\$?([\d,]+\.?\d*)',
    ])

    data['box2a_taxable_amount'] = extract_amount(text, [
        r'2a\s+[Tt]axable\s+[Aa]mount.*?\$?([\d,]+\.?\d*)',
    ])

    data['box4_fed_withholding'] = extract_amount(text, [
        r'4\s+[Ff]ederal.*?withheld.*?\$?([\d,]+\.?\d*)',
    ])

    code_match = re.search(r'7\s+[Dd]istribution\s+[Cc]ode.*?([0-9A-Z]{1,2})', text)
    if code_match:
        data['box7_distribution_code'] = code_match.group(1)

    return data


def parse_ssa1099(text, filename):
    """Parse SSA-1099 Social Security benefits."""
    data = {
        'description': 'Social Security Benefits',
        'box3_benefits_paid': 0,
        'box5_net_benefits': 0,
        'box6_fed_withholding': 0,
        'source_file': filename
    }

    # Box 3: Benefits Paid - look for dollar amount after "Box 3" or "Benefits Paid"
    data['box3_benefits_paid'] = extract_amount(text, [
        r'[Bb]ox\s*3.*?[Bb]enefits\s+[Pp]aid.*?\$([\d,]+\.?\d*)',
        r'[Bb]enefits\s+[Pp]aid\s+in\s+\d{4}\s*\$([\d,]+\.?\d*)',
        r'[Bb]ox\s*3[.\s]+[Bb]enefits.*?\$([\d,]+\.?\d*)',
        r'\$([\d,]+\.?\d*)\s*\n.*?DESCRIPTION\s+OF\s+AMOUNT',
    ])

    # Box 5: Net Benefits - often same as Box 3 if no repayments
    data['box5_net_benefits'] = extract_amount(text, [
        r'[Bb]ox\s*5.*?[Nn]et\s+[Bb]enefits.*?\$([\d,]+\.?\d*)',
        r'[Nn]et\s+[Bb]enefits\s+for\s+\d{4}.*?\$([\d,]+\.?\d*)',
        r'[Bb]enefits\s+for\s+\d{4}\s*\$([\d,]+\.?\d*)',
    ])

    # If Box 5 not found but Box 3 was, use Box 3 (common when no repayments)
    if data['box5_net_benefits'] == 0 and data['box3_benefits_paid'] > 0:
        data['box5_net_benefits'] = data['box3_benefits_paid']

    # Box 6: Voluntary Federal Withholding
    data['box6_fed_withholding'] = extract_amount(text, [
        r'[Bb]ox\s*6.*?[Ww]ithheld.*?\$([\d,]+\.?\d*)',
        r'[Ff]ederal.*?[Ww]ithheld.*?\$([\d,]+\.?\d*)',
    ])

    return data


def parse_1098(text, filename):
    """Parse 1098 Mortgage Interest Statement."""
    data = {
        'lender_name': '',
        'box1_mortgage_interest': 0,
        'box2_outstanding_principal': 0,
        'box5_mortgage_insurance': 0,
        'box10_property_tax': 0,
        'property_address': '',
        'source_file': filename
    }

    # Lender name - look for bank names at start of document (first few lines)
    # The format shows "FIFTH THIRD BANK, N.A." at the top
    lender_patterns = [
        r"(FIFTH\s+THIRD\s+BANK[,\s]*N\.?A\.?)",
        r"([A-Z]+\s+THIRD\s+BANK[,\s]*N\.?A\.?)",
        r"^([A-Z][A-Z\s]+BANK[,\s]+N\.A\.)",
        r"([A-Z]+\s+[A-Z]+\s+BANK,?\s*N\.?A\.?)",
        r"([A-Z][A-Z\s]+(?:BANK|MORTGAGE|CREDIT UNION)[A-Za-z\s\.,]*)",
    ]
    for pattern in lender_patterns:
        match = re.search(pattern, text, re.MULTILINE)
        if match:
            name = match.group(1).strip().split('\n')[0]
            # Filter out bad matches
            if len(name) > 5 and 'ZIP' not in name and 'PAYER' not in name and 'RECIPIENT' not in name:
                data['lender_name'] = name[:60]
                break

    # Box 1: Mortgage interest - format shows "1Mortgageinterestreceivedfrompayer(s)/borrower(s)*"
    # followed by newline then "$6,871.22"
    box1_patterns = [
        r'1[Mm]ortgageinterest[a-z\(\)/\*]+\n\$([\d,]+\.\d{2})',  # Run-on text then newline
        r'\$([\d,]+\.\d{2})\s*\nRECIPIENT',  # Amount before RECIPIENT'S TIN
        r'1\s*[Mm]ortgage\s*[Ii]nterest.*?\n\$\s*([\d,]+\.\d{2})',
        r'[Mm]ortgage\s*[Ii]nterest\s*[Rr]eceived.*?\$\s*([\d,]+\.\d{2})',
        r'[Mm]ortgage\s*[Ii]nterest.*?\$([\d,]+\.\d{2})',
    ]
    data['box1_mortgage_interest'] = extract_amount(text, box1_patterns)

    # Box 2: Outstanding mortgage principal
    box2_patterns = [
        r'2\s*[Oo]utstanding\s*[Mm]ortgage.*?\$\s*([\d,]+\.\d{2})',
        r'[Oo]utstanding\s*mortgage\s*\n?\s*principal\s*\$\s*([\d,]+\.\d{2})',
        r'\$\s*([\d,]+\.\d{2})\s*\n.*?[Mm]ortgage\s*origination',
    ]
    data['box2_outstanding_principal'] = extract_amount(text, box2_patterns)

    # Box 5: Mortgage insurance premiums
    data['box5_mortgage_insurance'] = extract_amount(text, [
        r'5\s*[Mm]ortgage\s*[Ii]nsurance.*?\$\s*([\d,]+\.\d{2})',
    ])

    # Box 10: Real estate/property tax
    data['box10_property_tax'] = extract_amount(text, [
        r'10\s*[Oo]ther.*?\$\s*([\d,]+\.\d{2})',
        r'[Rr]eal\s*[Ee]state\s*[Tt]ax.*?\$\s*([\d,]+\.\d{2})',
    ])

    # Property address from Box 8
    addr_match = re.search(r'(\d+\s+[A-Z]+\s+[A-Z]+\s+(?:BLVD|DR|ST|AVE|RD|LN|CT|WAY|HWY)[A-Za-z0-9\s,]*(?:FL|CA|NY|NJ|TX)\s*\d{5})', text)
    if addr_match:
        data['property_address'] = addr_match.group(1).strip()[:80]

    return data


def parse_property_tax(text, filename):
    """Parse property tax bill."""
    data = {
        'county': '',
        'property_address': '',
        'ad_valorem_taxes': 0,
        'total_taxes': 0,
        'taxable_value': 0,
        'parcel_number': '',
        'source_file': filename
    }

    # County name
    county_match = re.search(r'([A-Z]+)\s+(?:COUNTY|CO)\b', text)
    if county_match:
        data['county'] = county_match.group(1) + ' COUNTY'

    # Parcel number - format like "R092527-305700010720"
    parcel_patterns = [
        r'PARCEL\s+ACCOUNT\s+NUMBER[^\n]*\n([A-Z]?\d+[-\d]+)',
        r'([R]\d{6}-\d+)',
    ]
    for pattern in parcel_patterns:
        match = re.search(pattern, text)
        if match:
            data['parcel_number'] = match.group(1)
            break

    # Ad valorem taxes - look for dollar amounts near "AD VALOREM" or "Paid By"
    # The Osceola format shows "$5,087.58" followed by "Paid By"
    ad_valorem_patterns = [
        r'\$([\d,]+\.\d{2})\s*\n?Paid\s*By',  # Amount before "Paid By"
        r'\$([\d,]+\.\d{2})\s*\nPaid',  # Amount before "Paid"
        r'TOTAL\s+MILLAGE\s+AD\s+VALOREM\s*TAXES\s*\n[^\$]*\$([\d,]+\.\d{2})',
        r'AD\s*VALOREM\s*TAXES\s*\$?\s*([\d,]+\.?\d*)',
        r'ADVALOREMTAXES[^\$]*\$([\d,]+\.\d{2})',
    ]
    data['ad_valorem_taxes'] = extract_amount(text, ad_valorem_patterns)

    # Combined/total taxes - look for amount near COMBINED TAXES
    total_patterns = [
        r'COMBINEDTAXES[^\$]*\$([\d,]+\.\d{2})',
        r'COMBINED\s*TAXES[^\$]*\$\s*([\d,]+\.?\d*)',
        r'TOTAL.*?TAXES.*?\$\s*([\d,]+\.?\d*)',
    ]
    data['total_taxes'] = extract_amount(text, total_patterns)

    # If no ad valorem found but total found, use total
    if data['ad_valorem_taxes'] == 0 and data['total_taxes'] > 0:
        data['ad_valorem_taxes'] = data['total_taxes']

    # Taxable value
    value_match = re.search(r'TAXABLE\s+VALUE[^\d]*([\d,]+)', text)
    if value_match:
        data['taxable_value'] = float(value_match.group(1).replace(',', ''))

    # Property address - look for street address pattern
    addr_match = re.search(r'(\d+\s+[A-Z]+\s+[A-Z]+\s+(?:BLVD|DR|ST|AVE|RD|LN|CT|WAY|HWY|MEMORIAL))', text)
    if addr_match:
        data['property_address'] = addr_match.group(1)

    return data


def parse_1098t(text, filename, is_ocr=False):
    """Parse 1098-T Tuition Statement."""
    result = ExtractionResult('1098-T', filename, is_ocr)
    data = {
        'school_name': '',
        'box1_payments_received': 0,
        'box2_amounts_billed': 0,
        'box4_adjustments_prior_year': 0,
        'box5_scholarships': 0,
        'box6_adjustments_scholarships': 0,
        'box7_checked': False,  # Box 7 checked = amounts for academic period beginning in Jan-Mar of next year
        'box8_half_time': False,
        'box9_graduate': False,
        'source_file': filename
    }

    # School/Institution name - look at top of form
    school_patterns = [
        r"FILER'?S?\s+(?:name|NAME)[:\s]*\n?\s*([A-Z][A-Za-z0-9\s\.,&-]+(?:UNIVERSITY|COLLEGE|INSTITUTE|SCHOOL))",
        r"([A-Z][A-Za-z\s]+(?:UNIVERSITY|COLLEGE|INSTITUTE|SCHOOL)[A-Za-z\s]*)",
        r"^([A-Z][A-Z\s]+(?:UNIVERSITY|COLLEGE))",
    ]
    for pattern in school_patterns:
        match = re.search(pattern, text, re.MULTILINE)
        if match:
            name = match.group(1).strip().split('\n')[0]
            if len(name) > 5:
                data['school_name'] = name[:60]
                break

    # Box 1: Payments received for qualified tuition
    data['box1_payments_received'], q1 = extract_amount_with_quality(text, [
        r'1\s*[Pp]ayments\s+[Rr]eceived.*?\$?\s*([\d,]+\.?\d*)',
        r'[Bb]ox\s*1[:\s]+\$?([\d,]+\.?\d*)',
        r'[Pp]ayments\s+[Rr]eceived\s+for\s+[Qq]ualified.*?\$?\s*([\d,]+\.?\d*)',
    ], 'box1_payments_received')
    result.add_field('box1_payments_received', data['box1_payments_received'], q1)

    # Box 2: Amounts billed (older forms used this instead of Box 1)
    data['box2_amounts_billed'], q2 = extract_amount_with_quality(text, [
        r'2\s*[Aa]mounts\s+[Bb]illed.*?\$?\s*([\d,]+\.?\d*)',
        r'[Bb]ox\s*2[:\s]+\$?([\d,]+\.?\d*)',
    ], 'box2_amounts_billed')
    result.add_field('box2_amounts_billed', data['box2_amounts_billed'], q2)

    # Box 4: Adjustments made for a prior year
    data['box4_adjustments_prior_year'] = extract_amount(text, [
        r'4\s*[Aa]djustments\s+[Mm]ade.*?[Pp]rior.*?\$?\s*([\d,]+\.?\d*)',
        r'[Bb]ox\s*4[:\s]+\$?([\d,]+\.?\d*)',
    ])

    # Box 5: Scholarships or grants
    data['box5_scholarships'], q5 = extract_amount_with_quality(text, [
        r'5\s*[Ss]cholarships\s+[Oo]r\s+[Gg]rants.*?\$?\s*([\d,]+\.?\d*)',
        r'[Bb]ox\s*5[:\s]+\$?([\d,]+\.?\d*)',
        r'[Ss]cholarships.*?[Gg]rants.*?\$?\s*([\d,]+\.?\d*)',
    ], 'box5_scholarships')
    result.add_field('box5_scholarships', data['box5_scholarships'], q5)

    # Box 6: Adjustments to scholarships
    data['box6_adjustments_scholarships'] = extract_amount(text, [
        r'6\s*[Aa]djustments.*?[Ss]cholarships.*?\$?\s*([\d,]+\.?\d*)',
        r'[Bb]ox\s*6[:\s]+\$?([\d,]+\.?\d*)',
    ])

    # Box 7, 8, 9 are checkboxes
    data['box7_checked'] = bool(re.search(r'[Bb]ox\s*7.*?[Xx✓]|7\s*[Xx✓]', text))
    data['box8_half_time'] = bool(re.search(r'[Bb]ox\s*8.*?[Xx✓]|[Hh]alf.?[Tt]ime.*?[Xx✓]', text))
    data['box9_graduate'] = bool(re.search(r'[Bb]ox\s*9.*?[Xx✓]|[Gg]raduate.*?[Xx✓]', text))

    result.add_text_field('school_name', data['school_name'])
    result.check_required(['school_name'])

    return result.to_dict()


def parse_1099q(text, filename, is_ocr=False):
    """Parse 1099-Q Payments From Qualified Education Programs."""
    result = ExtractionResult('1099-Q', filename, is_ocr)
    data = {
        'payer_name': '',
        'box1_gross_distribution': 0,
        'box2_earnings': 0,
        'box3_basis': 0,
        'box4_trustee_transfer': False,
        'box5_distribution_type': '',  # 1=529, 2=Coverdell
        'box6_designated_beneficiary': False,
        'source_file': filename
    }

    # Payer/Trustee name
    payer_patterns = [
        r"[Pp]ayer'?s?/[Tt]rustee'?s?\s+name.*?\n\s*([A-Z][A-Za-z0-9\s\.,&-]+)",
        r"PAYER.*?\n([A-Z][A-Za-z0-9\s\.,&-]+)",
        r"TRUSTEE.*?\n([A-Z][A-Za-z0-9\s\.,&-]+)",
        r"(FIDELITY|VANGUARD|SCHWAB|AMERICAN FUNDS|T\. ROWE PRICE|TIAA)[A-Z\s]*",
    ]
    for pattern in payer_patterns:
        match = re.search(pattern, text)
        if match:
            data['payer_name'] = match.group(1).strip()[:50]
            break

    # Box 1: Gross distribution
    data['box1_gross_distribution'], q1 = extract_amount_with_quality(text, [
        r'1\s*[Gg]ross\s+[Dd]istribution.*?\$?\s*([\d,]+\.?\d*)',
        r'[Bb]ox\s*1[:\s]+\$?([\d,]+\.?\d*)',
        r'[Gg]ross\s+[Dd]istribution[^$\d]*([\d,]+\.?\d*)',
    ], 'box1_gross_distribution')
    result.add_field('box1_gross_distribution', data['box1_gross_distribution'], q1)

    # Box 2: Earnings
    data['box2_earnings'], q2 = extract_amount_with_quality(text, [
        r'2\s*[Ee]arnings.*?\$?\s*([\d,]+\.?\d*)',
        r'[Bb]ox\s*2[:\s]+\$?([\d,]+\.?\d*)',
    ], 'box2_earnings')
    result.add_field('box2_earnings', data['box2_earnings'], q2)

    # Box 3: Basis
    data['box3_basis'], q3 = extract_amount_with_quality(text, [
        r'3\s*[Bb]asis.*?\$?\s*([\d,]+\.?\d*)',
        r'[Bb]ox\s*3[:\s]+\$?([\d,]+\.?\d*)',
    ], 'box3_basis')
    result.add_field('box3_basis', data['box3_basis'], q3)

    # Box 4: Trustee-to-trustee transfer (checkbox)
    data['box4_trustee_transfer'] = bool(re.search(r'[Bb]ox\s*4.*?[Xx✓]|[Tt]rustee.*?[Tt]ransfer.*?[Xx✓]', text))

    # Box 5: Type of account (1=529, 2=Coverdell ESA)
    type_match = re.search(r'[Bb]ox\s*5.*?([12])|[Pp]rivate|[Ss]tate', text)
    if type_match:
        if type_match.group(1):
            data['box5_distribution_type'] = type_match.group(1)
        elif 'PRIVATE' in text.upper() or 'COVERDELL' in text.upper():
            data['box5_distribution_type'] = '2'
        else:
            data['box5_distribution_type'] = '1'

    # Box 6: Designated beneficiary (checkbox)
    data['box6_designated_beneficiary'] = bool(re.search(r'[Bb]ox\s*6.*?[Xx✓]', text))

    result.add_text_field('payer_name', data['payer_name'])
    result.check_required(['payer_name', 'box1_gross_distribution'])

    return result.to_dict()


def parse_k1(text, filename, is_ocr=False):
    """Parse Schedule K-1 (Form 1065/1120S/1041)."""
    result = ExtractionResult('K-1', filename, is_ocr)
    data = {
        'entity_name': '',
        'entity_ein': '',
        'partner_name': '',
        'k1_type': '',  # 1065 (partnership), 1120S (S-corp), 1041 (trust/estate)
        'box1_ordinary_income': 0,
        'box2_net_rental_income': 0,
        'box3_other_net_rental_income': 0,
        'box4_guaranteed_payments': 0,
        'box5_interest_income': 0,
        'box6a_ordinary_dividends': 0,
        'box6b_qualified_dividends': 0,
        'box7_royalties': 0,
        'box8_net_st_cap_gain': 0,
        'box9a_net_lt_cap_gain': 0,
        'box10_net_1231_gain': 0,
        'box11_other_income': 0,
        'box12_sec179_deduction': 0,
        'box13_other_deductions': 0,
        'box14_self_employment': 0,
        'box15_credits': 0,
        'box16_foreign_transactions': 0,
        'box19_distributions': 0,
        'box20_other_info': 0,
        'source_file': filename
    }

    # Determine K-1 type
    if 'FORM 1065' in text.upper() or "PARTNER'S SHARE" in text.upper():
        data['k1_type'] = '1065'
    elif 'FORM 1120S' in text.upper() or 'FORM 1120-S' in text.upper() or "SHAREHOLDER'S SHARE" in text.upper():
        data['k1_type'] = '1120S'
    elif 'FORM 1041' in text.upper() or "BENEFICIARY'S SHARE" in text.upper():
        data['k1_type'] = '1041'

    # Entity name (partnership/S-corp/trust name)
    entity_patterns = [
        r"[Pp]artnership'?s?\s+name.*?\n\s*([A-Z][A-Za-z0-9\s\.,&-]+)",
        r"[Cc]orporation'?s?\s+name.*?\n\s*([A-Z][A-Za-z0-9\s\.,&-]+)",
        r"[Ee]state'?s?\s+or\s+[Tt]rust'?s?\s+name.*?\n\s*([A-Z][A-Za-z0-9\s\.,&-]+)",
        r"Part\s+I[^\n]*\n([A-Z][A-Za-z0-9\s\.,&-]+(?:LLC|LP|LLP|INC|CORP)?)",
    ]
    for pattern in entity_patterns:
        match = re.search(pattern, text)
        if match:
            data['entity_name'] = match.group(1).strip()[:60]
            break

    # Entity EIN
    ein_match = re.search(r'[Ee]mployer.*?[Ii]dentification.*?(\d{2}-\d{7})', text)
    if ein_match:
        data['entity_ein'] = ein_match.group(1)

    # Partner/Shareholder name
    partner_patterns = [
        r"[Pp]artner'?s?\s+name.*?\n\s*([A-Z][A-Za-z\s,]+)",
        r"[Ss]hareholder'?s?\s+name.*?\n\s*([A-Z][A-Za-z\s,]+)",
        r"[Bb]eneficiary'?s?\s+name.*?\n\s*([A-Z][A-Za-z\s,]+)",
    ]
    for pattern in partner_patterns:
        match = re.search(pattern, text)
        if match:
            data['partner_name'] = match.group(1).strip()[:60]
            break

    # Box 1: Ordinary business income (loss)
    data['box1_ordinary_income'], q1 = extract_amount_with_quality(text, [
        r'1\s+[Oo]rdinary\s+[Bb]usiness\s+[Ii]ncome.*?\(?\$?\s*(-?[\d,]+\.?\d*)\)?',
        r'[Bb]ox\s*1[:\s]+\(?\$?\s*(-?[\d,]+\.?\d*)\)?',
        r'[Oo]rdinary\s+[Ii]ncome.*?\(?\$?\s*(-?[\d,]+\.?\d*)\)?',
    ], 'box1_ordinary_income')
    result.add_field('box1_ordinary_income', data['box1_ordinary_income'], q1)

    # Box 2: Net rental real estate income
    data['box2_net_rental_income'] = extract_amount(text, [
        r'2\s+[Nn]et\s+[Rr]ental\s+[Rr]eal\s+[Ee]state.*?\(?\$?\s*(-?[\d,]+\.?\d*)\)?',
    ])

    # Box 4: Guaranteed payments
    data['box4_guaranteed_payments'] = extract_amount(text, [
        r'4\s+[Gg]uaranteed\s+[Pp]ayments.*?\$?\s*([\d,]+\.?\d*)',
    ])

    # Box 5: Interest income
    data['box5_interest_income'] = extract_amount(text, [
        r'5\s+[Ii]nterest\s+[Ii]ncome.*?\$?\s*([\d,]+\.?\d*)',
    ])

    # Box 6a: Ordinary dividends
    data['box6a_ordinary_dividends'] = extract_amount(text, [
        r'6a\s+[Oo]rdinary\s+[Dd]ividends.*?\$?\s*([\d,]+\.?\d*)',
    ])

    # Box 6b: Qualified dividends
    data['box6b_qualified_dividends'] = extract_amount(text, [
        r'6b\s+[Qq]ualified\s+[Dd]ividends.*?\$?\s*([\d,]+\.?\d*)',
    ])

    # Box 8: Net short-term capital gain
    data['box8_net_st_cap_gain'] = extract_amount(text, [
        r'8\s+[Nn]et\s+[Ss]hort.*?[Cc]apital\s+[Gg]ain.*?\(?\$?\s*(-?[\d,]+\.?\d*)\)?',
    ])

    # Box 9a: Net long-term capital gain
    data['box9a_net_lt_cap_gain'] = extract_amount(text, [
        r'9a\s+[Nn]et\s+[Ll]ong.*?[Cc]apital\s+[Gg]ain.*?\(?\$?\s*(-?[\d,]+\.?\d*)\)?',
    ])

    # Box 10: Net section 1231 gain
    data['box10_net_1231_gain'] = extract_amount(text, [
        r'10\s+[Nn]et\s+[Ss]ection\s*1231\s+[Gg]ain.*?\(?\$?\s*(-?[\d,]+\.?\d*)\)?',
    ])

    # Box 19: Distributions
    data['box19_distributions'] = extract_amount(text, [
        r'19\s+[Dd]istributions.*?\$?\s*([\d,]+\.?\d*)',
    ])

    result.add_text_field('entity_name', data['entity_name'])
    result.add_text_field('k1_type', data['k1_type'])
    result.check_required(['entity_name'])

    return result.to_dict()


def parse_consolidated_1099(text, filename):
    """Parse consolidated 1099 (contains multiple 1099 types like Vanguard/Fidelity)."""
    results = {'1099-INT': [], '1099-DIV': [], '1099-B': []}

    # Extract payer name
    payer_name = ''
    payer_patterns = [
        r"(VANGUARD\s+(?:MARKETING|BROKERAGE)[A-Z\s]*)",
        r"(FIDELITY\s+INVESTMENTS[A-Z\s]*)",
        r"([A-Z][A-Za-z]+\s+(?:BANK|INVESTMENTS|SECURITIES|BROKERAGE|FINANCIAL))",
    ]
    for pattern in payer_patterns:
        match = re.search(pattern, text)
        if match:
            payer_name = match.group(1).strip()[:50]
            break

    # For Vanguard-style documents, the whole page may contain multiple sections
    # Try to extract DIV data from entire text first (Vanguard format)
    div_data = parse_1099div(text, filename)
    if payer_name:
        div_data['payer_name'] = payer_name
    if div_data.get('box1a_ordinary_dividends', 0) > 0:
        results['1099-DIV'].append(div_data)
    else:
        # Try section-based extraction (traditional consolidated format)
        div_section = re.search(r'1099-DIV.*?(?=1099-(?!DIV)|$)', text, re.DOTALL | re.IGNORECASE)
        if div_section:
            div_data = parse_1099div(div_section.group(0), filename)
            div_data['payer_name'] = payer_name or div_data.get('payer_name', '')
            if div_data['box1a_ordinary_dividends'] > 0:
                results['1099-DIV'].append(div_data)

    # Try INT extraction from entire text first
    int_data = parse_1099int(text, filename)
    if payer_name and not int_data.get('payer_name'):
        int_data['payer_name'] = payer_name
    if int_data.get('box1_interest', 0) > 0:
        results['1099-INT'].append(int_data)
    else:
        # Try section-based extraction
        int_section = re.search(r'1099-INT.*?(?=1099-(?!INT)|$)', text, re.DOTALL | re.IGNORECASE)
        if int_section:
            int_data = parse_1099int(int_section.group(0), filename)
            int_data['payer_name'] = payer_name or int_data.get('payer_name', '')
            if int_data['box1_interest'] > 0:
                results['1099-INT'].append(int_data)

    return results


def parse_folder(folder_path):
    """Parse all PDFs in a folder and return structured data."""
    folder = Path(folder_path)

    if not folder.exists():
        print(f"ERROR: Folder not found: {folder}")
        return None

    results = {
        'metadata': {
            'source_folder': str(folder),
            'parse_date': datetime.now().isoformat(),
            'files_processed': 0,
            'files_skipped': 0
        },
        'forms': {
            'W-2': [],
            'IRS-1099INT': [],
            'IRS-1099DIV': [],
            'IRS-1099R': [],
            'SSA-1099': [],
            '1098': [],
            '1098-T': [],
            '1099-B': [],
            '1099-Q': [],
            'K-1': []
        }
    }

    # Get unique PDF files (case-insensitive), excluding generated review files
    pdf_files = list(set(folder.glob('*.pdf')) | set(folder.glob('*.PDF')))
    # Exclude document_review.pdf files (generated combined PDFs)
    pdf_files = [f for f in pdf_files if 'document_review' not in f.name.lower()]
    print(f"Found {len(pdf_files)} PDF files in {folder}")

    for pdf_path in pdf_files:
        filename = pdf_path.name
        print(f"\nProcessing: {filename}")

        text, is_scanned = extract_text_from_pdf(pdf_path)
        if not text.strip():
            if OCR_AVAILABLE:
                print(f"  Skipped: OCR failed to extract text")
            else:
                print(f"  Skipped: Scanned document (install pytesseract & pdf2image for OCR)")
            results['metadata']['files_skipped'] += 1
            continue

        if is_scanned:
            print(f"  (OCR extracted)")

        doc_type = identify_document_type(text, filename)
        print(f"  Identified as: {doc_type}")

        if doc_type == 'W-2':
            data = parse_w2(text, filename, is_ocr=is_scanned)
            # Add quality metadata if not already present
            if '_quality' not in data:
                data = add_quality_metadata(data, 'W-2', is_scanned, ['employer_name', 'box1_wages'])
            results['forms']['W-2'].append(data)
            conf = data.get('_quality', {}).get('overall_confidence', 'N/A')
            ocr_tag = " [OCR]" if is_scanned else ""
            print(f"    Employer: {data.get('employer_name', '')}")
            print(f"    Wages: ${data.get('box1_wages', 0):,.2f}{ocr_tag} (Confidence: {conf}%)")

        elif doc_type == '1099-INT':
            data = parse_1099int(text, filename, is_ocr=is_scanned)
            if '_quality' not in data:
                data = add_quality_metadata(data, '1099-INT', is_scanned, ['payer_name', 'box1_interest'])
            results['forms']['IRS-1099INT'].append(data)
            conf = data.get('_quality', {}).get('overall_confidence', 'N/A')
            ocr_tag = " [OCR]" if is_scanned else ""
            print(f"    Payer: {data.get('payer_name', '')}")
            print(f"    Interest: ${data.get('box1_interest', 0):,.2f}{ocr_tag} (Confidence: {conf}%)")

        elif doc_type == '1099-DIV':
            data = parse_1099div(text, filename, is_ocr=is_scanned)
            data = add_quality_metadata(data, '1099-DIV', is_scanned, ['payer_name', 'box1a_ordinary_dividends'])
            results['forms']['IRS-1099DIV'].append(data)
            conf = data.get('_quality', {}).get('overall_confidence', 'N/A')
            ocr_tag = " [OCR]" if is_scanned else ""
            print(f"    Payer: {data.get('payer_name', '')}")
            print(f"    Dividends: ${data.get('box1a_ordinary_dividends', 0):,.2f}{ocr_tag} (Confidence: {conf}%)")

        elif doc_type == '1099-R':
            data = parse_1099r(text, filename)
            data = add_quality_metadata(data, '1099-R', is_scanned, ['payer_name', 'box1_gross_distribution'])
            results['forms']['IRS-1099R'].append(data)
            conf = data.get('_quality', {}).get('overall_confidence', 'N/A')
            ocr_tag = " [OCR]" if is_scanned else ""
            print(f"    Payer: {data.get('payer_name', '')}")
            print(f"    Distribution: ${data.get('box1_gross_distribution', 0):,.2f}{ocr_tag} (Confidence: {conf}%)")

        elif doc_type == 'SSA-1099':
            data = parse_ssa1099(text, filename)
            data = add_quality_metadata(data, 'SSA-1099', is_scanned, ['box5_net_benefits'])
            results['forms']['SSA-1099'].append(data)
            conf = data.get('_quality', {}).get('overall_confidence', 'N/A')
            ocr_tag = " [OCR]" if is_scanned else ""
            print(f"    Net Benefits: ${data.get('box5_net_benefits', 0):,.2f}{ocr_tag} (Confidence: {conf}%)")

        elif doc_type == '1098':
            data = parse_1098(text, filename)
            data = add_quality_metadata(data, '1098', is_scanned, ['lender_name', 'box1_mortgage_interest'])
            results['forms']['1098'].append(data)
            conf = data.get('_quality', {}).get('overall_confidence', 'N/A')
            ocr_tag = " [OCR]" if is_scanned else ""
            print(f"    Lender: {data.get('lender_name', '')}")
            print(f"    Interest: ${data.get('box1_mortgage_interest', 0):,.2f}{ocr_tag} (Confidence: {conf}%)")

        elif doc_type == '1098-T':
            data = parse_1098t(text, filename, is_ocr=is_scanned)
            results['forms']['1098-T'].append(data)
            conf = data.get('_quality', {}).get('overall_confidence', 'N/A')
            ocr_tag = " [OCR]" if is_scanned else ""
            print(f"    School: {data.get('school_name', '')}")
            print(f"    Payments: ${data.get('box1_payments_received', 0):,.2f}{ocr_tag}")
            print(f"    Scholarships: ${data.get('box5_scholarships', 0):,.2f} (Confidence: {conf}%)")

        elif doc_type == '1099-Q':
            data = parse_1099q(text, filename, is_ocr=is_scanned)
            results['forms']['1099-Q'].append(data)
            conf = data.get('_quality', {}).get('overall_confidence', 'N/A')
            ocr_tag = " [OCR]" if is_scanned else ""
            print(f"    Payer: {data.get('payer_name', '')}")
            print(f"    Gross Distribution: ${data.get('box1_gross_distribution', 0):,.2f}{ocr_tag}")
            print(f"    Earnings: ${data.get('box2_earnings', 0):,.2f} (Confidence: {conf}%)")

        elif doc_type == 'K-1':
            data = parse_k1(text, filename, is_ocr=is_scanned)
            results['forms']['K-1'].append(data)
            conf = data.get('_quality', {}).get('overall_confidence', 'N/A')
            ocr_tag = " [OCR]" if is_scanned else ""
            k1_type = data.get('k1_type', '')
            print(f"    Type: Form {k1_type}")
            print(f"    Entity: {data.get('entity_name', '')}")
            print(f"    Ordinary Income: ${data.get('box1_ordinary_income', 0):,.2f}{ocr_tag} (Confidence: {conf}%)")

        elif doc_type == 'CONSOLIDATED-1099':
            consolidated = parse_consolidated_1099(text, filename)
            for form_type, entries in consolidated.items():
                for entry in entries:
                    entry = add_quality_metadata(entry, form_type, is_scanned)
                if form_type == '1099-INT':
                    results['forms']['IRS-1099INT'].extend(entries)
                elif form_type == '1099-DIV':
                    results['forms']['IRS-1099DIV'].extend(entries)
            print(f"    Consolidated forms extracted")

        elif doc_type == 'PROPERTY-TAX':
            data = parse_property_tax(text, filename)
            data = add_quality_metadata(data, 'PROPERTY-TAX', is_scanned, ['ad_valorem_taxes'])
            if 'PROPERTY-TAX' not in results['forms']:
                results['forms']['PROPERTY-TAX'] = []
            results['forms']['PROPERTY-TAX'].append(data)
            conf = data.get('_quality', {}).get('overall_confidence', 'N/A')
            ocr_tag = " [OCR]" if is_scanned else ""
            print(f"    County: {data.get('county', '')}")
            print(f"    Property Tax: ${data.get('ad_valorem_taxes', 0):,.2f}{ocr_tag} (Confidence: {conf}%)")

        else:
            print(f"  Skipped: Unknown document type")
            results['metadata']['files_skipped'] += 1
            continue

        results['metadata']['files_processed'] += 1

    return results


def main():
    parser = argparse.ArgumentParser(
        description='Parse tax source documents and create JSON for checksheet population',
        epilog="""
Examples:
  python parse_source_docs.py "L:\\Client Folder\\2025"
  python parse_source_docs.py "L:\\Client Folder\\2025" --checksheet "C:\\Tax\\checksheet.xlsx"
        """
    )

    parser.add_argument('folder', help='Folder containing tax source documents')
    parser.add_argument('--output', '-o', help='Output JSON file')
    parser.add_argument('--checksheet', '-c', help='Run populate_checksheet.py with this template')
    parser.add_argument('--column', choices=['cch', 'source'], default='source',
                        help='Which column to fill (default: source)')

    args = parser.parse_args()
    folder_path = Path(args.folder)

    print("=" * 60)
    print("SOURCE DOCUMENT PARSER")
    print("=" * 60)
    print(f"Folder: {folder_path}")
    print()

    results = parse_folder(folder_path)
    if not results:
        sys.exit(1)

    # Output path
    if args.output:
        output_path = Path(args.output)
    else:
        folder_name = folder_path.name
        output_path = folder_path / f"{folder_name}_parsed.json"

    with open(output_path, 'w') as f:
        json.dump(results, f, indent=2)

    print()
    print("=" * 60)
    print("PARSE SUMMARY")
    print("=" * 60)
    print(f"Files processed: {results['metadata']['files_processed']}")
    print(f"Files skipped: {results['metadata']['files_skipped']}")
    print()
    for form_type, entries in results['forms'].items():
        if entries:
            print(f"  {form_type}: {len(entries)}")
    print()
    print(f"JSON saved to: {output_path}")

    # Run populate_checksheet if requested
    if args.checksheet:
        print()
        print("=" * 60)
        print("POPULATING CHECKSHEET")
        print("=" * 60)

        script_dir = Path(__file__).parent
        populate_script = script_dir / 'populate_checksheet.py'

        if not populate_script.exists():
            print(f"ERROR: populate_checksheet.py not found at {populate_script}")
            sys.exit(1)

        import subprocess
        cmd = [
            sys.executable,
            str(populate_script),
            str(output_path),
            '--checksheet', args.checksheet,
            '--column', args.column
        ]
        print(f"Running: {' '.join(cmd)}")
        subprocess.run(cmd)


if __name__ == '__main__':
    main()
