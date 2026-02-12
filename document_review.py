#!/usr/bin/env python3
# Python / Anaconda is located at: C:\Users\ngroom\AppData\Local\anaconda3\python.exe
"""
Document Review & PDF Combiner
==============================
Combines client PDFs in tax-form order and checks against expected documents.

Order:
1. W-2s
2. 1099-INT
3. 1099-DIV
4. 1099-R
5. SSA-1099
6. 1099-B (Brokerage)
7. 1098
8. Other 1099s
9. Everything else

Usage:
    python document_review.py "L:\\Client Folder\\2025"
    python document_review.py "L:\\Client Folder\\2025" --google-sheet SHEET_ID
"""

import argparse
import os
import re
import sys
from pathlib import Path
from datetime import datetime

NORMALIZE_AVAILABLE = False
try:
    from pypdf import PdfReader, PdfWriter, Transformation
    from pypdf.generic import RectangleObject
    NORMALIZE_AVAILABLE = True
except ImportError:
    try:
        from PyPDF2 import PdfReader, PdfWriter
    except ImportError:
        print("ERROR: pypdf or PyPDF2 not found. Install with: pip install pypdf")
        sys.exit(1)

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

try:
    from PIL import Image, ImageFilter
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import pillow_heif
    pillow_heif.register_heif_opener()
    HEIF_AVAILABLE = True
except ImportError:
    HEIF_AVAILABLE = False

# Supported image extensions for conversion
IMAGE_EXTENSIONS = {'.heic', '.heif', '.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp', '.webp'}

# Supported Office extensions for conversion
OFFICE_EXTENSIONS = {'.xlsx', '.xls', '.docx', '.doc', '.csv'}

# Optional Google Sheets support
GSHEETS_AVAILABLE = False
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSHEETS_AVAILABLE = True
except ImportError:
    pass

# Optional win32com for Office file conversion (Windows)
WIN32COM_AVAILABLE = False
try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    pass


# =============================================================================
# IMAGE TO PDF CONVERSION
# =============================================================================

# Standard US Letter dimensions at 300 DPI
LETTER_WIDTH_PX = 2550   # 8.5 inches * 300 DPI
LETTER_HEIGHT_PX = 3300  # 11 inches * 300 DPI

# Standard US Letter dimensions in points (72 DPI) for PDF normalization
LETTER_WIDTH_PT = 612.0   # 8.5 inches * 72 points/inch
LETTER_HEIGHT_PT = 792.0  # 11 inches * 72 points/inch


def auto_rotate_image(img):
    """
    Auto-rotate image to correct orientation.
    Handles EXIF orientation tags and detects landscape vs portrait.
    """
    # First, apply EXIF orientation (camera rotation metadata)
    try:
        from PIL import ExifTags
        exif = img.getexif()
        if exif:
            for tag, value in exif.items():
                if ExifTags.TAGS.get(tag) == 'Orientation':
                    if value == 3:
                        img = img.rotate(180, expand=True)
                    elif value == 6:
                        img = img.rotate(270, expand=True)
                    elif value == 8:
                        img = img.rotate(90, expand=True)
                    break
    except (AttributeError, KeyError, TypeError):
        pass

    # Tax documents are almost always portrait - rotate landscape images
    w, h = img.size
    if w > h * 1.2:  # Significantly wider than tall = landscape
        img = img.rotate(270, expand=True)

    return img


def trim_edges(img, threshold=240, min_border=20):
    """
    Trim whitespace and dark borders from photo edges.
    Detects the document region within the photo.
    """
    import numpy as np

    # Convert to grayscale numpy array
    gray = img.convert('L')
    arr = np.array(gray)

    # Find rows and columns that have significant content
    # (not nearly-white background and not black borders)
    row_means = arr.mean(axis=1)
    col_means = arr.mean(axis=0)

    # Content rows: not too bright (background) and not too dark (black border)
    content_rows = np.where((row_means < threshold) & (row_means > 15))[0]
    content_cols = np.where((col_means < threshold) & (col_means > 15))[0]

    if len(content_rows) < 50 or len(content_cols) < 50:
        # Not enough content detected, return original
        return img

    top = max(0, content_rows[0] - min_border)
    bottom = min(arr.shape[0], content_rows[-1] + min_border)
    left = max(0, content_cols[0] - min_border)
    right = min(arr.shape[1], content_cols[-1] + min_border)

    # Only crop if we're removing meaningful borders (at least 2% per side)
    h, w = arr.shape
    if (top > h * 0.02 or (h - bottom) > h * 0.02 or
            left > w * 0.02 or (w - right) > w * 0.02):
        img = img.crop((left, top, right, bottom))

    return img


def scale_to_letter(img):
    """
    Scale image to fit US Letter size (8.5x11) while maintaining aspect ratio.
    Adds white padding to center the document on the page.
    """
    w, h = img.size

    # Calculate scale to fit within letter size with small margins
    margin = 75  # ~0.25 inch margin at 300 DPI
    max_w = LETTER_WIDTH_PX - 2 * margin
    max_h = LETTER_HEIGHT_PX - 2 * margin

    scale = min(max_w / w, max_h / h)

    # Only scale down, don't upscale small images beyond 1.5x
    if scale > 1.5:
        scale = 1.5

    new_w = int(w * scale)
    new_h = int(h * scale)

    # Resize with high quality
    img = img.resize((new_w, new_h), Image.LANCZOS)

    # Create white letter-size canvas and paste centered
    canvas = Image.new('RGB', (LETTER_WIDTH_PX, LETTER_HEIGHT_PX), 'white')
    x_offset = (LETTER_WIDTH_PX - new_w) // 2
    y_offset = (LETTER_HEIGHT_PX - new_h) // 2
    canvas.paste(img, (x_offset, y_offset))

    return canvas


def convert_image_to_pdf(image_path, output_pdf_path):
    """
    Convert an image file to a properly oriented, trimmed, and scaled PDF.
    Returns the output path on success, None on failure.
    """
    try:
        img = Image.open(str(image_path))

        # Convert to RGB if needed (RGBA, P mode, etc.)
        if img.mode in ('RGBA', 'P', 'LA'):
            background = Image.new('RGB', img.size, 'white')
            if img.mode == 'RGBA' or img.mode == 'LA':
                background.paste(img, mask=img.split()[-1])
            else:
                background.paste(img)
            img = background
        elif img.mode != 'RGB':
            img = img.convert('RGB')

        # Step 1: Auto-rotate (EXIF + landscape detection)
        img = auto_rotate_image(img)

        # Step 2: Trim edges (remove photo borders/background)
        img = trim_edges(img)

        # Step 3: Scale to letter size with centering
        img = scale_to_letter(img)

        # Save as PDF
        img.save(str(output_pdf_path), 'PDF', resolution=300)
        return output_pdf_path

    except Exception as e:
        print(f"  WARNING: Could not convert {image_path.name}: {e}")
        return None


def convert_images_in_folder(folder):
    """
    Find all image files in folder, convert to PDF.
    Returns list of newly created PDF paths.
    """
    if not PIL_AVAILABLE:
        return []

    converted = []
    image_files = []

    for f in folder.iterdir():
        if f.suffix.lower() in IMAGE_EXTENSIONS:
            image_files.append(f)

    if not image_files:
        return []

    # Check HEIF support
    heic_files = [f for f in image_files if f.suffix.lower() in ('.heic', '.heif')]
    if heic_files and not HEIF_AVAILABLE:
        print(f"  WARNING: {len(heic_files)} HEIC files found but pillow-heif not installed")
        print(f"           Run: pip install pillow-heif")
        image_files = [f for f in image_files if f.suffix.lower() not in ('.heic', '.heif')]

    if not image_files:
        return []

    print(f"  Converting {len(image_files)} image(s) to PDF...")

    for img_path in image_files:
        pdf_path = img_path.with_suffix('.pdf')

        # Skip if PDF already exists and is newer than the image
        if pdf_path.exists() and pdf_path.stat().st_mtime > img_path.stat().st_mtime:
            converted.append(pdf_path)
            continue

        result = convert_image_to_pdf(img_path, pdf_path)
        if result:
            print(f"    {img_path.name} -> {pdf_path.name}")
            converted.append(result)

    return converted


# =============================================================================
# PDF PAGE NORMALIZATION
# =============================================================================

def normalize_pdf_page(page, auto_rotate=True):
    """
    Normalize a PDF page to US Letter (8.5x11).
    Rotates landscape pages to portrait only if auto_rotate is True.
    Scales non-letter pages to fit. Preserves vector quality.
    """
    if not NORMALIZE_AVAILABLE:
        return page

    # Transfer /Rotate into content stream so mediabox = visual dimensions
    try:
        page.transfer_rotation_to_content()
    except (AttributeError, Exception):
        pass

    mb = page.mediabox
    x0, y0 = float(mb.left), float(mb.bottom)
    w = float(mb.width)
    h = float(mb.height)

    # Shift content to origin if mediabox is offset
    if abs(x0) > 1 or abs(y0) > 1:
        page.add_transformation(Transformation().translate(-x0, -y0))
        page.mediabox = RectangleObject([0, 0, w, h])

    # Rotate landscape pages to portrait (only when auto_rotate enabled)
    if auto_rotate and w > h * 1.05:
        page.add_transformation(Transformation().rotate(270).translate(0, w))
        page.mediabox = RectangleObject([0, 0, h, w])
        w, h = h, w

    # Already letter size? (within ~0.5 inch tolerance)
    if abs(w - LETTER_WIDTH_PT) < 36 and abs(h - LETTER_HEIGHT_PT) < 36:
        return page

    # Scale to fit letter with small margin
    margin_pt = 18  # ~0.25 inch
    usable_w = LETTER_WIDTH_PT - 2 * margin_pt
    usable_h = LETTER_HEIGHT_PT - 2 * margin_pt
    scale = min(usable_w / w, usable_h / h)

    scaled_w = w * scale
    scaled_h = h * scale
    tx = (LETTER_WIDTH_PT - scaled_w) / 2
    ty = (LETTER_HEIGHT_PT - scaled_h) / 2

    page.add_transformation(
        Transformation().scale(scale, scale).translate(tx, ty)
    )
    page.mediabox = RectangleObject([0, 0, LETTER_WIDTH_PT, LETTER_HEIGHT_PT])

    # Remove conflicting box definitions
    for box in ('/CropBox', '/BleedBox', '/TrimBox', '/ArtBox'):
        if box in page:
            try:
                del page[box]
            except Exception:
                pass

    return page


# =============================================================================
# OFFICE FILE CONVERSION
# =============================================================================

def convert_office_files_in_folder(folder):
    """
    Find Word and Excel files in folder, convert to PDF via Office COM automation.
    Returns list of newly created PDF paths.
    """
    if not WIN32COM_AVAILABLE:
        return []

    office_files = [f for f in folder.iterdir() if f.suffix.lower() in OFFICE_EXTENSIONS]
    if not office_files:
        return []

    word_files = [f for f in office_files if f.suffix.lower() in ('.docx', '.doc')]
    excel_files = [f for f in office_files if f.suffix.lower() in ('.xlsx', '.xls', '.csv')
                   and 'checksheet' not in f.name.lower()]
    converted = []

    print(f"  Converting {len(office_files)} Office file(s) to PDF...")

    # Convert Word documents
    if word_files:
        try:
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False
            word.DisplayAlerts = False
            try:
                for wp in word_files:
                    pdf_path = wp.with_suffix('.pdf')
                    if pdf_path.exists() and pdf_path.stat().st_mtime > wp.stat().st_mtime:
                        converted.append(pdf_path)
                        continue
                    try:
                        doc = word.Documents.Open(str(wp.resolve()))
                        doc.SaveAs(str(pdf_path.resolve()), FileFormat=17)
                        doc.Close(False)
                        if pdf_path.exists():
                            print(f"    {wp.name} -> {pdf_path.name}")
                            converted.append(pdf_path)
                    except Exception as e:
                        print(f"    WARNING: Could not convert {wp.name}: {e}")
            finally:
                try:
                    word.Quit()
                except Exception:
                    pass
        except Exception as e:
            print(f"    WARNING: Could not start Word: {e}")

    # Convert Excel spreadsheets
    if excel_files:
        try:
            excel = win32com.client.Dispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                for ep in excel_files:
                    pdf_path = ep.with_suffix('.pdf')
                    if pdf_path.exists() and pdf_path.stat().st_mtime > ep.stat().st_mtime:
                        converted.append(pdf_path)
                        continue
                    try:
                        wb = excel.Workbooks.Open(str(ep.resolve()))
                        wb.ExportAsFixedFormat(0, str(pdf_path.resolve()))
                        wb.Close(False)
                        if pdf_path.exists():
                            print(f"    {ep.name} -> {pdf_path.name}")
                            converted.append(pdf_path)
                    except Exception as e:
                        print(f"    WARNING: Could not convert {ep.name}: {e}")
            finally:
                try:
                    excel.Quit()
                except Exception:
                    pass
        except Exception as e:
            print(f"    WARNING: Could not start Excel: {e}")

    return converted


# =============================================================================
# DOCUMENT CLASSIFICATION
# =============================================================================

def extract_pdf_text(pdf_path, max_pages=4):
    """Extract text from first few pages of PDF for classification."""
    if not PDFPLUMBER_AVAILABLE:
        return ""

    try:
        text = ""
        with pdfplumber.open(pdf_path) as pdf:
            # Read up to max_pages, but also check later pages for Vanguard-style docs
            # where first pages may be reversed text
            pages_to_read = min(max_pages, len(pdf.pages))
            for i, page in enumerate(pdf.pages[:pages_to_read]):
                page_text = page.extract_text() or ""
                text += page_text + "\n"

            # If text looks reversed (Vanguard security feature), try later pages
            if len(text) < 200 or 'snoitcasnarT' in text:  # "Transactions" reversed
                for page in pdf.pages[2:5]:  # Check pages 3-5
                    page_text = page.extract_text() or ""
                    text += page_text + "\n"
        return text
    except Exception:
        return ""


def classify_by_content(text):
    """Classify document by its content, return (form_type, payer, priority)."""
    import re
    text_upper = text.upper()

    # Pre-scan: non-W-2 form markers that override W-2 classification
    NON_W2_FORM_MARKERS = ['5498', '1099-SA', '1095-C', '1095-B',
                           '1099-MISC', '1099-NEC', '1099-G']
    has_non_w2_marker = any(marker in text_upper for marker in NON_W2_FORM_MARKERS)
    has_ss_wages = 'SOCIAL SECURITY WAGES' in text_upper

    # "Social security wages" can ONLY appear on a W-2
    if has_ss_wages:
        is_w2 = True
    elif has_non_w2_marker:
        is_w2 = False
    else:
        # Check for specific form types in content
        is_w2 = ('WAGE AND TAX STATEMENT' in text_upper or 'FORM W-2' in text_upper or
                 'WAGES, TIPS' in text_upper or 'W-2 WAGE' in text_upper)
        # Compact W-2 detection: EIN followed by two dollar amounts (no form labels)
        if not is_w2:
            is_w2 = bool(re.search(r'\d{2}-\d{7}\s*\n\s*[\d,]+\.\d{2}\s+[\d,]+\.\d{2}', text))

    if is_w2:
        # Extract employer name - try multiple patterns
        payer = ""
        emp_patterns = [
            r"[Ee]mployer'?s?\s+name[:\s]*\n?\s*([A-Z][A-Za-z0-9\s\.,&-]+)",
            r"c\s+[Ee]mployer'?s?\s+name.*?\n\s*([A-Z][A-Za-z0-9\s\.,&-]+)",
            # Compact format: after SSN, before dollar amounts
            r'[\dX]{3}-[\dX]{2}-\d{4}\s*\n\s*([A-Z][A-Z0-9&\s\.,]+?)\s+[\d,]+\.\d{2}',
            # LLC/INC/CORP in text
            r'^([A-Z][A-Za-z0-9\s\.,&]+?(?:LLC|INC|CORP|L\.L\.C\.|ENTERPRISES)[A-Za-z\s\.]*)',
        ]
        for pattern in emp_patterns:
            emp_match = re.search(pattern, text, re.MULTILINE)
            if emp_match:
                name = emp_match.group(1).strip()[:40]
                # Filter garbage names containing dollar amounts or form text
                if not re.search(r'\d+\.\d{2}', name) and len(name) > 3:
                    payer = name
                    break
        return ('W-2', payer, 1)

    # For consolidated 1099s (Vanguard/Fidelity), check which section has actual values
    has_div = '1099-DIV' in text_upper or 'DIVIDENDS AND DISTRIBUTIONS' in text_upper
    has_int = '1099-INT' in text_upper or 'INTEREST INCOME' in text_upper

    if has_div and has_int:
        # Multiple patterns for dividend values - try each until we find a match
        div_patterns = [
            # Vanguard: "1a- Total ordinary dividends (includes lines...) 14,292.84"
            r'1a[-\s]+[Tt]otal\s+[Oo]rdinary\s+[Dd]ividends[^\d]+([\d,]+\.\d{2})',
            # Fidelity: "Total ordinary dividends $ 5,344.93"
            r'[Tt]otal\s+[Oo]rdinary\s+[Dd]ividends[^\d$]*\$?\s*([\d,]+\.\d{2})',
            # Generic: number after "ordinary dividends"
            r'[Oo]rdinary\s+[Dd]ividends[^\d]*([\d,]+\.\d{2})',
            # Box 1a with amount
            r'[Bb]ox\s*1a[^\d]*([\d,]+\.\d{2})',
        ]

        int_patterns = [
            # Vanguard: "1- Interest income $ 0.00" or "1- Interest income 123.45"
            r'1[-\s]+[Ii]nterest\s+[Ii]ncome[^\d$]*\$?\s*([\d,]+\.\d{2})',
            # Fidelity: "Interest income $ 3.00"
            r'[Ii]nterest\s+[Ii]ncome[^\d$]*\$?\s*([\d,]+\.\d{2})',
            # Box 1 interest
            r'[Bb]ox\s*1[^\da][^\d]*([\d,]+\.\d{2})',
        ]

        div_amt = 0
        for pattern in div_patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                try:
                    div_amt = float(match.group(1).replace(',', ''))
                    break
                except (ValueError, AttributeError):
                    continue

        int_amt = 0
        for pattern in int_patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                try:
                    int_amt = float(match.group(1).replace(',', ''))
                    break
                except (ValueError, AttributeError):
                    continue

        payer = extract_payer_from_text(text)

        # Classify based on which has higher value
        if div_amt > int_amt and div_amt > 0:
            return ('1099-DIV', payer, 3)
        elif int_amt > 0:
            return ('1099-INT', payer, 2)
        elif has_div:
            return ('1099-DIV', payer, 3)
        else:
            return ('1099-INT', payer, 2)

    if has_int:
        payer = extract_payer_from_text(text)
        return ('1099-INT', payer, 2)

    if has_div:
        payer = extract_payer_from_text(text)
        return ('1099-DIV', payer, 3)

    if '1099-R' in text_upper or 'DISTRIBUTIONS FROM PENSIONS' in text_upper:
        payer = extract_payer_from_text(text)
        return ('1099-R', payer, 4)

    if 'SSA-1099' in text_upper or 'SOCIAL SECURITY BENEFIT' in text_upper:
        return ('SSA-1099', 'Social Security Admin', 5)

    if '1099-B' in text_upper or 'PROCEEDS FROM BROKER' in text_upper:
        payer = extract_payer_from_text(text)
        return ('1099-B', payer, 6)

    if '1098-T' in text_upper or 'TUITION STATEMENT' in text_upper:
        payer = extract_payer_from_text(text)
        return ('1098-T', payer, 7)

    if '1098' in text_upper and 'MORTGAGE' in text_upper:
        payer = extract_payer_from_text(text)
        return ('1098', payer, 7)

    if '1099-Q' in text_upper or 'QUALIFIED EDUCATION' in text_upper:
        payer = extract_payer_from_text(text)
        return ('1099-Q', payer, 8)

    if '1099-NEC' in text_upper or 'NONEMPLOYEE COMPENSATION' in text_upper:
        payer = extract_payer_from_text(text)
        return ('1099-NEC', payer, 8)

    if '1099-MISC' in text_upper:
        payer = extract_payer_from_text(text)
        return ('1099-MISC', payer, 8)

    if '1099-G' in text_upper or 'GOVERNMENT PAYMENTS' in text_upper or 'UNEMPLOYMENT COMPENSATION' in text_upper:
        payer = extract_payer_from_text(text)
        return ('1099-G', payer, 8)

    if 'K-1' in text_upper or 'SCHEDULE K-1' in text_upper:
        payer = extract_payer_from_text(text)
        return ('K-1', payer, 8)

    if '1095' in text_upper or 'HEALTH COVERAGE' in text_upper:
        return ('1095', '', 9)

    if 'PROPERTY TAX' in text_upper or 'AD VALOREM' in text_upper or 'TAX COLLECTOR' in text_upper:
        return ('Property Tax', '', 9)

    return (None, '', 99)


def extract_payer_from_text(text):
    """Extract payer/institution name from document text."""
    import re

    # Garbage words that indicate we matched form instructions, not a name
    garbage_words = ['zip', 'foreign', 'postal', 'telephone', 'province', 'omb',
                     'payer', 'form ', 'instructions', 'copy b', 'recipient']

    # Try form-header patterns first (most reliable)
    header_patterns = [
        # After "1099-INT\n" or "1099-DIV\n" header
        r'1099-(?:INT|DIV|R|NEC|G)\s*\n\s*([A-Z][A-Za-z0-9\s\.,&-]+?)(?:\s+Form|\s+\d|\n)',
        # First line: institution name (CAPITAL ONE N.A., VANGUARD MARKETING, etc.)
        r'^([A-Z][A-Z\s\.,]+(?:N\.A\.|BANK|SAVINGS|CREDIT UNION|FINANCIAL|INC|LLC|CORP)\.?)',
    ]
    for pattern in header_patterns:
        match = re.search(pattern, text, re.MULTILINE)
        if match:
            name = match.group(1).strip()[:40]
            if not any(bad in name.lower() for bad in garbage_words) and len(name) > 3:
                return name

    # Common patterns for payer names
    patterns = [
        r"PAYER'?S?\s+NAME[:\s]*\n?\s*([A-Z][A-Za-z0-9\s\.,&-]+)",
        r"(VANGUARD[A-Z\s]*)",
        r"(FIDELITY[A-Z\s]*)",
        r"(CHASE[A-Z\s]*)",
        r"(SCHWAB[A-Z\s]*)",
        r"(WELLS FARGO[A-Z\s]*)",
        r"(BANK OF AMERICA[A-Z\s]*)",
        r"(CAPITAL ONE[A-Z\s\.]*)",
        r"(MORGAN STANLEY[A-Z\s]*)",
        r"(TD AMERITRADE[A-Z\s]*)",
        r"(E\*?TRADE[A-Z\s]*)",
    ]

    text_upper = text.upper()
    for pattern in patterns:
        match = re.search(pattern, text_upper)
        if match:
            name = match.group(1).strip()[:40]
            if not any(bad in name.lower() for bad in garbage_words):
                return name

    return ""


def classify_pdf(filename, pdf_path=None):
    """
    Classify a PDF by tax form type based on filename and content.
    Returns (category, payer, priority) where lower priority = earlier in order.
    """
    fname_upper = filename.upper()
    payer = ""

    # First try filename-based classification for clear cases
    # W-2 (priority 1)
    if 'W2' in fname_upper or 'W-2' in fname_upper:
        return ('W-2', payer, 1)

    # 1099-INT (priority 2)
    if '1099-INT' in fname_upper or '1099INT' in fname_upper:
        return ('1099-INT', payer, 2)

    # 1099-DIV (priority 3)
    if '1099-DIV' in fname_upper or '1099DIV' in fname_upper:
        return ('1099-DIV', payer, 3)

    # 1099-R (priority 4)
    if '1099-R' in fname_upper or '1099R' in fname_upper:
        return ('1099-R', payer, 4)

    # SSA-1099 (priority 5)
    if 'SSA' in fname_upper or 'SSI' in fname_upper or 'SOCIAL SECURITY' in fname_upper:
        return ('SSA-1099', payer, 5)

    # 1099-B / Brokerage (priority 6)
    if '1099-B' in fname_upper or '1099B' in fname_upper:
        return ('1099-B', payer, 6)

    # 1099-Q (priority 8 - education)
    if '1099-Q' in fname_upper or '1099Q' in fname_upper:
        return ('1099-Q', payer, 8)

    # 1098-T Tuition (priority 7)
    if '1098-T' in fname_upper or '1098T' in fname_upper:
        return ('1098-T', payer, 7)

    # 1098 Mortgage (priority 7)
    if '1098' in fname_upper:
        return ('1098', payer, 7)

    # K-1 (priority 8)
    if 'K-1' in fname_upper or 'K1' in fname_upper:
        return ('K-1', payer, 8)

    # For ambiguous "1099" filenames, try to read the content
    if '1099' in fname_upper and pdf_path and PDFPLUMBER_AVAILABLE:
        text = extract_pdf_text(pdf_path)
        if text:
            form_type, payer, priority = classify_by_content(text)
            if form_type:
                return (form_type, payer, priority)
        # Fallback to 1099-Other
        return ('1099-Other', payer, 8)

    # Other 1099s (priority 8)
    if '1099' in fname_upper:
        return ('1099-Other', payer, 8)

    # For unknown files, try content-based classification
    if pdf_path and PDFPLUMBER_AVAILABLE:
        text = extract_pdf_text(pdf_path)
        if text:
            form_type, payer, priority = classify_by_content(text)
            if form_type:
                return (form_type, payer, priority)

    # Everything else (priority 9)
    return ('Other', payer, 9)


def detect_multi_form_contents(pdf_path):
    """For multi-form PDFs, detect what form types are inside."""
    if not PDFPLUMBER_AVAILABLE:
        return []

    try:
        forms_found = {}  # {form_type: count}
        seen_eins = {}    # {ein: form_type} to avoid double-counting copies
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                if len(text.strip()) < 50:
                    continue

                text_upper = text.upper()
                ein_match = re.search(r'(\d{2}-\d{7})', text)
                ein = ein_match.group(1) if ein_match else None

                # Determine form type for this page
                form_type = None
                if '1099-G' in text_upper or 'UNEMPLOYMENT COMPENSATION' in text_upper:
                    form_type = '1099-G'
                elif '1099-NEC' in text_upper or 'NONEMPLOYEE COMPENSATION' in text_upper:
                    form_type = '1099-NEC'
                elif ('WAGE AND TAX' in text_upper or 'FORM W-2' in text_upper or
                      'WAGES, TIPS' in text_upper or
                      bool(re.search(r'\d{2}-\d{7}\s*\n\s*[\d,]+\.\d{2}\s+[\d,]+\.\d{2}', text))):
                    form_type = 'W-2'

                if form_type and ein:
                    key = (ein, form_type)
                    if key not in seen_eins:
                        seen_eins[key] = True
                        forms_found[form_type] = forms_found.get(form_type, 0) + 1

        return [(ft, count) for ft, count in forms_found.items()]
    except Exception:
        return []


def classify_pdf_pages(pdf_path):
    """
    Classify individual pages of a multi-page PDF to detect mixed form types.
    Returns list of (page_indices, category, payer, priority) groups,
    or None if the PDF should be treated as a single document.
    """
    if not PDFPLUMBER_AVAILABLE:
        return None

    try:
        with pdfplumber.open(pdf_path) as pdf:
            num_pages = len(pdf.pages)
            if num_pages <= 2:
                return None

            # Classify each page
            page_classes = []
            for i, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                if len(text.strip()) < 50:
                    page_classes.append((i, None, '', 99))
                    continue
                form_type, payer, priority = classify_by_content(text)
                page_classes.append((i, form_type, payer, priority))
    except Exception:
        return None

    # Check if multiple form types exist
    types_found = set(t[1] for t in page_classes if t[1] is not None)
    if len(types_found) <= 1:
        return None  # Single form type, use file-level classification

    # Group consecutive pages by form type
    groups = []
    current_pages = [page_classes[0][0]]
    current_type = page_classes[0][1]
    current_payer = page_classes[0][2]
    current_priority = page_classes[0][3]

    for i in range(1, len(page_classes)):
        idx, form_type, payer, priority = page_classes[i]

        if form_type is None:
            # Unclassified page - attach to current group
            current_pages.append(idx)
        elif form_type == current_type:
            # Same type - continue group
            current_pages.append(idx)
            if payer and not current_payer:
                current_payer = payer
        else:
            # New form type - save current group, start new one
            groups.append((list(current_pages), current_type or 'Other',
                          current_payer, current_priority))
            current_pages = [idx]
            current_type = form_type
            current_payer = payer
            current_priority = priority

    # Save last group
    groups.append((list(current_pages), current_type or 'Other',
                  current_payer, current_priority))

    return groups


def sort_pdfs_by_category(pdf_files, file_origins=None):
    """
    Sort PDF files by tax form category, reading content when needed.
    Multi-form PDFs are split into separate page groups.
    Returns list of (priority, category, payer, pdf_path, page_indices) tuples.
    page_indices is None for whole files, or a list of page numbers for splits.
    """
    if file_origins is None:
        file_origins = {}
    categorized = []
    print("  Analyzing document contents...")

    for pdf_path in pdf_files:
        # Try page-level classification for multi-form detection
        page_groups = classify_pdf_pages(pdf_path)

        if page_groups and len(page_groups) > 1:
            # Multi-form PDF - add each group separately
            group_desc = ", ".join(f"{cat}" for _, cat, _, _ in page_groups)
            print(f"    Split {pdf_path.name} -> [{group_desc}]")
            for pages, category, payer, priority in page_groups:
                if category == 'Other' and file_origins:
                    origin = file_origins.get(str(pdf_path), 'original')
                    if origin == 'image':
                        category = 'Other (Photo)'
                        priority = 10
                    elif origin == 'office':
                        category = 'Other (Office)'
                        priority = 11
                categorized.append((priority, category, payer, pdf_path, pages))
        else:
            # Single form type or small file - classify whole file
            category, payer, priority = classify_pdf(pdf_path.name, pdf_path)

            if category == 'Other' and file_origins:
                origin = file_origins.get(str(pdf_path), 'original')
                if origin == 'image':
                    category = 'Other (Photo)'
                    priority = 10
                elif origin == 'office':
                    category = 'Other (Office)'
                    priority = 11

            categorized.append((priority, category, payer, pdf_path, None))

    # Sort by priority, then by filename
    categorized.sort(key=lambda x: (x[0], x[3].name.lower()))
    return categorized


# =============================================================================
# PDF COMBINATION
# =============================================================================

def combine_pdfs(categorized, output_path):
    """
    Combine PDFs into a single document, normalizing all pages to letter size.
    Accepts categorized entries with page indices for page-level sorting.
    Only auto-rotates landscape pages that have no readable text (sideways scans).
    """
    writer = PdfWriter()

    for priority, category, payer, pdf_path, page_indices in categorized:
        try:
            reader = PdfReader(str(pdf_path))

            # Determine which pages to include
            if page_indices is None:
                pages_to_process = list(range(len(reader.pages)))
            else:
                pages_to_process = page_indices

            # Check which pages have extractable text (for rotation decisions)
            page_has_text = {}
            if PDFPLUMBER_AVAILABLE:
                try:
                    with pdfplumber.open(str(pdf_path)) as plumber:
                        for i in pages_to_process:
                            if i < len(plumber.pages):
                                text = plumber.pages[i].extract_text() or ""
                                page_has_text[i] = len(text.strip()) > 100
                except Exception:
                    pass

            for i in pages_to_process:
                if i < len(reader.pages):
                    page = reader.pages[i]
                    # Don't auto-rotate pages with readable text (intentionally landscape)
                    has_text = page_has_text.get(i, False)
                    page = normalize_pdf_page(page, auto_rotate=not has_text)
                    writer.add_page(page)
        except Exception as e:
            print(f"  WARNING: Could not read {pdf_path.name}: {e}")
            continue

    with open(output_path, 'wb') as output_file:
        writer.write(output_file)

    return output_path


# =============================================================================
# GOOGLE SHEETS INTEGRATION
# =============================================================================

def get_gsheet_client(credentials_path=None):
    """Get authenticated Google Sheets client."""
    if not GSHEETS_AVAILABLE:
        print("ERROR: gspread not installed. Run: pip install gspread google-auth")
        return None

    # Look for credentials file
    if credentials_path is None:
        possible_paths = [
            Path.home() / '.config' / 'gspread' / 'service_account.json',
            Path.home() / 'gspread_credentials.json',
            Path('C:/Tax/google_credentials.json'),
        ]
        for p in possible_paths:
            if p.exists():
                credentials_path = p
                break

    if credentials_path is None or not Path(credentials_path).exists():
        print("ERROR: Google credentials file not found.")
        print("  Expected locations:")
        print("    - ~/.config/gspread/service_account.json")
        print("    - C:/Tax/google_credentials.json")
        return None

    try:
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        creds = Credentials.from_service_account_file(str(credentials_path), scopes=scopes)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        print(f"ERROR: Could not authenticate with Google: {e}")
        return None


def check_expected_documents(gsheet_client, sheet_id, client_name, found_docs):
    """
    Check Google Sheet 'Document Details' for expected documents.

    Sheet format:
    Column A: Client Name
    Column B: Form Expected (W-2, 1099-DIV, etc.)
    Column C: Payer/Source
    Column D: T/S/J
    Column E: Prior Year Amount
    Column F: Category
    Column G: Received (?, Y, N)
    Column H: Notes

    Args:
        gsheet_client: Authenticated gspread client
        sheet_id: Google Sheet ID
        client_name: Client name from folder (e.g., "DiMascio Jim and Leila")
        found_docs: Dict of found documents {form_type: [filenames]}
    """
    try:
        spreadsheet = gsheet_client.open_by_key(sheet_id)
        ws = spreadsheet.worksheet('Document Details')

        all_data = ws.get_all_values()
        headers = all_data[0] if all_data else []

        # Find column indices
        col_map = {h.lower(): i for i, h in enumerate(headers)}
        client_col = col_map.get('client name', 0)
        form_col = col_map.get('form expected', 1)
        payer_col = col_map.get('payer/source', 2)
        received_col = col_map.get('received', 6)
        notes_col = col_map.get('notes', 7)

        # Normalize client name for matching
        # "DiMascio Jim and Leila" -> match "DIMASCIO, JAMES & LEILA"
        client_parts = client_name.upper().replace(' AND ', ' ').replace('&', ' ').split()

        # Find all rows for this client
        client_rows = []
        for i, row in enumerate(all_data[1:], start=2):  # start=2 for 1-indexed + skip header
            if len(row) > client_col:
                sheet_client = row[client_col].upper()
                # Check if any significant part of the name matches
                if any(part in sheet_client for part in client_parts if len(part) > 2):
                    client_rows.append((i, row))

        if not client_rows:
            print(f"  Client '{client_name}' not found in Document Details")
            return None

        print(f"  Found {len(client_rows)} expected documents for client")

        # Check each expected document
        missing = []
        found_count = 0
        updates = []

        for row_num, row in client_rows:
            form_type = row[form_col] if len(row) > form_col else ''
            expected_payer = row[payer_col] if len(row) > payer_col else ''
            current_received = row[received_col] if len(row) > received_col else '?'

            # Check if we found this form type (and optionally match payer)
            form_found = False
            matched_payer = ""

            for found_type, file_list in found_docs.items():
                # Check if form types match
                type_match = (
                    form_type.upper() in found_type.upper() or
                    found_type.upper() in form_type.upper() or
                    (form_type.upper() == 'W-2' and found_type.upper() == 'W-2') or
                    (form_type.upper().replace('-', '') == found_type.upper().replace('-', ''))
                )

                if type_match:
                    # Check if any file matches the expected payer
                    for filename, found_payer in file_list:
                        # Normalize payer names for comparison
                        exp_payer_norm = expected_payer.upper().split('-')[0].split('#')[0].strip()
                        found_payer_norm = found_payer.upper() if found_payer else ""
                        filename_upper = filename.upper()

                        # Match by payer name in found_payer or filename
                        if (exp_payer_norm and (
                            exp_payer_norm in found_payer_norm or
                            exp_payer_norm in filename_upper or
                            any(word in found_payer_norm or word in filename_upper
                                for word in exp_payer_norm.split() if len(word) > 3)
                        )) or not exp_payer_norm:
                            form_found = True
                            matched_payer = found_payer or filename
                            break

                    # If no payer match but form type matches, still count it
                    if not form_found and file_list:
                        form_found = True
                        matched_payer = file_list[0][1] or file_list[0][0]

                if form_found:
                    break

            if form_found:
                found_count += 1
                if current_received == '?':
                    updates.append((row_num, received_col + 1, 'Y'))  # +1 for 1-indexed
            else:
                missing.append(f"{form_type} ({expected_payer})")
                if current_received == '?':
                    updates.append((row_num, received_col + 1, 'N'))

        # Apply updates
        for row_num, col_num, value in updates:
            ws.update_cell(row_num, col_num, value)

        print(f"  Found: {found_count}/{len(client_rows)} expected documents")

        if missing:
            print(f"  MISSING:")
            for m in missing:
                print(f"    - {m}")

            # Update Client Tracker notes
            try:
                tracker = spreadsheet.worksheet('Client Tracker')
                # Find client row
                tracker_data = tracker.get_all_values()
                for i, row in enumerate(tracker_data[1:], start=2):
                    if len(row) > 0:
                        sheet_client = row[0].upper()
                        if any(part in sheet_client for part in client_parts if len(part) > 2):
                            # Found client - update notes column (column K = 11)
                            notes_text = f"[{datetime.now().strftime('%m/%d')}] Missing: {', '.join([m.split('(')[0].strip() for m in missing[:3]])}"
                            if len(missing) > 3:
                                notes_text += f" +{len(missing)-3} more"

                            current_notes = row[10] if len(row) > 10 else ''
                            if current_notes:
                                notes_text = f"{current_notes} | {notes_text}"

                            tracker.update_cell(i, 11, notes_text)
                            print(f"  Updated Client Tracker notes")
                            break
            except Exception as e:
                print(f"  Could not update Client Tracker: {e}")

        return missing

    except Exception as e:
        print(f"  ERROR accessing Google Sheet: {e}")
        import traceback
        traceback.print_exc()
        return None


# =============================================================================
# MAIN WORKFLOW
# =============================================================================

def process_folder(folder_path, google_sheet_id=None, credentials_path=None):
    """Process a client folder - combine PDFs and check expected documents."""
    folder = Path(folder_path)

    if not folder.exists():
        print(f"ERROR: Folder not found: {folder}")
        return None

    # Get client name from folder path
    client_name = folder.parent.name
    year = folder.name

    print("=" * 60)
    print("DOCUMENT REVIEW")
    print("=" * 60)
    print(f"Client: {client_name}")
    print(f"Year: {year}")
    print(f"Folder: {folder}")
    print()

    # Convert any image files (HEIC, JPG, etc.) to PDF first
    converted_images = convert_images_in_folder(folder)

    # Convert any Office files (Word, Excel) to PDF
    converted_office = convert_office_files_in_folder(folder)

    # Track file origins for sort priority of unidentified docs
    file_origins = {}
    for p in converted_images:
        file_origins[str(p)] = 'image'
    for p in converted_office:
        file_origins[str(p)] = 'office'

    # Find all PDFs (exclude previously generated combined PDFs)
    pdf_files = list(folder.glob('*.pdf')) + list(folder.glob('*.PDF'))
    pdf_files = list(set(pdf_files))  # Remove duplicates
    pdf_files = [f for f in pdf_files if '_document_review' not in f.name.lower()]

    if not pdf_files:
        print("No PDF files found!")
        return None

    print(f"Found {len(pdf_files)} PDF files")
    print()

    # Classify and sort PDFs
    categorized = sort_pdfs_by_category(pdf_files, file_origins)

    # Print categorization and build found_docs dict
    print("\nDocument Classification:")
    print("-" * 40)
    current_category = None
    found_categories = set()
    found_docs = {}  # {category: [(filename, payer)]}

    for priority, category, payer, pdf_path, page_indices in categorized:
        if category != current_category:
            print(f"\n  [{category}]")
            current_category = category
        found_categories.add(category)
        if category not in found_docs:
            found_docs[category] = []
        found_docs[category].append((pdf_path.name, payer))
        payer_str = f" ({payer})" if payer else ""
        pages_str = f" (pages {page_indices[0]+1}-{page_indices[-1]+1})" if page_indices else ""
        print(f"    - {pdf_path.name}{payer_str}{pages_str}")
    print()

    # Combine PDFs - client_name_year_document_review.pdf
    output_filename = f"{client_name}_{year}_document_review.pdf"
    output_path = folder / output_filename

    print(f"Combining PDFs...")
    try:
        combine_pdfs(categorized, output_path)
        print(f"  Created: {output_path}")
        print(f"  Size: {output_path.stat().st_size / 1024:.1f} KB")
    except PermissionError:
        print(f"  WARNING: Could not write {output_path} (file may be open)")
        print(f"  Close the file and rerun, or output will be skipped.")
    print()

    # Check against Google Sheet if provided
    if google_sheet_id:
        print("Checking Google Sheet for expected documents...")
        gsheet_client = get_gsheet_client(credentials_path)
        if gsheet_client:
            missing = check_expected_documents(gsheet_client, google_sheet_id, client_name, found_docs)

    # Summary
    print()
    print("=" * 60)
    print("SUMMARY")
    print("=" * 60)
    num_entries = len(categorized)
    print(f"Document groups combined: {num_entries}")
    print(f"Output file: {output_filename}")
    print()
    print("Categories found:")
    for cat in sorted(found_categories):
        count = len([c for _, c, _, _, _ in categorized if c == cat])
        print(f"  {cat}: {count}")

    return output_path


def main():
    parser = argparse.ArgumentParser(
        description='Combine client PDFs in tax-form order and check for missing documents',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python document_review.py "L:\\DiMascio Jim and Leila\\2025"
  python document_review.py "L:\\Client\\2025" --google-sheet 1ABC123xyz

Document Order:
  1. W-2s
  2. 1099-INT
  3. 1099-DIV
  4. 1099-R
  5. SSA-1099
  6. 1099-B (Brokerage)
  7. 1098
  8. Other 1099s
  9. Everything else
        """
    )

    parser.add_argument('folder', help='Client folder path (e.g., "L:\\Client\\2025")')
    parser.add_argument('--google-sheet', '-g', dest='sheet_id',
                        help='Google Sheet ID for expected document tracking')
    parser.add_argument('--credentials', '-c',
                        help='Path to Google service account credentials JSON')

    args = parser.parse_args()

    process_folder(args.folder, args.sheet_id, args.credentials)


if __name__ == '__main__':
    main()
