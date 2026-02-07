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

try:
    from pypdf import PdfReader, PdfWriter
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

# Optional Google Sheets support
GSHEETS_AVAILABLE = False
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSHEETS_AVAILABLE = True
except ImportError:
    pass


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


def sort_pdfs_by_category(pdf_files):
    """Sort PDF files by tax form category, reading content when needed."""
    categorized = []
    multi_form_details = {}  # {pdf_path: [(form_type, count)]}
    print("  Analyzing document contents...")

    for pdf_path in pdf_files:
        category, payer, priority = classify_pdf(pdf_path.name, pdf_path)

        # Check for multi-form PDFs (e.g., "all W2.pdf" with multiple forms inside)
        if category == 'W-2' and PDFPLUMBER_AVAILABLE:
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    if len(pdf.pages) > 5:
                        # Large PDF - check for multiple form types
                        text = ""
                        for page in pdf.pages[:20]:
                            text += (page.extract_text() or "") + "\n"
                        eins = set(re.findall(r'\d{2}-\d{7}', text))
                        if len(eins) > 1:
                            contents = detect_multi_form_contents(pdf_path)
                            if contents:
                                multi_form_details[str(pdf_path)] = contents
            except Exception:
                pass

        categorized.append((priority, category, payer, pdf_path))

    # Sort by priority, then by filename
    categorized.sort(key=lambda x: (x[0], x[3].name.lower()))
    return categorized, multi_form_details


# =============================================================================
# PDF COMBINATION
# =============================================================================

def combine_pdfs(pdf_files, output_path):
    """Combine multiple PDFs into a single document."""
    writer = PdfWriter()

    for pdf_path in pdf_files:
        try:
            reader = PdfReader(str(pdf_path))
            for page in reader.pages:
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

    # Find all PDFs (exclude previously generated combined PDFs)
    pdf_files = list(folder.glob('*.pdf')) + list(folder.glob('*.PDF'))
    pdf_files = list(set(pdf_files))  # Remove duplicates
    pdf_files = [f for f in pdf_files if '_document_review.pdf' not in f.name.lower()]

    if not pdf_files:
        print("No PDF files found!")
        return None

    print(f"Found {len(pdf_files)} PDF files")
    print()

    # Classify and sort PDFs
    categorized, multi_form_details = sort_pdfs_by_category(pdf_files)

    # Print categorization and build found_docs dict
    print("\nDocument Classification:")
    print("-" * 40)
    current_category = None
    found_categories = set()
    found_docs = {}  # {category: [(filename, payer)]}

    for priority, category, payer, pdf_path in categorized:
        if category != current_category:
            print(f"\n  [{category}]")
            current_category = category
        found_categories.add(category)
        if category not in found_docs:
            found_docs[category] = []
        found_docs[category].append((pdf_path.name, payer))
        payer_str = f" ({payer})" if payer else ""
        print(f"    - {pdf_path.name}{payer_str}")
        # Show multi-form contents if detected
        contents = multi_form_details.get(str(pdf_path))
        if contents:
            forms_desc = ", ".join(f"{count}x {ft}" for ft, count in contents)
            print(f"      ^ Multi-form PDF contains: {forms_desc}")
            # Add sub-forms to found_docs for Google Sheet matching
            for ft, count in contents:
                if ft != category:
                    found_categories.add(ft)
                    if ft not in found_docs:
                        found_docs[ft] = []
                    found_docs[ft].append((pdf_path.name, f"(inside {pdf_path.name})"))
    print()

    # Combine PDFs - client_name_year_document_review.pdf
    output_filename = f"{client_name}_{year}_document_review.pdf"
    output_path = folder / output_filename

    print(f"Combining PDFs...")
    sorted_pdfs = [pdf_path for _, _, _, pdf_path in categorized]
    try:
        combine_pdfs(sorted_pdfs, output_path)
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
    print(f"Documents combined: {len(sorted_pdfs)}")
    print(f"Output file: {output_filename}")
    print()
    print("Categories found:")
    for cat in sorted(found_categories):
        count = len([c for _, c, _, _ in categorized if c == cat])
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
