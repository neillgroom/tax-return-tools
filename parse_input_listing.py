#!/usr/bin/env python3
"""
CCH Input Listing Parser
========================
Parses CCH ProSystem fx Input Listing export and outputs structured JSON.

Usage:
    python parse_input_listing.py input_listing.pdf
    python parse_input_listing.py input_listing.pdf --output data.json
    python parse_input_listing.py input_listing.pdf --format text  (if already extracted to text)

Output:
    Creates JSON file with all parsed data, ready for checksheet population.
"""

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path
from collections import defaultdict
from datetime import datetime


# =============================================================================
# CCH FIELD MAPPINGS (Validated via Gemini/CCH Documentation)
# =============================================================================

CCH_FIELD_MAP = {
    'W-2': {
        'name': 'W-2 Wage Statement',
        'fields': {
            40: 'employer_name',
            50: 'box1_wages',
            51: 'box2_fed_withholding',
            52: 'box3_ss_wages',
            54: 'box5_medicare_wages',
            68: 'box12_code',
            69: 'box12_amount',
        }
    },
    'IRS-1099INT': {
        'name': '1099-INT Interest',
        'fields': {
            40: 'payer_name',
            71: 'box1_interest',
            74: 'box4_fed_withholding',
            78: 'box8_tax_exempt_interest',
        }
    },
    'IRS-1099DIV': {
        'name': '1099-DIV Dividends',
        'fields': {
            40: 'payer_name',
            70: 'box1a_ordinary_dividends',
            71: 'box1b_qualified_dividends',
            72: 'box2a_total_cap_gain',
            74: 'box4_fed_withholding',
        }
    },
    'IRS-1099R': {
        'name': '1099-R Retirement',
        'fields': {
            40: 'payer_name',
            70: 'box1_gross_distribution',
            71: 'box2a_taxable_amount',
            74: 'box4_fed_withholding',
            81: 'box7_distribution_code',
            85: 'ira_sep_simple',
        }
    },
    'SSA-1099': {
        'name': 'SSA-1099 Social Security',
        'fields': {
            30: 'description',
            33: 'box5_net_benefits',
            35: 'box6_fed_withholding',
        }
    },
    'IRS-1098': {
        'name': '1098 Mortgage Interest',
        'fields': {
            40: 'lender_name',
            70: 'box1_mortgage_interest',
            71: 'box6_points',
        }
    },
    'D-1': {
        'name': 'Schedule D Capital Transactions',
        'fields': {
            30: 'description',
            31: 'shares',
            32: 'date_acquired',
            33: 'date_sold',
            34: 'proceeds',
            35: 'cost_basis',
            44: 'term',  # S=short, L=long
            100: 'proceeds_actual',
            101: 'cost_actual',
        }
    },
    'C-1': {
        'name': 'Schedule C Business',
        'fields': {
            30: 'business_description',
            31: 'naics_code',
            36: 'business_name',
            37: 'address',
            38: 'city',
            39: 'state',
            40: 'zip',
            80: 'gross_receipts',
        }
    },
    'C-2': {
        'name': 'Schedule C Expenses',
        'fields': {
            42: 'contract_labor',
            55: 'wages',
            61: 'other_expense_description',
            63: 'other_expense_amount',
        }
    },
    'E-1': {
        'name': 'Schedule E Rental',
        'fields': {
            30: 'property_type',
            32: 'state',
            42: 'address',
            43: 'city',
            44: 'state2',
            45: 'zip',
            52: 'line12_mortgage_interest',
            60: 'line3_rents_received',
            69: 'line5_advertising',
            72: 'line7_cleaning_maintenance',
            75: 'line8_commissions',
            78: 'line9_insurance',
            81: 'line10_legal_professional',
            84: 'line16_taxes',
            87: 'line13_other_interest',
            102: 'line14_repairs',
            105: 'line15_supplies',
            108: 'line17_utilities',
            111: 'line18_depreciation',
            140: 'line19_other_description',
            141: 'line19_other_amount',
        }
    },
    'K-1': {
        'name': 'K-1 Pass-through (1065/1120S)',
        'fields': {
            # 1065 Partnership
            43: 'box1_ordinary_income_1065',
            44: 'box2_rental_income_1065',
            52: 'box5_interest_1065',
            53: 'box6a_dividends_1065',
            # 1120S S-Corp
            103: 'box1_ordinary_income_1120s',
            104: 'box2_rental_income_1120s',
            111: 'box4_interest_1120s',
            112: 'box5a_dividends_1120s',
        }
    },
    'IRS-K1 1041': {
        'name': 'K-1 Estate/Trust',
        'fields': {
            40: 'ein',
            41: 'entity_name',
            42: 'entity_type',
            104: 'box1_interest',
            340: 'box3_net_st_cap_gain',
            380: 'box5_other_portfolio',
        }
    },
    'DP-1': {
        'name': 'Depreciation Asset',
        'fields': {
            30: 'asset_number',
            31: 'description',
            32: 'date_placed_in_service',
            33: 'method',
            34: 'life_years',
            35: 'cost_basis',
            47: 'section_179',
            49: 'current_depreciation',
            54: 'bonus_depreciation',
        }
    },
    'NJ1': {
        'name': 'NJ-1040 Resident Return',
        'fields': {
            30: 'filing_status',
            33: 'exemptions',
            36: 'county_code',
            39: 'municipality_code',
        }
    },
}


def extract_text_from_pdf(pdf_path: Path) -> str:
    """Extract text from PDF using pdftotext."""
    try:
        result = subprocess.run(
            ['pdftotext', '-layout', str(pdf_path), '-'],
            capture_output=True,
            text=True,
            check=True
        )
        return result.stdout
    except FileNotFoundError:
        print("ERROR: pdftotext not found. Install poppler-utils:")
        print("  Ubuntu/Debian: sudo apt-get install poppler-utils")
        print("  Mac: brew install poppler")
        print("  Windows: Download from https://github.com/oschwartz10612/poppler-windows")
        sys.exit(1)
    except subprocess.CalledProcessError as e:
        print(f"ERROR: Failed to extract text from PDF: {e}")
        sys.exit(1)


def parse_field_value(value_str: str):
    """Parse a field value, converting to appropriate type."""
    # Remove quotes if present
    if value_str.startswith('"') and value_str.endswith('"'):
        return value_str[1:-1]
    
    # Try to convert to number
    try:
        if '.' in value_str:
            return float(value_str)
        else:
            return int(value_str)
    except ValueError:
        return value_str


def parse_input_listing(text: str) -> dict:
    """
    Parse CCH Input Listing text into structured data.
    
    Returns dict with:
        - 'forms': dict of form_type -> list of entries
        - 'summary': summary totals
        - 'metadata': parsing info
    """
    results = {
        'forms': defaultdict(list),
        'metadata': {
            'parsed_at': datetime.now().isoformat(),
            'parser_version': '1.0',
        }
    }
    
    # Split by separator lines
    sections = re.split(r'~{50,}', text)
    
    for section in sections:
        section = section.strip()
        if not section:
            continue
        
        lines = section.split('\n')
        if not lines:
            continue
        
        # Parse header: "IRS-1099INT, Sheet #1, Entity 1 Box Cnt 3"
        header_match = re.match(
            r'^([A-Za-z0-9\-\s]+),\s*Sheet\s*#(\d+),\s*Entity\s*(\d+)',
            lines[0]
        )
        
        if not header_match:
            continue
        
        form_type = header_match.group(1).strip()
        sheet_num = int(header_match.group(2))
        entity_num = int(header_match.group(3))
        
        # Initialize entry
        entry = {
            '_form_type': form_type,
            '_sheet': sheet_num,
            '_entity': entity_num,
            '_raw_fields': {},
        }
        
        # Combine all field data lines
        field_text = ' '.join(lines[1:])
        
        # Parse fields: "40: \"BCB COMMUNITY BANK\", 71: 10041"
        # Pattern handles: quoted strings, negative numbers, decimals, dates
        field_pattern = r'(\d+):\s*(?:"([^"]+)"|(\d+/\s*\d+/\d+)|(-?\d+\.?\d*))'
        
        for match in re.finditer(field_pattern, field_text):
            field_num = int(match.group(1))
            
            # Get the value from whichever group matched
            if match.group(2):  # Quoted string
                value = match.group(2)
            elif match.group(3):  # Date
                value = match.group(3)
            elif match.group(4):  # Number
                value = parse_field_value(match.group(4))
            else:
                continue
            
            entry['_raw_fields'][field_num] = value
            
            # Map to friendly name if we know the form type
            if form_type in CCH_FIELD_MAP:
                field_map = CCH_FIELD_MAP[form_type]['fields']
                if field_num in field_map:
                    friendly_name = field_map[field_num]
                    entry[friendly_name] = value
        
        results['forms'][form_type].append(entry)
    
    return results


def calculate_summary(parsed_data: dict) -> dict:
    """Calculate summary totals from parsed data."""
    summary = {
        'income': {},
        'withholding': {},
        'schedules': {},
    }
    
    forms = parsed_data['forms']
    
    # 1099-INT totals
    total_interest = 0
    for entry in forms.get('IRS-1099INT', []):
        total_interest += entry.get('box1_interest', 0) or 0
    summary['income']['total_interest'] = total_interest
    
    # 1099-DIV totals
    total_ordinary_div = 0
    total_qualified_div = 0
    for entry in forms.get('IRS-1099DIV', []):
        total_ordinary_div += entry.get('box1a_ordinary_dividends', 0) or 0
        total_qualified_div += entry.get('box1b_qualified_dividends', 0) or 0
    summary['income']['total_ordinary_dividends'] = total_ordinary_div
    summary['income']['total_qualified_dividends'] = total_qualified_div
    
    # W-2 totals
    total_wages = 0
    total_fed_wh = 0
    for entry in forms.get('W-2', []):
        total_wages += entry.get('box1_wages', 0) or 0
        total_fed_wh += entry.get('box2_fed_withholding', 0) or 0
    summary['income']['total_wages'] = total_wages
    summary['withholding']['w2_federal'] = total_fed_wh
    
    # 1099-R totals
    total_pension_taxable = 0
    for entry in forms.get('IRS-1099R', []):
        total_pension_taxable += entry.get('box2a_taxable_amount', 0) or 0
    summary['income']['total_pension_taxable'] = total_pension_taxable
    
    # SSA-1099
    total_ss_benefits = 0
    for entry in forms.get('SSA-1099', []):
        total_ss_benefits += entry.get('box5_net_benefits', 0) or 0
    summary['income']['total_ss_benefits'] = total_ss_benefits
    
    # Schedule C
    total_sch_c_gross = 0
    for entry in forms.get('C-1', []):
        total_sch_c_gross += entry.get('gross_receipts', 0) or 0
    summary['schedules']['schedule_c_gross'] = total_sch_c_gross
    
    # Schedule E
    total_rents = 0
    for entry in forms.get('E-1', []):
        total_rents += entry.get('line3_rents_received', 0) or 0
    summary['schedules']['schedule_e_rents'] = total_rents
    
    return summary


def main():
    parser = argparse.ArgumentParser(
        description='Parse CCH ProSystem fx Input Listing',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python parse_input_listing.py return_input.pdf
  python parse_input_listing.py return_input.pdf --output parsed_data.json
  python parse_input_listing.py return_input.txt --format text
  
The output JSON can be used with populate_checksheet.py to fill the verification checksheet.
        """
    )
    
    parser.add_argument(
        'input_file',
        help='Input Listing file (PDF or text)'
    )
    
    parser.add_argument(
        '--output', '-o',
        help='Output JSON file (default: <input>_parsed.json)'
    )
    
    parser.add_argument(
        '--format', '-f',
        choices=['auto', 'pdf', 'text'],
        default='auto',
        help='Input format (default: auto-detect from extension)'
    )
    
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='Print detailed parsing info'
    )
    
    args = parser.parse_args()
    
    input_path = Path(args.input_file)
    
    if not input_path.exists():
        print(f"ERROR: File not found: {input_path}")
        sys.exit(1)
    
    # Determine format
    if args.format == 'auto':
        if input_path.suffix.lower() == '.pdf':
            file_format = 'pdf'
        else:
            file_format = 'text'
    else:
        file_format = args.format
    
    # Extract text
    print(f"Reading: {input_path}")
    
    if file_format == 'pdf':
        print("Extracting text from PDF...")
        text = extract_text_from_pdf(input_path)
    else:
        text = input_path.read_text()
    
    # Parse
    print("Parsing Input Listing...")
    parsed = parse_input_listing(text)
    
    # Calculate summary
    parsed['summary'] = calculate_summary(parsed)
    
    # Convert defaultdict to regular dict for JSON serialization
    parsed['forms'] = dict(parsed['forms'])
    
    # Determine output path
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_suffix('.json')
        output_path = output_path.with_stem(input_path.stem + '_parsed')
    
    # Write output
    with open(output_path, 'w') as f:
        json.dump(parsed, f, indent=2, default=str)
    
    print(f"\nOutput: {output_path}")
    
    # Print summary
    print("\n" + "="*60)
    print("PARSING SUMMARY")
    print("="*60)
    
    for form_type, entries in parsed['forms'].items():
        form_name = CCH_FIELD_MAP.get(form_type, {}).get('name', form_type)
        print(f"  {form_name}: {len(entries)} entries")
    
    print("\n" + "-"*60)
    print("INCOME TOTALS")
    print("-"*60)
    
    summary = parsed['summary']
    if summary['income'].get('total_wages'):
        print(f"  Wages (W-2 Box 1):        ${summary['income']['total_wages']:>12,}")
    if summary['income'].get('total_interest'):
        print(f"  Interest (1099-INT):      ${summary['income']['total_interest']:>12,}")
    if summary['income'].get('total_ordinary_dividends'):
        print(f"  Dividends (1099-DIV):     ${summary['income']['total_ordinary_dividends']:>12,}")
    if summary['income'].get('total_pension_taxable'):
        print(f"  Pension (1099-R):         ${summary['income']['total_pension_taxable']:>12,}")
    if summary['income'].get('total_ss_benefits'):
        print(f"  Social Security:          ${summary['income']['total_ss_benefits']:>12,}")
    if summary['schedules'].get('schedule_c_gross'):
        print(f"  Schedule C Gross:         ${summary['schedules']['schedule_c_gross']:>12,}")
    if summary['schedules'].get('schedule_e_rents'):
        print(f"  Schedule E Rents:         ${summary['schedules']['schedule_e_rents']:>12,}")
    
    print("\n" + "="*60)
    
    if args.verbose:
        print("\nDETAILED DATA:")
        print(json.dumps(parsed, indent=2, default=str))


if __name__ == '__main__':
    main()
