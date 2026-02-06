#!/usr/bin/env python3
"""
Tax Verification Tool
=====================
Main entry point for tax verification workflow.

Usage:
    python run.py                    # Interactive menu
    python run.py parse FILE         # Parse CCH Input Listing
    python run.py fill JSON_FILE     # Fill checksheet from parsed data
    python run.py full FILE          # Parse and fill in one step

Requirements:
    pip install openpyxl
    
    For PDF support:
    - Ubuntu/Debian: sudo apt-get install poppler-utils
    - Mac: brew install poppler
    - Windows: Install poppler and add to PATH
"""

import argparse
import sys
import os
from pathlib import Path

# Add script directory to path
script_dir = Path(__file__).parent
sys.path.insert(0, str(script_dir))

from parse_input_listing import parse_input_listing, extract_text_from_pdf, calculate_summary, CCH_FIELD_MAP
from populate_checksheet import populate_checksheet

import json
from datetime import datetime

try:
    from openpyxl import load_workbook
except ImportError:
    print("ERROR: openpyxl not found. Run:")
    print("  pip install openpyxl")
    sys.exit(1)


def print_banner():
    """Print tool banner."""
    print()
    print("=" * 60)
    print("  TAX VERIFICATION TOOL")
    print("  CCH Input Listing Parser + Checksheet Populator")
    print("=" * 60)
    print()


def interactive_menu():
    """Run interactive menu."""
    print_banner()
    
    print("What would you like to do?")
    print()
    print("  1. Parse CCH Input Listing (PDF or text)")
    print("  2. Fill checksheet from parsed data (JSON)")
    print("  3. Full workflow (parse + fill)")
    print("  4. Help")
    print("  5. Exit")
    print()
    
    choice = input("Enter choice (1-5): ").strip()
    
    if choice == '1':
        do_parse_interactive()
    elif choice == '2':
        do_fill_interactive()
    elif choice == '3':
        do_full_interactive()
    elif choice == '4':
        show_help()
    elif choice == '5':
        print("Goodbye!")
        sys.exit(0)
    else:
        print("Invalid choice.")
        interactive_menu()


def do_parse_interactive():
    """Interactive parse workflow."""
    print()
    print("-" * 60)
    print("PARSE CCH INPUT LISTING")
    print("-" * 60)
    print()
    
    file_path = input("Enter path to Input Listing file (PDF or text): ").strip()
    
    if not file_path:
        print("No file specified.")
        return interactive_menu()
    
    file_path = Path(file_path)
    if not file_path.exists():
        print(f"ERROR: File not found: {file_path}")
        return interactive_menu()
    
    do_parse(file_path)
    
    print()
    input("Press Enter to continue...")
    interactive_menu()


def do_fill_interactive():
    """Interactive fill workflow."""
    print()
    print("-" * 60)
    print("FILL CHECKSHEET")
    print("-" * 60)
    print()
    
    json_path = input("Enter path to parsed JSON file: ").strip()
    
    if not json_path:
        print("No file specified.")
        return interactive_menu()
    
    json_path = Path(json_path)
    if not json_path.exists():
        print(f"ERROR: File not found: {json_path}")
        return interactive_menu()
    
    checksheet_path = input("Enter path to checksheet template (or press Enter for default): ").strip()
    if checksheet_path:
        checksheet_path = Path(checksheet_path)
    else:
        checksheet_path = script_dir / 'tax_verification_checksheet.xlsx'
    
    column = input("Fill which column? (cch/source) [cch]: ").strip().lower() or 'cch'
    
    do_fill(json_path, checksheet_path, column)
    
    print()
    input("Press Enter to continue...")
    interactive_menu()


def do_full_interactive():
    """Interactive full workflow."""
    print()
    print("-" * 60)
    print("FULL WORKFLOW (Parse + Fill)")
    print("-" * 60)
    print()
    
    file_path = input("Enter path to Input Listing file (PDF or text): ").strip()
    
    if not file_path:
        print("No file specified.")
        return interactive_menu()
    
    file_path = Path(file_path)
    if not file_path.exists():
        print(f"ERROR: File not found: {file_path}")
        return interactive_menu()
    
    checksheet_path = input("Enter path to checksheet template (or press Enter for default): ").strip()
    if checksheet_path:
        checksheet_path = Path(checksheet_path)
    else:
        checksheet_path = script_dir / 'tax_verification_checksheet.xlsx'
    
    do_full(file_path, checksheet_path)
    
    print()
    input("Press Enter to continue...")
    interactive_menu()


def do_parse(file_path: Path, output_path: Path = None):
    """Parse Input Listing file."""
    print(f"\nParsing: {file_path}")
    
    # Extract text
    if file_path.suffix.lower() == '.pdf':
        print("Extracting text from PDF...")
        text = extract_text_from_pdf(file_path)
    else:
        text = file_path.read_text()
    
    # Parse
    print("Parsing data...")
    parsed = parse_input_listing(text)
    parsed['summary'] = calculate_summary(parsed)
    parsed['forms'] = dict(parsed['forms'])
    
    # Output path
    if not output_path:
        output_path = file_path.with_suffix('.json')
        output_path = output_path.with_stem(file_path.stem + '_parsed')
    
    # Save
    with open(output_path, 'w') as f:
        json.dump(parsed, f, indent=2, default=str)
    
    print(f"\nSaved: {output_path}")
    
    # Print summary
    print_parse_summary(parsed)
    
    return output_path, parsed


def do_fill(json_path: Path, checksheet_path: Path, column: str = 'cch', output_path: Path = None):
    """Fill checksheet from parsed data."""
    print(f"\nLoading: {json_path}")
    
    with open(json_path) as f:
        parsed_data = json.load(f)
    
    if not checksheet_path.exists():
        print(f"ERROR: Checksheet not found: {checksheet_path}")
        return None
    
    print(f"Loading checksheet: {checksheet_path}")
    workbook = load_workbook(checksheet_path)
    
    print(f"Filling '{column}' column...")
    counts = populate_checksheet(workbook, parsed_data, column)
    
    # Output path
    if not output_path:
        output_path = json_path.with_suffix('.xlsx')
        output_path = output_path.with_stem(json_path.stem.replace('_parsed', '') + '_checksheet')
    
    workbook.save(output_path)
    
    print(f"\nSaved: {output_path}")
    
    # Print summary
    print("\nPopulated:")
    for form_type, count in counts.items():
        print(f"  {form_type}: {count} entries")
    
    return output_path


def do_full(file_path: Path, checksheet_path: Path, column: str = 'cch'):
    """Full workflow: parse + fill."""
    # Parse
    json_path, parsed_data = do_parse(file_path)
    
    # Fill
    output_path = do_fill(json_path, checksheet_path, column)
    
    return output_path


def print_parse_summary(parsed):
    """Print parsing summary."""
    print("\n" + "=" * 50)
    print("PARSED DATA SUMMARY")
    print("=" * 50)
    
    for form_type, entries in parsed['forms'].items():
        form_name = CCH_FIELD_MAP.get(form_type, {}).get('name', form_type)
        print(f"  {form_name}: {len(entries)}")
    
    print("\n" + "-" * 50)
    summary = parsed['summary']
    
    if summary['income'].get('total_wages'):
        print(f"  Total Wages:      ${summary['income']['total_wages']:>12,}")
    if summary['income'].get('total_interest'):
        print(f"  Total Interest:   ${summary['income']['total_interest']:>12,}")
    if summary['income'].get('total_ordinary_dividends'):
        print(f"  Total Dividends:  ${summary['income']['total_ordinary_dividends']:>12,}")
    if summary['schedules'].get('schedule_c_gross'):
        print(f"  Sch C Gross:      ${summary['schedules']['schedule_c_gross']:>12,}")
    if summary['schedules'].get('schedule_e_rents'):
        print(f"  Sch E Rents:      ${summary['schedules']['schedule_e_rents']:>12,}")


def show_help():
    """Show help information."""
    print()
    print("=" * 60)
    print("HELP")
    print("=" * 60)
    print("""
WORKFLOW:

1. Export Input Listing from CCH ProSystem fx
   - In CCH: File > Print > Input Listing
   - Print to PDF or save as text file

2. Run this tool to parse the Input Listing
   - Creates JSON file with all extracted data

3. Run this tool to fill the checksheet
   - Populates the CCH column in the Excel checksheet

4. Manually enter Source column from source documents
   - Or use OCR tool (coming soon)

5. Review checksheet
   - Check for variances (should be 0)
   - Review Summary tab


COMMAND LINE USAGE:

  python run.py                     # Interactive menu
  python run.py parse FILE.pdf      # Parse Input Listing
  python run.py fill DATA.json      # Fill checksheet
  python run.py full FILE.pdf       # Parse + fill in one step


FILES:

  parse_input_listing.py    - Standalone parser script
  populate_checksheet.py    - Standalone population script
  tax_verification_checksheet.xlsx  - Blank checksheet template


REQUIREMENTS:

  pip install openpyxl

  For PDF support, install poppler-utils:
  - Ubuntu/Debian: sudo apt-get install poppler-utils
  - Mac: brew install poppler
  - Windows: Download from github.com/oschwartz10612/poppler-windows
""")
    input("\nPress Enter to continue...")
    interactive_menu()


def main():
    parser = argparse.ArgumentParser(
        description='Tax Verification Tool',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Commands')
    
    # Parse command
    parse_parser = subparsers.add_parser('parse', help='Parse CCH Input Listing')
    parse_parser.add_argument('file', help='Input Listing file (PDF or text)')
    parse_parser.add_argument('--output', '-o', help='Output JSON file')
    
    # Fill command
    fill_parser = subparsers.add_parser('fill', help='Fill checksheet from parsed data')
    fill_parser.add_argument('json_file', help='Parsed data JSON file')
    fill_parser.add_argument('--checksheet', '-c', help='Checksheet template')
    fill_parser.add_argument('--output', '-o', help='Output checksheet file')
    fill_parser.add_argument('--column', choices=['cch', 'source'], default='cch')
    
    # Full command
    full_parser = subparsers.add_parser('full', help='Full workflow (parse + fill)')
    full_parser.add_argument('file', help='Input Listing file (PDF or text)')
    full_parser.add_argument('--checksheet', '-c', help='Checksheet template')
    full_parser.add_argument('--column', choices=['cch', 'source'], default='cch')
    
    args = parser.parse_args()
    
    if args.command == 'parse':
        file_path = Path(args.file)
        output_path = Path(args.output) if args.output else None
        do_parse(file_path, output_path)
    
    elif args.command == 'fill':
        json_path = Path(args.json_file)
        checksheet_path = Path(args.checksheet) if args.checksheet else script_dir / 'tax_verification_checksheet.xlsx'
        output_path = Path(args.output) if args.output else None
        do_fill(json_path, checksheet_path, args.column, output_path)
    
    elif args.command == 'full':
        file_path = Path(args.file)
        checksheet_path = Path(args.checksheet) if args.checksheet else script_dir / 'tax_verification_checksheet.xlsx'
        do_full(file_path, checksheet_path, args.column)
    
    else:
        # No command - run interactive menu
        interactive_menu()


if __name__ == '__main__':
    main()
