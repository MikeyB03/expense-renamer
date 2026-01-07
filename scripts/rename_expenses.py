#!/usr/bin/env python3
"""
Document PDF Renamer

Processes and renames PDFs based on document type:
- Expenses/receipts → VendorName_YYYY-MM-DD.pdf
- Bank statements → BankName_YYYY-MM-DD_YYYY-MM-DD.pdf (date range)
- SprintPoint invoices → Skipped (outgoing invoices)

Optionally matches expenses against an Excel file and marks them as uploaded.

Usage:
    python rename_expenses.py /path/to/folder [--dry-run] [--excel /path/to/expenses.xlsx]

Options:
    --dry-run    Preview renames without actually renaming files
    --excel      Path to Excel file to match expenses against and mark as uploaded

Requirements:
    pip install pdfplumber anthropic openpyxl pandas

Environment:
    ANTHROPIC_API_KEY must be set
"""

import os
import re
import sys
import json
import shutil
import argparse
from pathlib import Path
from datetime import datetime
from typing import Optional, Tuple, Dict, Any, List

try:
    import pdfplumber
except ImportError:
    print("Error: pdfplumber is required. Install with: pip install pdfplumber")
    sys.exit(1)

try:
    import anthropic
except ImportError:
    print("Error: anthropic is required. Install with: pip install anthropic")
    sys.exit(1)


# Initialize Anthropic client
client = None


def get_client() -> anthropic.Anthropic:
    """Get or create the Anthropic client."""
    global client
    if client is None:
        api_key = os.environ.get("ANTHROPIC_API_KEY")
        if not api_key:
            print("Error: ANTHROPIC_API_KEY environment variable is not set")
            sys.exit(1)
        client = anthropic.Anthropic(api_key=api_key)
    return client


def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract all text from a PDF file."""
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        print(f"  Warning: Could not read {pdf_path}: {e}")
    return text


def extract_with_ai(text: str) -> Dict[str, Any]:
    """
    Use Claude to extract document information.
    Returns a dict with document type and relevant fields.
    """
    # Truncate text if too long (keep first ~4000 chars which should contain key info)
    if len(text) > 5000:
        text = text[:5000]

    prompt = f"""Analyze this document and extract information based on its type.

FIRST, determine the document type:
1. "bank_statement" - A bank statement showing account transactions over a period
2. "expense" - A receipt, bill, or invoice FROM another company (expense to be paid/already paid)
3. "sprintpoint_invoice" - An invoice issued BY SprintPoint Ltd (outgoing invoice)

THEN extract the relevant information:

For BANK STATEMENTS:
- bank_name: The bank's name (e.g., "HSBC", "Barclays", "NatWest")
- start_date: Statement period start date in YYYY-MM-DD format
- end_date: Statement period end date in YYYY-MM-DD format

For EXPENSES:
- vendor: The company that issued the document (use well-known brand names, not legal entities)
  Examples: Use "Uber" not "DÉCADA OUSADA LDA", "Amazon" not "Amazon EU S.à r.l."
- date: The document date in YYYY-MM-DD format

For SPRINTPOINT INVOICES:
- Just identify it as type "sprintpoint_invoice"

Return ONLY a JSON object:
{{
  "document_type": "bank_statement" | "expense" | "sprintpoint_invoice",
  // Include relevant fields based on type:
  // For bank_statement: "bank_name", "start_date", "end_date"
  // For expense: "vendor", "date"
}}

Document text:
{text}

JSON response:"""

    try:
        response = get_client().messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=300,
            messages=[
                {"role": "user", "content": prompt}
            ]
        )

        # Parse the response
        response_text = response.content[0].text.strip()

        # Try to extract JSON from the response
        # Handle cases where response might have markdown code blocks
        if "```" in response_text:
            json_match = re.search(r'```(?:json)?\s*(.*?)\s*```', response_text, re.DOTALL)
            if json_match:
                response_text = json_match.group(1)

        data = json.loads(response_text)
        return data

    except json.JSONDecodeError as e:
        print(f"  Warning: Could not parse AI response as JSON: {e}")
        return {"document_type": "unknown"}
    except Exception as e:
        print(f"  Warning: AI extraction failed: {e}")
        return {"document_type": "unknown"}


def sanitize_filename(name: str) -> str:
    """
    Sanitize a string to be safe for use as a filename.
    """
    # Replace problematic characters
    name = re.sub(r'[<>:"/\\|?*]', '', name)
    # Replace multiple spaces/underscores with single hyphen
    name = re.sub(r'[\s_]+', '-', name)
    # Remove leading/trailing whitespace and dots
    name = name.strip('. ')
    # Limit length
    if len(name) > 50:
        name = name[:50].rsplit('-', 1)[0]  # Try to break at word boundary
    return name or "Unknown"


def get_unique_filename(folder: Path, base_name: str, extension: str) -> str:
    """
    Get a unique filename by adding _1, _2, etc. if file exists.
    """
    filename = f"{base_name}{extension}"
    if not (folder / filename).exists():
        return filename

    counter = 1
    while True:
        filename = f"{base_name}_{counter}{extension}"
        if not (folder / filename).exists():
            return filename
        counter += 1
        if counter > 999:  # Safety limit
            raise ValueError(f"Too many duplicates for {base_name}")


def get_month_folder_name(date_str: str) -> str:
    """
    Convert a date string to a month folder name like '10 October'.

    Args:
        date_str: Date in YYYY-MM-DD format

    Returns:
        Folder name like '01 January', '10 October', '12 December'
    """
    try:
        date = datetime.strptime(date_str, "%Y-%m-%d")
        month_num = date.strftime("%m")  # Zero-padded month number
        month_name = date.strftime("%B")  # Full month name
        return f"{month_num} {month_name}"
    except ValueError:
        return None


def move_to_month_folder(file_path: Path, date_str: str, dry_run: bool = False) -> Optional[Path]:
    """
    Move a file to the appropriate monthly folder based on date.
    Creates the folder if it doesn't exist.

    Args:
        file_path: Path to the file to move
        date_str: Date in YYYY-MM-DD format
        dry_run: If True, don't actually move the file

    Returns:
        New path if moved successfully, None otherwise
    """
    month_folder = get_month_folder_name(date_str)
    if not month_folder:
        return None

    # Create the month folder path
    dest_folder = file_path.parent / month_folder
    dest_path = dest_folder / file_path.name

    if dry_run:
        print(f"  Would move to: {month_folder}/{file_path.name}")
        return dest_path

    # Create folder if it doesn't exist
    dest_folder.mkdir(exist_ok=True)

    # Handle if destination file already exists
    if dest_path.exists():
        # Get unique filename in destination folder
        base_name = file_path.stem
        extension = file_path.suffix
        unique_name = get_unique_filename(dest_folder, base_name, extension)
        dest_path = dest_folder / unique_name

    # Move the file
    try:
        shutil.move(str(file_path), str(dest_path))
        print(f"  Moved to: {month_folder}/{dest_path.name}")
        return dest_path
    except Exception as e:
        print(f"  Warning: Failed to move file: {e}")
        return None


def process_document(pdf_path: Path, dry_run: bool = False) -> Tuple[bool, str, Optional[Dict]]:
    """
    Process a single PDF and rename it based on document type.
    Returns (success, message, expense_info).
    expense_info contains vendor and date for expense documents (for Excel matching).
    """
    print(f"\nProcessing: {pdf_path.name}")

    # Extract text from PDF
    text = extract_text_from_pdf(str(pdf_path))
    if not text:
        return False, "Could not extract text from PDF", None

    # Use AI to extract document info
    doc_info = extract_with_ai(text)
    doc_type = doc_info.get("document_type", "unknown")

    expense_info = None

    # Handle based on document type
    if doc_type == "sprintpoint_invoice":
        print(f"  Type: SprintPoint Invoice (skipping)")
        return True, "Skipped (SprintPoint invoice)", None

    elif doc_type == "bank_statement":
        bank_name = doc_info.get("bank_name")
        start_date = doc_info.get("start_date")
        end_date = doc_info.get("end_date")

        if not bank_name:
            return False, "Could not identify bank name", None
        if not start_date or not end_date:
            return False, "Could not identify statement date range", None

        print(f"  Type: Bank Statement")
        print(f"  Bank: {bank_name}")
        print(f"  Period: {start_date} to {end_date}")

        # Build filename: BankName_StartDate_EndDate.pdf
        safe_bank = sanitize_filename(bank_name)
        base_name = f"{safe_bank}_{start_date}_{end_date}"

    elif doc_type == "expense":
        vendor = doc_info.get("vendor")
        date = doc_info.get("date")

        if not vendor:
            return False, "Could not identify vendor name", None
        if not date:
            return False, "Could not identify document date", None

        print(f"  Type: Expense")
        print(f"  Vendor: {vendor}")
        print(f"  Date: {date}")

        # Build filename: VendorName_Date.pdf
        safe_vendor = sanitize_filename(vendor)
        base_name = f"{safe_vendor}_{date}"

        # Store expense info for Excel matching
        expense_info = {"vendor": vendor, "date": date}

    else:
        return False, f"Unknown document type: {doc_type}", None

    # Get unique filename
    new_filename = get_unique_filename(pdf_path.parent, base_name, ".pdf")
    new_path = pdf_path.parent / new_filename

    if pdf_path.name == new_filename:
        # File already correctly named, but still move expenses to monthly folder
        if doc_type == "expense":
            date = doc_info.get("date")
            move_to_month_folder(pdf_path, date, dry_run)
        return True, "Already correctly named", expense_info

    print(f"  New name: {new_filename}")

    if dry_run:
        # Show what would happen
        if doc_type == "expense":
            date = doc_info.get("date")
            move_to_month_folder(pdf_path.parent / new_filename, date, dry_run)
        return True, f"Would rename to: {new_filename}", expense_info

    # Rename the file
    try:
        pdf_path.rename(new_path)

        # Move expenses to monthly folder
        if doc_type == "expense":
            date = doc_info.get("date")
            move_to_month_folder(new_path, date, dry_run)

        return True, f"Renamed to: {new_filename}", expense_info
    except Exception as e:
        return False, f"Failed to rename: {e}", None


def match_expenses_to_excel(excel_path: Path, expenses: List[Dict], dry_run: bool = False) -> int:
    """
    Match processed expenses against an Excel file and mark them as uploaded.
    Returns the number of matches found.
    """
    try:
        import pandas as pd
    except ImportError:
        print("Error: pandas is required for Excel matching. Install with: pip install pandas openpyxl")
        return 0

    if not excel_path.exists():
        print(f"Error: Excel file not found: {excel_path}")
        return 0

    print(f"\n{'='*50}")
    print(f"Matching expenses against: {excel_path.name}")

    # Read the Excel file
    df = pd.read_excel(excel_path)

    # Check if required columns exist
    if 'Description' not in df.columns:
        print("Error: Excel file must have a 'Description' column")
        return 0
    if 'Date' not in df.columns:
        print("Error: Excel file must have a 'Date' column")
        return 0
    if 'Uploaded' not in df.columns:
        print("Warning: 'Uploaded' column not found, creating it")
        df['Uploaded'] = None

    matches = []
    for exp in expenses:
        vendor = exp['vendor'].upper()
        try:
            exp_date = pd.Timestamp(exp['date'])
        except:
            continue

        for idx, row in df.iterrows():
            desc = str(row['Description']).upper()
            try:
                row_date = pd.Timestamp(row['Date'])
            except:
                continue

            # Match by vendor name in description and same month/year
            # Use word boundary matching to avoid partial matches (e.g., "EE" in "FEE")
            vendor_pattern = r'\b' + re.escape(vendor) + r'\b'
            if re.search(vendor_pattern, desc) and row_date.month == exp_date.month and row_date.year == exp_date.year:
                matches.append({
                    'excel_idx': idx,
                    'excel_date': row_date,
                    'excel_desc': row['Description'],
                    'expense_vendor': exp['vendor'],
                    'expense_date': exp['date']
                })

    if matches:
        print(f"\nMatches found ({len(matches)}):")
        for m in matches:
            print(f"  {m['excel_desc'][:40]}... ({m['excel_date'].strftime('%Y-%m-%d')}) <- {m['expense_vendor']} ({m['expense_date']})")

        if not dry_run:
            # Update the Uploaded column for matched rows
            for m in matches:
                df.at[m['excel_idx'], 'Uploaded'] = 'Yes'

            # Save the updated Excel file
            df.to_excel(excel_path, index=False)
            print(f"\nExcel file updated: {len(matches)} rows marked as 'Yes' in Uploaded column")
        else:
            print(f"\n(Dry run - Excel file not modified)")
    else:
        print("No matches found")

    return len(matches)


def main():
    parser = argparse.ArgumentParser(
        description="Rename PDFs based on document type (expenses, bank statements) using AI extraction"
    )
    parser.add_argument(
        "folder",
        type=str,
        help="Path to folder containing PDFs to process"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview renames without actually renaming files"
    )
    parser.add_argument(
        "--excel",
        type=str,
        help="Path to Excel file to match expenses against and mark as uploaded"
    )

    args = parser.parse_args()

    folder = Path(args.folder).expanduser().resolve()

    if not folder.exists():
        print(f"Error: Folder does not exist: {folder}")
        sys.exit(1)

    if not folder.is_dir():
        print(f"Error: Path is not a directory: {folder}")
        sys.exit(1)

    # Find all PDF files
    pdf_files = list(folder.glob("*.pdf")) + list(folder.glob("*.PDF"))

    if not pdf_files:
        print(f"No PDF files found in: {folder}")
        sys.exit(0)

    print(f"Found {len(pdf_files)} PDF file(s) in {folder}")
    print("Using AI (Claude) for extraction\n")
    if args.dry_run:
        print("DRY RUN - No files will be renamed\n")

    # Process each PDF
    success_count = 0
    fail_count = 0
    processed_expenses = []

    for pdf_path in sorted(pdf_files):
        success, message, expense_info = process_document(pdf_path, args.dry_run)
        if success:
            success_count += 1
            if expense_info:
                processed_expenses.append(expense_info)
        else:
            fail_count += 1
            print(f"  FAILED: {message}")

    # Summary
    print(f"\n{'='*50}")
    print(f"Summary: {success_count} successful, {fail_count} failed")
    if args.dry_run:
        print("(Dry run - no files were actually renamed)")

    # Match expenses against Excel if provided
    if args.excel and processed_expenses:
        excel_path = Path(args.excel).expanduser().resolve()
        match_count = match_expenses_to_excel(excel_path, processed_expenses, args.dry_run)
        print(f"Matched {match_count} expense(s) in Excel file")
    elif args.excel and not processed_expenses:
        print("\nNo expenses to match against Excel file")


if __name__ == "__main__":
    main()
