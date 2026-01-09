# Expense Renamer

A Python tool that uses AI (Claude) to process, rename, and organize financial PDF documents.

## Project Overview

This tool processes PDF documents and automatically:
- **Expenses/Receipts**: Renames to `VendorName_YYYY-MM-DD.pdf` and moves to monthly folders
- **Bank Statements**: Renames to `BankName_YYYY-MM-DD_YYYY-MM-DD.pdf` (date range)
- **Incoming Invoices**: Matches by amount against Excel, moves to folder based on payment date
- **SprintPoint Invoices**: Skipped (outgoing invoices)

## Project Structure

```
expense-renamer/
├── scripts/
│   └── rename_expenses.py   # Main script (single file application)
├── SKILL.md                 # Skill definition for Claude Code
└── CLAUDE.md                # This file
```

## Key Commands

```bash
# Basic usage (rename only)
python3 scripts/rename_expenses.py /path/to/pdfs

# With Excel matching (marks expenses as uploaded)
python3 scripts/rename_expenses.py /path/to/pdfs --excel /path/to/expenses.xlsx

# Preview changes without modifying files
python3 scripts/rename_expenses.py /path/to/pdfs --dry-run
```

## Dependencies

```bash
pip install pdfplumber anthropic openpyxl pandas
export ANTHROPIC_API_KEY="your-key-here"
```

## Architecture Notes

### Document Classification
The AI classifies documents into 4 types:
1. `expense` - Receipts/bills SprintPoint pays (small recurring amounts)
2. `incoming_invoice` - Contractor invoices where SprintPoint receives payment (large amounts, TalentHawk)
3. `bank_statement` - Bank account statements with date ranges
4. `sprintpoint_invoice` - Outgoing invoices issued by SprintPoint (skipped)

### Excel Matching Strategies
1. **Vendor matching**: Vendor name in Description + same month/year
2. **First word matching**: First significant word (4+ chars) for partial matches
3. **URL/domain matching**: Vendor name embedded in URLs (5+ chars)
4. **Amazon combined totals**: Groups Amazon receipts by date, matches combined total by amount

### Auto-marked Entries
Entries that don't need receipts are auto-marked as `-` (not applicable):
- Internal transfers (BARTLEY, DIRECTOR loans)
- Tax payments (HMRC, PAYE, TAX)
- Salary/wages/pension
- Non-sterling transaction fees

## Code Conventions

- Single-file Python application
- Uses Claude Sonnet for AI extraction (claude-sonnet-4-20250514)
- Truncates PDF text to 5000 chars for API calls
- Files are moved to folders like `10 October/` based on date
- Duplicate handling via `_1`, `_2` suffixes
