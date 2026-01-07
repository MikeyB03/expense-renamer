---
name: document-renamer
description: Rename expense and bank statement PDFs using AI extraction. Expenses are renamed to VendorName_Date.pdf and moved to monthly folders (e.g., "10 October"). Bank statements are renamed to BankName_StartDate_EndDate.pdf. Optionally matches expenses against an Excel file and marks them as uploaded. Automatically skips SprintPoint invoices (outgoing). Use when organizing financial documents, renaming expenses/receipts, or processing bank statements.
---

# Document Renamer

Renames PDFs based on document type using AI extraction:

- **Expenses/Receipts** → `VendorName_YYYY-MM-DD.pdf` → moved to monthly folder (e.g., `10 October/`)
- **Bank Statements** → `BankName_YYYY-MM-DD_YYYY-MM-DD.pdf` (date range)
- **SprintPoint Invoices** → Skipped (outgoing invoices)

Expenses are automatically moved to monthly subfolders based on their date. Optionally matches expenses against an Excel file and marks them as "Yes" in the Uploaded column.

## Usage

### Basic (rename only)

```bash
python3 scripts/rename_expenses.py /path/to/folder
```

### With Excel matching

```bash
python3 scripts/rename_expenses.py /path/to/folder --excel /path/to/expenses.xlsx
```

### Dry Run (Preview)

```bash
python3 scripts/rename_expenses.py /path/to/folder --excel /path/to/expenses.xlsx --dry-run
```

## Requirements

```bash
pip install pdfplumber anthropic openpyxl pandas
export ANTHROPIC_API_KEY="your-key-here"
```

## Excel File Format

The Excel file must have these columns:
- **Date** - Transaction date
- **Description** - Transaction description (used for vendor matching)
- **Uploaded** - Will be set to "Yes" when matched (created if missing)

Matching logic:
- Vendor name must appear in the Description (word boundary match)
- Transaction must be in the same month/year as the expense

## Example Output

```
Processing: V02394802112_2025-10-04.pdf
  Type: Expense
  Vendor: EE
  Date: 2025-10-04
  New name: EE_2025-10-04.pdf
  Moved to: 10 October/EE_2025-10-04.pdf

Processing: 20251003_91673173.pdf
  Type: Bank Statement
  Bank: HSBC
  Period: 2025-09-04 to 2025-10-03
  New name: HSBC_2025-09-04_2025-10-03.pdf

==================================================
Summary: 2 successful, 0 failed

==================================================
Matching expenses against: Expenses.xlsx

Matches found (1):
  EE LIMITED... (2025-10-13) <- EE (2025-10-04)

Excel file updated: 1 rows marked as 'Yes' in Uploaded column
Matched 1 expense(s) in Excel file
```

## Supported Document Types

| Type | Naming Format | Folder | Example |
|------|---------------|--------|---------|
| Expense | `Vendor_Date.pdf` | `MM MonthName/` | `10 October/EE_2025-10-04.pdf` |
| Bank Statement | `Bank_StartDate_EndDate.pdf` | (stays in place) | `HSBC_2025-09-04_2025-10-03.pdf` |
| SprintPoint Invoice | Skipped | - | - |
