---
name: document-renamer
description: Rename and organize expense, invoice, and bank statement PDFs using AI extraction. Expenses are renamed to VendorName_Date.pdf and moved to monthly folders. Incoming invoices (e.g., TalentHawk) are matched by amount against Excel and moved to folders based on payment date. Bank statements are renamed to BankName_StartDate_EndDate.pdf. Optionally matches against Excel and marks as uploaded. Use when organizing financial documents.
---

# Document Renamer

Processes PDFs based on document type using AI extraction:

- **Expenses/Receipts** → `VendorName_YYYY-MM-DD.pdf` → moved to monthly folder based on expense date
- **Incoming Invoices** → Keep original filename → matched by amount → moved to monthly folder based on payment date
- **Bank Statements** → `BankName_YYYY-MM-DD_YYYY-MM-DD.pdf` (date range)
- **SprintPoint Invoices** → Skipped (outgoing invoices)

## Usage

### Basic (rename only)

```bash
python3 scripts/rename_expenses.py /path/to/folder
```

### With Excel matching (required for invoices)

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
- **Date** - Transaction date (used as payment date for invoices)
- **Description** - Transaction description (used for expense vendor matching)
- **Amount** - Transaction amount (positive for credits/invoices, negative for expenses)
- **Uploaded** - Will be set to "Yes" when matched (created if missing)

### Matching Logic

**Expenses:** Vendor name in Description + same month/year

**Invoices:** Exact amount match on positive (credit) entries

## Example Output

```
Processing: V02394802112_2025-10-04.pdf
  Type: Expense
  Vendor: EE
  Date: 2025-10-04
  New name: EE_2025-10-04.pdf
  Moved to: 10 October/EE_2025-10-04.pdf

Processing: Invoice INV-0107.pdf
  Type: Incoming Invoice
  Vendor: TalentHawk
  Invoice: INV-0107
  Amount: 21841.25
  Matched to: TALENTHAWK LIMITED... (2025-10-13)
  Excel updated: Uploaded = Yes
  Moved to: 10 October/Invoice INV-0107.pdf

Processing: 20251003_91673173.pdf
  Type: Bank Statement
  Bank: HSBC
  Period: 2025-09-04 to 2025-10-03
  New name: HSBC_2025-09-04_2025-10-03.pdf

==================================================
Summary: 3 successful, 0 failed
```

## Supported Document Types

| Type | Rename? | Folder Based On | Example |
|------|---------|-----------------|---------|
| Expense | Yes → `Vendor_Date.pdf` | Expense date | `10 October/EE_2025-10-04.pdf` |
| Incoming Invoice | No (keep original) | Payment date from Excel | `10 October/Invoice INV-0107.pdf` |
| Bank Statement | Yes → `Bank_Start_End.pdf` | (stays in place) | `HSBC_2025-09-04_2025-10-03.pdf` |
| SprintPoint Invoice | Skipped | - | - |
