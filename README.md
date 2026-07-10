# Financial Record Checker

A Python command-line tool that automates the comparison of bank statement transactions against internal records to identify discrepancies and ensure accurate financial settlement.

## What It Does

This tool performs transaction reconciliation by:

1. **Loading Data**: Reads two CSV files - `bank_statement.csv` (actual bank transactions) and `internal_records.csv` (company's internal records)
2. **Matching Transactions**: Compares transactions based on amount and date, with a ±1 day tolerance to account for real-world posting delays
3. **Classifying Results**: Categorizes each transaction as:
   - **Matched** - Found in both files with matching amount/date
   - **Missing in Records** - Present in bank statement but not recorded internally
   - **Missing in Bank Statement** - Recorded internally but not received by the bank
4. **Generating Reports**: Creates a formatted Excel file with three sheets:
   - **Matched** - All successfully reconciled transactions
   - **Discrepancies** - Unmatched transactions (highlighted in red)
   - **Summary** - Key metrics and statistics

## Real-World Use Cases

Operations teams use this tool for:

- **Transaction Reconciliation**: Daily/monthly verification that all bank transactions are properly recorded in internal systems
- **Spotting Processing Errors**: Identifying missing entries, duplicate records, or data entry mistakes
- **Ensuring Accurate Settlement**: Confirming that expected payments have been received and recorded correctly
- **Audit Trail**: Maintaining documented evidence of reconciliation activities for compliance
- **Cash Flow Management**: Quickly identifying discrepancies that could impact cash position

## Installation

Install required dependencies:

```bash
pip install pandas openpyxl
```

## Usage

### Basic Usage (with sample data)

```bash
python reconcile.py
```

This uses the default files:
- Input: `bank_statement.csv`, `internal_records.csv`
- Output: `bank_reconciliation_report.xlsx`

### Custom Files

```bash
python reconcile.py --bank custom_bank.csv --records custom_records.csv --output custom_report.xlsx
```

### Command-Line Arguments

- `--bank`: Path to bank statement CSV file (default: `bank_statement.csv`)
- `--records`: Path to internal records CSV file (default: `internal_records.csv`)
- `--output`: Path for output Excel report (default: `bank_reconciliation_report.xlsx`)

## Input File Format

Both CSV files must contain these columns:

| Column | Description | Example |
|--------|-------------|---------|
| `transaction_id` | Unique identifier for the transaction | `BANK001`, `LEDGER001` |
| `date` | Transaction date (YYYY-MM-DD) | `2024-01-02` |
| `description` | Transaction description | `Customer Payment - Invoice #1001` |
| `amount` | Transaction amount (positive for inflows) | `1500.00` |

## Sample Data

The repository includes sample CSV files with 25-27 synthetic transactions, including intentional mismatches to demonstrate the tool's capabilities:

- **bank_statement.csv**: 25 bank transactions
- **internal_records.csv**: 27 ledger transactions (2 extra entries not in bank)

## Output

The tool generates:

1. **Console Summary**: Real-time statistics displayed in the terminal
2. **Excel Report**: Formatted workbook with:
   - Bold headers with blue background
   - Red highlighting for unmatched transactions
   - Auto-adjusted column widths
   - Summary metrics sheet

## Example Output

```
============================================================
Bank Transaction Reconciliation Tool
============================================================

Loading bank statement from: bank_statement.csv
Loading internal ledger from: internal_ledger.csv
Bank statement: 25 transactions
Internal ledger: 27 transactions

Matching transactions...

============================================================
RECONCILIATION SUMMARY
============================================================
Total Transactions Processed: 52
Matched Transactions: 25
Unmatched Transactions: 2
  - Missing in Ledger: 0
  - Missing in Bank Statement: 2
Total Value of Unmatched Transactions: $13,700.00
============================================================

Generating Excel report...
Excel report saved to: reconciliation_report.xlsx

Reconciliation complete!
```

## How It Works

1. **Date Tolerance**: The tool allows ±1 day matching to handle posting delays (e.g., a transaction recorded on Jan 15 in the ledger may appear as Jan 16 in the bank statement)
2. **Amount Matching**: Transactions must match exactly on amount
3. **One-to-One Matching**: Each transaction can only be matched once to prevent false positives
4. **Comprehensive Reporting**: All transactions are accounted for in the output, with clear categorization

## Error Handling

The tool validates:
- Required columns exist in both files
- Date formats are parseable
- Files are accessible

Errors are reported with clear messages to help diagnose issues quickly.
