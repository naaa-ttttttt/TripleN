# Adashe Ledger Automation

A Python script that automates the generation of a monthly Adashe/Ajo savings ledger from customer contribution data, producing a structured Excel report.

## Overview

This project replaces manual data entry for Adashe ("Triple N plus one") savings tracking. Customer details are stored in a CSV file, and the script calculates total collections, company earnings, and customer payouts for a given month, exporting the results to an Excel workbook.

## Requirements

- Python 3
- pandas
- openpyxl

Install dependencies:

\`\`\`bash
pip install pandas openpyxl
\`\`\`

## Project Structure

\`\`\`
ADASHE/
├── app.py              # Main script
├── customers.csv       # Customer data (name, daily_rate)
├── Adashe_ledger.xlsx  # Generated output
└── README.md
\`\`\`

## Input Format

Customer data is read from customers.csv, which must contain the following columns:

| Column      | Description                          |
|-------------|---------------------------------------|
| name        | Customer's full name                  |
| daily_rate  | Fixed daily contribution amount (NGN) |

Example:

\`\`\`csv
name,daily_rate
Nathaniel,5000
Noel,3000
\`\`\`

## Usage

Run the script from the project directory:

\`\`\`bash
python3 app.py
\`\`\`

This will:

1. Read customer data from customers.csv
2. Calculate total days in the specified month and year
3. Compute total collected, company earnings, and customer payout per customer
4. Write the results to Adashe_ledger.xlsx, on a sheet named Adashe_<Month>_<Year>

## Output Columns

| Column              | Description                                  |
|---------------------|-----------------------------------------------|
| Customer Name       | Name of the contributor                       |
| Daily rate (NGN)    | Fixed daily contribution amount               |
| Total Days          | Number of days in the contribution month      |
| Total Collected (NGN) | Daily rate multiplied by total days         |
| Company's Earnings (NGN) | Amount retained as company earnings       |
| Customer Payout (NGN) | Amount due to the customer                  |

## Current Limitations

- Assumes every customer contributes for the full month at a fixed daily rate; does not currently account for missed or partial contribution days.
- Running the script overwrites the existing Adashe_ledger.xlsx file; there is no automatic archiving of previous months.
- Column headers in customers.csv must match name and daily_rate exactly.

## Planned Improvements

- Move from an assumed monthly total to a daily contribution log, allowing missed or variable payments to be tracked and flagged.
- Automatic output file naming per month to preserve historical records.
- Input validation for missing, duplicate, or malformed customer entries.

## License

Internal project — not yet licensed for external distribution.
