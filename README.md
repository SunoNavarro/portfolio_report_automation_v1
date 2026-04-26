# Automated Sales Report Generator

A practical Python automation project that turns raw sales data into a clean Excel business report with KPIs, summary tables and charts.

This project is designed as a freelance portfolio example for small businesses that still create reports manually in Excel.

## Business problem

Many teams export sales data from CRMs, ERPs or spreadsheets and spend hours every week cleaning data, calculating KPIs and creating management reports.

This script automates that process:

- Reads raw CSV sales data
- Cleans and validates the data
- Calculates revenue KPIs
- Creates summary tables by month, region, product and client
- Generates an Excel report with formatted sheets and charts

## Example use case

A small sales team sends a weekly CSV export. Instead of manually preparing the report, they can run one command and get a ready-to-share Excel file.

Potential value: saves 2–5 hours per week and reduces manual reporting errors.

## Tech stack

- Python
- pandas
- openpyxl
- Excel reporting automation

## Project structure

```text
portfolio_report_automation_v1/
├── data/
│   └── sales_sample.csv
├── output/
│   └── sales_report.xlsx
├── src/
│   ├── __init__.py
│   ├── generate_report.py
│   └── visual_script.py
├── tests/
│   └── test_generate_report.py
├── requirements.txt
└── README.md
```

## How to run

Create a virtual environment:

```bash
python -m venv .venv
source .venv/bin/activate  # macOS/Linux
.venv\Scripts\activate     # Windows
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Generate the report:

```bash
python src/generate_report.py --input data/sales_sample.csv --output output/sales_report.xlsx
```

Optional CSV text encoding (defaults to `utf-8`; useful on Windows with Excel exports):

```bash
python src/generate_report.py --input data/sales_sample.csv --output output/sales_report.xlsx --encoding cp1252
```

## Branch `visual_script` (desktop GUI)

Work on the lightweight **tkinter** front-end lives on the **`visual_script`** branch. It uses only the Python standard library for UI (no extra GUI packages) and reuses `generate_report.py` unchanged.

**What it does**

- Opens a native file dialog so you can pick a sales **CSV**.
- Writes the Excel workbook **in the same folder as the CSV**, named `{original_csv_stem}_report.xlsx` (for example `sales_sample.csv` → `sales_sample_report.xlsx`).
- Shows a confirmation dialog with the full output path, or an error dialog if generation fails (for example the output file is open in Excel).

**How to use it**

Check out the branch, install dependencies as usual, then run:

```bash
git fetch origin
git checkout visual_script
pip install -r requirements.txt
python src/visual_script.py
```

If you work only locally and the branch already exists:

```bash
git checkout visual_script
python src/visual_script.py
```

Cancelling the file dialog closes the app without generating a report.

## Testing

Install dependencies (includes `pytest`), then from the project root:

```bash
python -m pytest tests/ -v
```

## Large exports and the “Clean Data” sheet

Summary sheets (KPIs, monthly revenue, etc.) stay small, so the script applies full formatting: header styling, borders on every cell, and column widths based on all rows.

The **Clean Data** sheet can contain tens of thousands of transaction rows. Formatting every cell that way is slow and memory-heavy in openpyxl. For **Clean Data** only, the script uses a **lightweight** layout:

- Same header row styling and frozen panes under the header
- **No** per-cell borders on the data body (much faster)
- Column widths are estimated from the **header plus the first 1,000 data rows** instead of scanning the entire column

Business logic and values are unchanged; only presentation differs on that sheet for large dumps.

## Output file locked (Excel)

If the target workbook is **open in Excel** (or another program), saving may fail with a **permission** error. Close the file or pick a different `--output` path. When run as a script, a clear message is printed to stderr and the process exits with code **1**.

## Input format

The input CSV must contain these columns:

```text
date, client, region, product, units, unit_price, status
```

Allowed status values:

```text
paid, pending, cancelled
```

## Output

The script creates an Excel report with these sheets:

- KPIs
- Monthly Revenue
- Revenue by Region
- Revenue by Product
- Top Clients
- Clean Data

It also adds charts for monthly revenue and product revenue.

## Freelance service angle

This project demonstrates how I can help businesses automate repetitive reporting tasks with Python.

Examples of possible client projects:

- Automated weekly sales reports
- Excel/CSV data cleaning scripts
- CRM export processing
- KPI report generation
- Invoice or order summary reports

## Possible improvements

- Send the report automatically by email
- Generate a PDF version
- Connect directly to Google Sheets or a database
- Schedule the script to run daily or weekly
- Add a simple web interface
