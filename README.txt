# LP Reporting Automation

Automated quarterly Capital Account Statement (CAS) reporting
for Limited Partners in a Private Equity fund.

## The problem it solves
Each quarter, fund accountants manually copy NAV data from Excel
into Word letters, convert them to PDFs, and email each LP
individually. For 20+ LPs this takes hours and risks human error.

## What this script does
1. Reads all LP data from a CAS Excel file (one row per investor)
2. Fills a personalised Word template for each LP
3. Converts each Word doc to a named PDF  (e.g. ABC_Holdings_Q1_2026.pdf)
4. Creates an Outlook draft per LP with the PDF attached
   and their email address pre-filled — ready to review and send

## Tech stack
- Python 3.8+
- openpyxl — reads Excel CAS data
- python-docx — fills Word template via XML manipulation
- pywin32 — Word COM automation for PDF conversion + Outlook drafts

## Key challenges solved
- Word COM not releasing documents between LPs (fixed with
  CoInitialize / CoUninitialize per document + gc.collect())
- Duplicate investor names causing file overwrites (fixed with
  LP Number prefix on every output filename)
- Date fields showing timestamps (stripped to date-only)
- Negative values displayed in accounting format: ($2,313)
- Zero-activity rows showing "-" instead of "$0"

## How to run
pip install openpyxl python-docx pywin32
python create_word_template.py   # run once to build the template
python lp_report_automation.py   # run each quarter

## Background
Built by a fund accountant learning Python to automate
repetitive quarterly workflows in Private Equity operations.
Developed with AI-Assisted pair programming (Claude by Anthropic).
