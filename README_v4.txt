## Y-Axis LP Reporting Automation — v4

Generates per-LP Capital Account Statement PDFs from a single Excel CAS file,
organized automatically into per-Entity folders.

**Pipeline:** Excel CAS → Word template (filled per LP) → PDF (grouped by Entity)

> Outlook draft creation has been removed from this version — the pipeline
> now ends at PDF generation.

### Requirements

```bash
pip install openpyxl python-docx pywin32
```

### Run

```bash
python lp_report_automation.py
```

### Configuration

Set these at the top of the script:

| Variable        | Purpose                                        |
|-----------------|-------------------------------------------------|
| `CAS_FILE`      | Path to the source Excel CAS workbook            |
| `CAS_SHEET`     | Sheet name to read (e.g. `"Summary"`)            |
| `WORD_TEMPLATE` | Path to the `{{TOKEN}}`-based Word template      |
| `OUTPUT_FOLDER` | Root folder where entity subfolders are created  |
| `QUARTER_LABEL` | Reporting period label (e.g. `"Q1 2026"`)        |

### How output is organized

PDFs are grouped into one subfolder per unique **Entity** name, with each
file named:

```
[Full Entity Name] CAS - MM.DD.YYYY - [Investor Name].pdf
```

- The entity's trailing legal suffix (`LP`, `LLC`, `LLP`, `Inc`, `Ltd`, etc.)
  is stripped from the *filename* only — the folder name keeps the full
  entity name as-is.
- If two rows share the same entity + date + investor name, the script
  appends `(2)`, `(3)`, etc. instead of silently overwriting the earlier PDF.

Example output tree:

```
Output_PDFs/
├── Ram Group Partners, LP/
│   ├── Ram Group Partners CAS - 03.31.2026 - ABC Holdings, LLC.pdf
│   └── Ram Group Partners CAS - 03.31.2026 - ABC Holdings, LLC (2).pdf
├── Ajay Group Partners, LP/
│   └── Ajay Group Partners CAS - 03.31.2026 - PQR Holdings, LLC.pdf
└── Delhi Group Partners/
    └── Delhi Group Partners CAS - 03.31.2026 - XYZ.pdf
```

### Internal uniqueness (Number column)

Column A (`Number`) in the CAS Excel sheet is the unique key used
internally for temp `.docx` naming and PDF-map tracking — it's independent
of the final PDF filename described above.

### Word → PDF reliability

PDF conversion uses `win32com.client.DispatchEx()` (not `Dispatch()`) to
guarantee a fresh Word process per LP rather than risking attachment to a
still-closing instance from the previous one, and actively waits for that
process to fully exit before moving to the next LP — eliminating the
intermittent timing/lag issue seen with rapid consecutive Word automation
calls.