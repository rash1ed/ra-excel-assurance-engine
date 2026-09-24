# RA Excel Assurance Engine

A lightweight, dependency-free quality assurance tool for `.xlsx` workbooks used in operations, PMO and business reporting.

## What v0.1 checks

1. Excel error values such as `#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, `#NAME?`, `#NULL!` and `#NUM!`.
2. Exact duplicate data rows.
3. Blank values in required columns defined in a JSON config.
4. Hidden or very-hidden worksheets.
5. Formulas that reference worksheets that do not exist in the workbook.

The engine reads the workbook's OOXML structure directly using Python's standard library (`zipfile` + `xml.etree.ElementTree`). It does not require Excel, LibreOffice, openpyxl or pandas.

## Why this project exists

Operational spreadsheets often become business-critical before anyone adds formal QA. This tool creates a repeatable review step before a workbook is used for reporting or handoff.

## Verified v0.1 status

The included validation suite currently passes **5/5 tests**. The unit suite is written with Python's built-in `unittest` framework and is executed by `scripts/run_validation.py`.

Three generated sample workbooks are exercised end-to-end:

- `issue_rich.xlsx`: 6 findings across all five audit categories.
- `clean.xlsx`: 0 findings.
- `quoted_refs.xlsx`: 1 expected required-column finding while a quoted sheet reference remains valid.

The full validation transcript is stored in [`docs/validation.txt`](docs/validation.txt).

CI reruns the same validation suite on every push and pull request via [`.github/workflows/ci.yml`](.github/workflows/ci.yml).

## Run it

    python -m ra_excel_assurance samples/issue_rich.xlsx --config samples/config.json

Default outputs are written beside the source workbook under `output/`:

- `report.txt`
- `report.json`
- `review_copy.xlsx`

## Important v0.1 safety behavior

The file named `review_copy.xlsx` is intentionally a **byte-identical review copy** in v0.1.

The engine does not silently delete duplicate rows, fill missing values, unhide sheets or rewrite formulas. Those actions require business context and could corrupt legitimate data.

Automatic remediation is reserved for a later version with explicit opt-in rules.

## Required-column config

Example:

    {
      "required_columns": {
        "Data": ["ID", "Name"]
      }
    }

## Project structure

- `ra_excel_assurance/` — OOXML reader, audit logic and CLI.
- `tests/` — unit and end-to-end tests.
- `scripts/generate_samples.py` — builds deterministic sample workbooks.
- `scripts/run_validation.py` — runs the full proof suite and writes the transcript.
- `samples/` — reproducible workbook fixtures.
- `docs/validation.txt` — latest verified execution evidence.

## Current limitations

v0.1 does **not** calculate Excel formulas. It can identify stored Excel error values and missing-sheet references, but it cannot determine whether a mathematically valid formula produces the intended business result.

It also does not infer control totals automatically; those require a declared business rule.

## Validation command

    python scripts/run_validation.py

A successful run must return exit code 0 for the unit suite and each sample workbook.

## License

MIT
