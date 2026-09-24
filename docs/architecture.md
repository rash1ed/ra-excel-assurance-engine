# Architecture

## Flow

1. CLI receives an `.xlsx` path and optional required-column JSON.
2. `xlsx.py` opens the workbook as an OOXML ZIP package.
3. Workbook metadata resolves worksheet names, states and XML targets.
4. Worksheet XML is parsed into lightweight cell records.
5. `auditor.py` runs five deterministic QA checks.
6. Results are emitted as text and JSON reports.
7. v0.1 creates a byte-identical review copy and never mutates business data automatically.

## Design choices

- No spreadsheet application is required.
- No third-party Python dependency is required at runtime.
- Checks are deterministic and covered by generated fixtures.
- Remediation is intentionally separated from detection.
- Validation evidence is reproducible with `python scripts/run_validation.py`.

## v0.1 trust boundary

The engine detects structural and stored-value risks. It does not calculate formulas, infer business meaning, or modify source workbooks.
