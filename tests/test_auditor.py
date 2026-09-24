from __future__ import annotations

import hashlib
import json
from pathlib import Path
import tempfile
import unittest

from ra_excel_assurance.auditor import (
    AuditConfig,
    audit_workbook,
    create_safe_copy,
)
from ra_excel_assurance.cli import main as cli_main
from tests.xlsx_factory import make_xlsx


class ExcelAssuranceTests(unittest.TestCase):
    def _issue_workbook(self, root: Path) -> Path:
        return make_xlsx(
            root / "issue_rich.xlsx",
            [
                {
                    "name": "Data",
                    "rows": [
                        ["ID", "Name", "Amount", "Calc"],
                        [1, "Alpha", 10, {"formula": "MissingSheet!A1"}],
                        [1, "Alpha", 10, {"formula": "MissingSheet!A1"}],
                        [2, "", 20, {"error": "#REF!"}],
                    ],
                },
                {
                    "name": "HiddenData",
                    "state": "hidden",
                    "rows": [["Key"], ["x"]],
                },
            ],
        )

    def test_issue_rich_workbook_hits_all_five_checks(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            book = self._issue_workbook(root)
            config = AuditConfig(required_columns={"Data": ["ID", "Name"]})
            result = audit_workbook(book, config)
            categories = {issue.category for issue in result.issues}
            expected = {
                "ERROR_VALUE",
                "DUPLICATE_ROW",
                "BLANK_REQUIRED",
                "HIDDEN_SHEET",
                "MISSING_SHEET_REFERENCE",
            }
            self.assertTrue(expected.issubset(categories))
    def test_clean_workbook_has_no_issues(self):
        with tempfile.TemporaryDirectory() as tmp:
            book = make_xlsx(
                Path(tmp) / "clean.xlsx",
                [
                    {
                        "name": "Data",
                        "rows": [
                            ["ID", "Name", "Lookup"],
                            [1, "Alpha", {"formula": "Lookup!A1"}],
                            [2, "Beta", {"formula": "Lookup!A2"}],
                        ],
                    },
                    {
                        "name": "Lookup",
                        "rows": [["Value"], [100], [200]],
                    },
                ],
            )
            config = AuditConfig(required_columns={"Data": ["ID", "Name"]})
            result = audit_workbook(book, config)
            self.assertEqual(result.issues, [])

    def test_quoted_sheet_reference_is_valid(self):
        with tempfile.TemporaryDirectory() as tmp:
            book = make_xlsx(
                Path(tmp) / "quoted.xlsx",
                [
                    {
                        "name": "Data",
                        "rows": [
                            ["ID", "Calc"],
                            [1, {"formula": "'Valid Sheet'!A1"}],
                        ],
                    },
                    {
                        "name": "Valid Sheet",
                        "rows": [["Value"], [10]],
                    },
                ],
            )
            result = audit_workbook(
                book,
                AuditConfig(required_columns={"Data": ["ID", "Owner"]}),
            )
            categories = [issue.category for issue in result.issues]
            self.assertIn("BLANK_REQUIRED", categories)
            self.assertNotIn("MISSING_SHEET_REFERENCE", categories)

    def test_safe_copy_is_byte_identical(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = self._issue_workbook(root)
            target = root / "review_copy.xlsx"
            digest = create_safe_copy(source, target)
            self.assertEqual(source.read_bytes(), target.read_bytes())
            self.assertEqual(
                digest,
                hashlib.sha256(source.read_bytes()).hexdigest(),
            )

    def test_cli_creates_reports_and_review_copy(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = self._issue_workbook(root)
            config = root / "config.json"
            config.write_text(
                json.dumps({"required_columns": {"Data": ["ID", "Name"]}}),
                encoding="utf-8",
            )
            code = cli_main([str(source), "--config", str(config)])
            self.assertEqual(code, 0)
            self.assertTrue((root / "output" / "report.txt").exists())
            self.assertTrue((root / "output" / "report.json").exists())
            self.assertTrue((root / "output" / "review_copy.xlsx").exists())


if __name__ == "__main__":
    unittest.main()
