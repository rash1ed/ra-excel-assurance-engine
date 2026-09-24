from __future__ import annotations

from argparse import ArgumentParser
from pathlib import Path

from .auditor import (
    AuditConfig,
    audit_workbook,
    create_safe_copy,
    write_report_json,
    write_report_text,
)


def build_parser() -> ArgumentParser:
    parser = ArgumentParser(
        prog="ra-excel-audit",
        description="Audit .xlsx workbooks for five high-value QA risks.",
    )
    parser.add_argument("workbook", help="Path to the .xlsx workbook")
    parser.add_argument("--config", help="JSON config with required columns")
    parser.add_argument("--report", help="Text report path")
    parser.add_argument("--json-report", help="JSON report path")
    parser.add_argument("--corrected", help="Safe review-copy path")
    return parser
def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    source = Path(args.workbook)
    output_dir = source.parent / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    report_path = Path(args.report) if args.report else output_dir / "report.txt"
    json_path = (
        Path(args.json_report)
        if args.json_report
        else output_dir / "report.json"
    )
    corrected_path = (
        Path(args.corrected)
        if args.corrected
        else output_dir / "corrected.xlsx"
    )

    config = AuditConfig.from_json(args.config)
    result = audit_workbook(source, config)
    write_report_text(result, report_path)
    write_report_json(result, json_path)
    digest = create_safe_copy(source, corrected_path)
    print("RA Excel Assurance Engine v0.1")
    print(f"Workbook: {source}")
    print(f"Issues: {len(result.issues)}")
    for category, count in sorted(result.counts.items()):
        print(f"  {category}: {count}")
    print(f"Report: {report_path}")
    print(f"JSON: {json_path}")
    print(f"Corrected review copy: {corrected_path}")
    print(f"Safe-copy SHA256: {digest}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
