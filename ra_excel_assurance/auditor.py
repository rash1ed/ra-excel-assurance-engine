from __future__ import annotations

from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path
import hashlib
import json
import re
import shutil

from .xlsx import Cell, XlsxBook

ERROR_MARKERS = {
    "#REF!",
    "#DIV/0!",
    "#VALUE!",
    "#N/A",
    "#NAME?",
    "#NULL!",
    "#NUM!",
}
SHEET_REF_RE = re.compile(
    r"(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_. ]*))!"
)


def _norm(value: object | None) -> str:
    if value is None:
        return ""
    return str(value).strip()
def _formula_sheet_refs(formula: str) -> set[str]:
    cleaned = re.sub(r"\[[^\]]+\][^!]+!", "", formula)
    refs: set[str] = set()
    for match in SHEET_REF_RE.finditer(cleaned):
        name = match.group(1) or match.group(2) or ""
        name = name.replace("''", "'").strip()
        if name:
            refs.add(name)
    return refs


@dataclass(frozen=True)
class Issue:
    category: str
    severity: str
    sheet: str
    location: str
    problem: str
    action: str

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class AuditConfig:
    required_columns: dict[str, list[str]]

    @classmethod
    def from_json(cls, path: str | Path | None) -> "AuditConfig":
        if path is None:
            return cls(required_columns={})
        payload = json.loads(Path(path).read_text(encoding="utf-8"))
        raw = payload.get("required_columns", {})
        return cls(
            required_columns={
                str(sheet): [str(col) for col in columns]
                for sheet, columns in raw.items()
            }
        )


@dataclass
class AuditResult:
    workbook: str
    issues: list[Issue]

    @property
    def counts(self) -> dict[str, int]:
        return dict(Counter(issue.category for issue in self.issues))

    def to_dict(self) -> dict:
        return {
            "workbook": self.workbook,
            "issue_count": len(self.issues),
            "counts": self.counts,
            "issues": [issue.to_dict() for issue in self.issues],
        }


def _row_key(row: dict[int, Cell], width: int) -> tuple[str, ...]:
    values: list[str] = []
    for col in range(1, width + 1):
        cell = row.get(col)
        if cell is None:
            values.append("")
        elif cell.formula:
            values.append(cell.formula.strip())
        else:
            values.append(_norm(cell.value))
    return tuple(values)


def _audit_duplicates(sheet_name: str, rows: dict[int, dict[int, Cell]]) -> list[Issue]:
    issues: list[Issue] = []
    width = max((max(row.keys(), default=0) for row in rows.values()), default=0)
    seen: dict[tuple[str, ...], int] = {}
    for row_num in sorted(n for n in rows if n >= 2):
        key = _row_key(rows[row_num], width)
        if not any(key):
            continue
        if key in seen:
            issues.append(
                Issue(
                    category="DUPLICATE_ROW",
                    severity="MEDIUM",
                    sheet=sheet_name,
                    location=f"Row {row_num}",
                    problem=f"Exact duplicate of row {seen[key]}",
                    action="Review both rows before removing either record.",
                )
            )
        else:
            seen[key] = row_num
    return issues
def _audit_required(
    sheet_name: str,
    rows: dict[int, dict[int, Cell]],
    required: list[str],
) -> list[Issue]:
    if not required:
        return []
    issues: list[Issue] = []
    header = rows.get(1, {})
    header_map = {
        _norm(cell.value).casefold(): col
        for col, cell in header.items()
        if _norm(cell.value)
    }
    data_rows = [n for n in sorted(rows) if n >= 2 and rows[n]]

    for column_name in required:
        col = header_map.get(column_name.strip().casefold())
        if col is None:
            issues.append(
                Issue(
                    category="BLANK_REQUIRED",
                    severity="HIGH",
                    sheet=sheet_name,
                    location="Row 1",
                    problem=f"Required column not found: {column_name}",
                    action="Add the required column or update the audit config.",
                )
            )
            continue
        for row_num in data_rows:
            cell = rows[row_num].get(col)
            if cell is None or not _norm(cell.value):
                issues.append(
                    Issue(
                        category="BLANK_REQUIRED",
                        severity="HIGH",
                        sheet=sheet_name,
                        location=cell.ref if cell else f"Column {col}, Row {row_num}",
                        problem=f"Blank value in required column: {column_name}",
                        action="Supply the missing business value.",
                    )
                )
    return issues


def audit_workbook(
    path: str | Path,
    config: AuditConfig | None = None,
) -> AuditResult:
    config = config or AuditConfig(required_columns={})
    issues: list[Issue] = []

    with XlsxBook(path) as book:
        sheet_names = {sheet.name for sheet in book.sheets}
        for sheet in book.sheets:
            rows = book.rows(sheet)

            if sheet.state != "visible":
                issues.append(
                    Issue(
                        category="HIDDEN_SHEET",
                        severity="MEDIUM",
                        sheet=sheet.name,
                        location="-",
                        problem=f"Sheet state is {sheet.state}",
                        action="Confirm the sheet is intentionally hidden.",
                    )
                )

            for row in rows.values():
                for cell in row.values():
                    value_text = _norm(cell.value)
                    if cell.data_type == "e" or value_text in ERROR_MARKERS:
                        issues.append(
                            Issue(
                                category="ERROR_VALUE",
                                severity="HIGH",
                                sheet=sheet.name,
                                location=cell.ref,
                                problem=f"Excel error value: {value_text}",
                                action="Review the source formula or broken reference.",
                            )
                        )

                    if cell.formula:
                        missing = sorted(
                            ref for ref in _formula_sheet_refs(cell.formula)
                            if ref not in sheet_names
                        )
                        for ref in missing:
                            issues.append(
                                Issue(
                                    category="MISSING_SHEET_REFERENCE",
                                    severity="HIGH",
                                    sheet=sheet.name,
                                    location=cell.ref,
                                    problem=f"Formula references missing sheet: {ref}",
                                    action="Repair the formula or restore the referenced sheet.",
                                )
                            )

            issues.extend(_audit_duplicates(sheet.name, rows))
            issues.extend(
                _audit_required(
                    sheet.name,
                    rows,
                    config.required_columns.get(sheet.name, []),
                )
            )

    return AuditResult(workbook=str(Path(path)), issues=issues)


def write_report_text(result: AuditResult, path: str | Path) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        "RA Excel Assurance Engine v0.1",
        f"Workbook: {result.workbook}",
        f"Issues: {len(result.issues)}",
        "",
        "Summary by category:",
    ]
    for category, count in sorted(result.counts.items()):
        lines.append(f"- {category}: {count}")
    if not result.counts:
        lines.append("- None")

    lines.extend(["", "Detailed findings:"])
    if not result.issues:
        lines.append("OK - no issues found.")
    for index, issue in enumerate(result.issues, 1):
        lines.extend(
            [
                f"[{index:03d}] {issue.severity} | {issue.category}",
                f"Sheet: {issue.sheet} | Location: {issue.location}",
                f"Problem: {issue.problem}",
                f"Action: {issue.action}",
                "",
            ]
        )
    path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def write_report_json(result: AuditResult, path: str | Path) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(result.to_dict(), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
def create_safe_copy(source: str | Path, target: str | Path) -> str:
    """Create the v0.1 review copy without changing business data."""
    source = Path(source)
    target = Path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, target)

    digest = hashlib.sha256()
    with target.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()
