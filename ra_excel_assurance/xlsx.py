from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

CELL_REF_RE = re.compile(r"([A-Z]+)(\d+)$")


def _q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def column_index(cell_ref: str) -> int:
    match = CELL_REF_RE.match(cell_ref.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    letters = match.group(1)
    value = 0
    for char in letters:
        value = value * 26 + (ord(char) - 64)
    return value
def _coerce_number(raw: str | None):
    if raw is None or raw == "":
        return None
    try:
        if "." not in raw and "e" not in raw.lower():
            return int(raw)
        return float(raw)
    except ValueError:
        return raw


@dataclass(frozen=True)
class Cell:
    ref: str
    value: object | None
    formula: str | None
    data_type: str | None


@dataclass(frozen=True)
class Sheet:
    name: str
    path: str
    state: str = "visible"


class XlsxBook:
    """Minimal OOXML reader for the auditing checks in v0.1."""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        self._zip = ZipFile(self.path, "r")
        self.shared_strings = self._load_shared_strings()
        self.sheets = self._load_sheets()

    def close(self) -> None:
        self._zip.close()

    def __enter__(self) -> "XlsxBook":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def _load_shared_strings(self) -> list[str]:
        try:
            root = ET.fromstring(self._zip.read("xl/sharedStrings.xml"))
        except KeyError:
            return []
        values: list[str] = []
        for item in root.findall(_q(MAIN_NS, "si")):
            parts = [node.text or "" for node in item.iter(_q(MAIN_NS, "t"))]
            values.append("".join(parts))
        return values

    def _load_sheets(self) -> list[Sheet]:
        workbook = ET.fromstring(self._zip.read("xl/workbook.xml"))
        rels_root = ET.fromstring(self._zip.read("xl/_rels/workbook.xml.rels"))
        targets = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels_root.findall(_q(PKG_REL_NS, "Relationship"))
        }
        result: list[Sheet] = []
        sheets_node = workbook.find(_q(MAIN_NS, "sheets"))
        if sheets_node is None:
            return result
        for node in sheets_node.findall(_q(MAIN_NS, "sheet")):
            rel_id = node.attrib.get(_q(REL_NS, "id"))
            target = targets.get(rel_id or "")
            if not target:
                continue
            target = target.lstrip("/")
            if not target.startswith("xl/"):
                target = f"xl/{target}"
            result.append(
                Sheet(
                    name=node.attrib.get("name", ""),
                    path=target,
                    state=node.attrib.get("state", "visible"),
                )
            )
        return result

    def rows(self, sheet: Sheet) -> dict[int, dict[int, Cell]]:
        root = ET.fromstring(self._zip.read(sheet.path))
        data = root.find(_q(MAIN_NS, "sheetData"))
        rows: dict[int, dict[int, Cell]] = {}
        if data is None:
            return rows
        for row_node in data.findall(_q(MAIN_NS, "row")):
            row_num = int(row_node.attrib.get("r", "0") or 0)
            row_cells: dict[int, Cell] = {}
            for cell_node in row_node.findall(_q(MAIN_NS, "c")):
                ref = cell_node.attrib.get("r", "")
                if not ref:
                    continue
                col = column_index(ref)
                data_type = cell_node.attrib.get("t")
                formula_node = cell_node.find(_q(MAIN_NS, "f"))
                formula = None
                if formula_node is not None and formula_node.text:
                    formula = "=" + formula_node.text
                value = self._cell_value(cell_node, data_type)
                row_cells[col] = Cell(
                    ref=ref,
                    value=value,
                    formula=formula,
                    data_type=data_type,
                )
            rows[row_num] = row_cells
        return rows
    def _cell_value(self, node: ET.Element, data_type: str | None):
        if data_type == "inlineStr":
            inline = node.find(_q(MAIN_NS, "is"))
            if inline is None:
                return ""
            return "".join(x.text or "" for x in inline.iter(_q(MAIN_NS, "t")))

        value_node = node.find(_q(MAIN_NS, "v"))
        raw = value_node.text if value_node is not None else None
        if data_type == "s" and raw is not None:
            try:
                return self.shared_strings[int(raw)]
            except (ValueError, IndexError):
                return raw
        if data_type == "b":
            return raw == "1"
        if data_type in {"str", "e"}:
            return raw or ""
        return _coerce_number(raw)
