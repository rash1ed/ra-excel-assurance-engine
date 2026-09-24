from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)


def _q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def _col_letter(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result
def _write_cell(parent, row_num: int, col_num: int, value) -> None:
    if value is None:
        return
    ref = f"{_col_letter(col_num)}{row_num}"
    if isinstance(value, dict) and "formula" in value:
        cell = ET.SubElement(parent, _q(MAIN_NS, "c"), {"r": ref})
        formula = ET.SubElement(cell, _q(MAIN_NS, "f"))
        formula.text = str(value["formula"]).lstrip("=")
        cached = ET.SubElement(cell, _q(MAIN_NS, "v"))
        cached.text = str(value.get("cached", 0))
        return
    if isinstance(value, dict) and "error" in value:
        cell = ET.SubElement(
            parent, _q(MAIN_NS, "c"), {"r": ref, "t": "e"}
        )
        node = ET.SubElement(cell, _q(MAIN_NS, "v"))
        node.text = str(value["error"])
        return
    if isinstance(value, str):
        cell = ET.SubElement(
            parent, _q(MAIN_NS, "c"), {"r": ref, "t": "inlineStr"}
        )
        inline = ET.SubElement(cell, _q(MAIN_NS, "is"))
        text = ET.SubElement(inline, _q(MAIN_NS, "t"))
        text.text = value
        return
    if isinstance(value, bool):
        cell = ET.SubElement(
            parent, _q(MAIN_NS, "c"), {"r": ref, "t": "b"}
        )
        node = ET.SubElement(cell, _q(MAIN_NS, "v"))
        node.text = "1" if value else "0"
        return
    cell = ET.SubElement(parent, _q(MAIN_NS, "c"), {"r": ref})
    node = ET.SubElement(cell, _q(MAIN_NS, "v"))
    node.text = str(value)


def _worksheet_xml(rows: list[list[object]]) -> bytes:
    root = ET.Element(_q(MAIN_NS, "worksheet"))
    data = ET.SubElement(root, _q(MAIN_NS, "sheetData"))
    for row_num, row_values in enumerate(rows, start=1):
        row = ET.SubElement(data, _q(MAIN_NS, "row"), {"r": str(row_num)})
        for col_num, value in enumerate(row_values, start=1):
            _write_cell(row, row_num, col_num, value)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)
def make_xlsx(path: str | Path, sheets: list[dict]) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    workbook = ET.Element(_q(MAIN_NS, "workbook"))
    sheets_node = ET.SubElement(workbook, _q(MAIN_NS, "sheets"))
    rels = ET.Element(_q(PKG_REL_NS, "Relationships"))

    for index, sheet in enumerate(sheets, start=1):
        attrs = {
            "name": sheet["name"],
            "sheetId": str(index),
            _q(REL_NS, "id"): f"rId{index}",
        }
        if sheet.get("state", "visible") != "visible":
            attrs["state"] = sheet["state"]
        ET.SubElement(sheets_node, _q(MAIN_NS, "sheet"), attrs)
        ET.SubElement(
            rels,
            _q(PKG_REL_NS, "Relationship"),
            {
                "Id": f"rId{index}",
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                "Target": f"worksheets/sheet{index}.xml",
            },
        )
    content = ET.Element(_q(CONTENT_NS, "Types"))
    ET.SubElement(
        content,
        _q(CONTENT_NS, "Default"),
        {"Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml"},
    )
    ET.SubElement(
        content,
        _q(CONTENT_NS, "Default"),
        {"Extension": "xml", "ContentType": "application/xml"},
    )
    ET.SubElement(
        content,
        _q(CONTENT_NS, "Override"),
        {
            "PartName": "/xl/workbook.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        },
    )
    for index in range(1, len(sheets) + 1):
        ET.SubElement(
            content,
            _q(CONTENT_NS, "Override"),
            {
                "PartName": f"/xl/worksheets/sheet{index}.xml",
                "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            },
        )
    root_rels = ET.Element(_q(PKG_REL_NS, "Relationships"))
    ET.SubElement(
        root_rels,
        _q(PKG_REL_NS, "Relationship"),
        {
            "Id": "rId1",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "Target": "xl/workbook.xml",
        },
    )

    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            ET.tostring(content, encoding="utf-8", xml_declaration=True),
        )
        archive.writestr(
            "_rels/.rels",
            ET.tostring(root_rels, encoding="utf-8", xml_declaration=True),
        )
        archive.writestr(
            "xl/workbook.xml",
            ET.tostring(workbook, encoding="utf-8", xml_declaration=True),
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            ET.tostring(rels, encoding="utf-8", xml_declaration=True),
        )
        for index, sheet in enumerate(sheets, start=1):
            archive.writestr(
                f"xl/worksheets/sheet{index}.xml",
                _worksheet_xml(sheet["rows"]),
            )
    return path
