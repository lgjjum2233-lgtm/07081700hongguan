from __future__ import annotations

import argparse
import json
import xml.etree.ElementTree as ET
from pathlib import Path, PurePosixPath
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

import openpyxl


NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "officeRel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

for prefix, uri in [("", NS["main"]), ("r", NS["officeRel"])]:
    ET.register_namespace(prefix, uri)


def q(tag: str) -> str:
    return f"{{{NS['main']}}}{tag}"


def read_xml(z: ZipFile, name: str) -> ET.Element:
    return ET.fromstring(z.read(name))


def write_xml(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def rel_target(base: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base_dir = str(PurePosixPath(base).parent)
    parts: list[str] = []
    for part in (base_dir + "/" + target).split("/"):
        if part in {"", "."}:
            continue
        if part == "..":
            if parts:
                parts.pop()
        else:
            parts.append(part)
    return "/".join(parts)


def workbook_sheet_paths(z: ZipFile) -> dict[str, str]:
    wb = read_xml(z, "xl/workbook.xml")
    rels = read_xml(z, "xl/_rels/workbook.xml.rels")
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    out: dict[str, str] = {}
    for sheet in wb.findall(".//main:sheets/main:sheet", NS):
        rid = sheet.attrib[f"{{{NS['officeRel']}}}id"]
        out[sheet.attrib["name"]] = rel_target("xl/workbook.xml", rel_map[rid])
    return out


def cell_col_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    return openpyxl.utils.column_index_from_string(letters)


def find_row(sheet_root: ET.Element, row_num: int) -> ET.Element:
    sheet_data = sheet_root.find("main:sheetData", NS)
    if sheet_data is None:
        raise KeyError("sheetData")
    for row in sheet_data.findall("main:row", NS):
        if int(row.attrib["r"]) == row_num:
            return row
    new_row = ET.Element(q("row"), {"r": str(row_num)})
    sheet_data.append(new_row)
    return new_row


def find_cell(sheet_root: ET.Element, ref: str) -> ET.Element:
    row_num = int("".join(ch for ch in ref if ch.isdigit()))
    row = find_row(sheet_root, row_num)
    for cell in row.findall("main:c", NS):
        if cell.attrib.get("r") == ref:
            return cell
    cell = ET.Element(q("c"), {"r": ref})
    row.append(cell)
    return cell


def set_inline_string(cell: ET.Element, value: str) -> None:
    style = cell.attrib.get("s")
    for child in list(cell):
        cell.remove(child)
    cell.attrib.clear()
    cell.attrib["r"] = cell.attrib.get("r", "")
    if style is not None:
        cell.attrib["s"] = style
    cell.attrib["t"] = "inlineStr"
    is_el = ET.SubElement(cell, q("is"))
    t = ET.SubElement(is_el, q("t"))
    t.text = value


def set_cell_text(sheet_root: ET.Element, ref: str, value: str) -> None:
    cell = find_cell(sheet_root, ref)
    old_ref = cell.attrib.get("r", ref)
    style = cell.attrib.get("s")
    for child in list(cell):
        cell.remove(child)
    cell.attrib.clear()
    cell.attrib["r"] = old_ref
    if style is not None:
        cell.attrib["s"] = style
    cell.attrib["t"] = "inlineStr"
    is_el = ET.SubElement(cell, q("is"))
    t = ET.SubElement(is_el, q("t"))
    t.text = value


MASTER_TEXT = {
    "F75": (
        "A股和港股估值数据已补齐：沪深300PE为14.69倍、近三年分位98.4%，"
        "创业板指PE为51.81倍、近三年分位99.9%；恒生指数PE为11.99倍、"
        "近三年分位83.9%，恒生科技PE为22.50倍、近三年分位57.1%。"
        "整体看，A股核心和成长估值仍处近三年高位，港股传统指数估值偏高但边际回落，"
        "港股科技成长估值处于中位偏上。"
    ),
    "F81": "上证指数：PE为17.18倍，近三年分位约94.4%，环比下行0.06倍，估值处于高位。",
    "F82": "深成指数：PE为37.36倍，近三年分位约99.7%，环比上行0.28倍，估值处于高位。",
    "F83": "沪深300：PE为14.69倍，近三年分位约98.4%，环比下行0.10倍，估值处于高位。",
    "F84": "创业板指：PE为51.81倍，近三年分位约99.9%，环比上行0.90倍，估值处于高位。",
    "F85": "恒生指数：PE为11.99倍，近三年分位约83.9%，环比下行0.14倍，估值处于偏高位置。",
    "F86": "恒生科技：PE为22.50倍，近三年分位约57.1%，环比下行0.49倍，估值处于中性略偏高位置。",
}

DISPLAY_TEXT = {
    "U51": (
        "估值面：中性偏谨慎：A股主要指数估值仍处近三年高位，估值性价比偏弱；"
        "港股估值边际回落，成长板块压力相对较小。核心资产方面，沪深300PE 14.69倍、"
        "近三年约98.4%分位；成长板块方面，创业板指PE 51.81倍、近三年约99.9%分位；"
        "港股方面，恒生指数PE 11.99倍、近三年约83.9%分位，恒生科技PE 22.50倍、"
        "近三年约57.1%分位。"
    )
}


def update_texts(input_path: Path, output_path: Path) -> dict[str, Any]:
    replacements: dict[str, bytes] = {}
    updates: list[dict[str, str]] = []
    with ZipFile(input_path, "r") as zin:
        sheet_paths = workbook_sheet_paths(zin)
        for sheet_name, mapping in [
            ("重点策略跟踪情况(V3.0)", MASTER_TEXT),
            ("重点策略跟踪情况(V3)", DISPLAY_TEXT),
        ]:
            sheet_path = sheet_paths.get(sheet_name)
            if not sheet_path:
                continue
            root = read_xml(zin, sheet_path)
            for ref, text in mapping.items():
                set_cell_text(root, ref, text)
                updates.append({"sheet": sheet_name, "cell": ref})
            replacements[sheet_path] = write_xml(root)
        with ZipFile(output_path, "w", ZIP_DEFLATED) as zout:
            for name in zin.namelist():
                zout.writestr(name, replacements.get(name, zin.read(name)))
    return {"input": str(input_path), "output": str(output_path), "updates": updates}


def main() -> None:
    parser = argparse.ArgumentParser(description="Sync PE judgement text in TREE after PE source correction.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    print(json.dumps(update_texts(Path(args.input), Path(args.output)), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
