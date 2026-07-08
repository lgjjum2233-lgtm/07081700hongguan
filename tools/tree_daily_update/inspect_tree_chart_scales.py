from __future__ import annotations

import json
from pathlib import Path, PurePosixPath
from zipfile import ZipFile
import xml.etree.ElementTree as ET


PATH = Path("（6月9日）TREE宏观分析数据资讯新增指标.xlsx")
SHEET = "重点策略跟踪情况(V3.0)"
ROWS = [25, 42, 60, 65, 88, 98, 107, 139, 155]

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "officeRel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}


def read_xml(z, name):
    return ET.fromstring(z.read(name))


def rel_target(base, target):
    if target.startswith("/"):
        return target.lstrip("/")
    base_dir = str(PurePosixPath(base).parent)
    parts = []
    for part in (base_dir + "/" + target).split("/"):
        if part in ("", "."):
            continue
        if part == "..":
            if parts:
                parts.pop()
        else:
            parts.append(part)
    return "/".join(parts)


def workbook_sheet_path(z, sheet_name):
    wb = read_xml(z, "xl/workbook.xml")
    rels = read_xml(z, "xl/_rels/workbook.xml.rels")
    rel_map = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
    for sheet in wb.findall(".//main:sheets/main:sheet", NS):
        if sheet.attrib["name"] == sheet_name:
            rid = sheet.attrib[f"{{{NS['officeRel']}}}id"]
            return rel_target("xl/workbook.xml", rel_map[rid])
    raise KeyError(sheet_name)


def sheet_drawing_path(z, sheet_path):
    sheet = read_xml(z, sheet_path)
    drawing = sheet.find("main:drawing", NS)
    rid = drawing.attrib[f"{{{NS['officeRel']}}}id"]
    rels_name = str(PurePosixPath(sheet_path).parent / "_rels" / (PurePosixPath(sheet_path).name + ".rels"))
    rels = read_xml(z, rels_name)
    for rel in rels:
        if rel.attrib["Id"] == rid:
            return rel_target(sheet_path, rel.attrib["Target"])
    raise KeyError(rid)


def drawing_rels_path(drawing_path):
    p = PurePosixPath(drawing_path)
    return str(p.parent / "_rels" / (p.name + ".rels"))


def chart_rows(z, drawing_path):
    rels = read_xml(z, drawing_rels_path(drawing_path))
    rel_map = {r.attrib["Id"]: rel_target(drawing_path, r.attrib["Target"]) for r in rels}
    drawing = read_xml(z, drawing_path)
    rows = {}
    for anchor in list(drawing):
        frm = anchor.find("xdr:from", NS)
        chart = anchor.find(".//a:graphicData/c:chart", NS)
        if frm is None or chart is None:
            continue
        rid = chart.attrib[f"{{{NS['officeRel']}}}id"]
        rows[int(frm.findtext("xdr:row", namespaces=NS)) + 1] = rel_map[rid]
    return rows


def cache_values(chart_xml):
    root = ET.fromstring(chart_xml)
    out = []
    for pt in root.findall(".//c:lineChart/c:ser/c:val/c:numRef/c:numCache/c:pt", NS):
        text = pt.findtext("c:v", namespaces=NS)
        if text is None:
            continue
        try:
            out.append(float(text))
        except ValueError:
            pass
    return out


def main():
    with ZipFile(PATH) as z:
        sheet_path = workbook_sheet_path(z, SHEET)
        drawing_path = sheet_drawing_path(z, sheet_path)
        rows = chart_rows(z, drawing_path)
        result = {}
        for row in ROWS:
            path = rows.get(row)
            values = cache_values(z.read(path)) if path else []
            result[row] = {
                "chart_path": path,
                "count": len(values),
                "first": values[0] if values else None,
                "last": values[-1] if values else None,
                "min": min(values) if values else None,
                "max": max(values) if values else None,
            }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
