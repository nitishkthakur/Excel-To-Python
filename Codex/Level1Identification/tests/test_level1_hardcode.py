#!/usr/bin/env python3
from __future__ import annotations

import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from level1_hardcode import NS_MAIN, parse_xml, scan_target_cached_errors, transform_workbook


def _xlsx_template_parts() -> dict[str, str]:
    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""

    workbook = """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Input 1" sheetId="1" r:id="rId1"/>
    <sheet name="Calc2" sheetId="2" r:id="rId2"/>
    <sheet name="Target" sheetId="3" r:id="rId3"/>
    <sheet name="Other" sheetId="4" r:id="rId4"/>
  </sheets>
  <definedNames>
    <definedName name="RATE">'Input 1'!$A$1</definedName>
    <definedName name="LOCAL_TMP" localSheetId="3">Other!$A$1</definedName>
  </definedNames>
</workbook>
"""

    workbook_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""

    styles = """<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="1"><xf xfId="0"/></cellXfs>
</styleSheet>
"""

    sheet1 = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>10</v></c>
      <c r="B1"><f>A1*2</f><v>20</v></c>
    </row>
  </sheetData>
</worksheet>
"""

    sheet2 = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>'Input 1'!B1+5</f><v>25</v></c>
    </row>
  </sheetData>
</worksheet>
"""

    sheet3 = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>'Input 1'!B1+Calc2!A1</f><v>45</v></c>
      <c r="B1"><f>RATE*2</f><v>20</v></c>
    </row>
  </sheetData>
</worksheet>
"""

    sheet4 = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>1+2</f><v>3</v></c>
    </row>
  </sheetData>
</worksheet>
"""

    return {
        "[Content_Types].xml": content_types,
        "_rels/.rels": root_rels,
        "xl/workbook.xml": workbook,
        "xl/_rels/workbook.xml.rels": workbook_rels,
        "xl/styles.xml": styles,
        "xl/worksheets/sheet1.xml": sheet1,
        "xl/worksheets/sheet2.xml": sheet2,
        "xl/worksheets/sheet3.xml": sheet3,
        "xl/worksheets/sheet4.xml": sheet4,
    }


def _write_template_workbook(path: Path) -> None:
    parts = _xlsx_template_parts()
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, xml in parts.items():
            zf.writestr(name, xml.encode("utf-8"))


def _sheet_formula_count(sheet_xml: bytes) -> int:
    root = parse_xml(sheet_xml)
    return len(root.xpath(".//x:f", namespaces={"x": NS_MAIN}))


def _sheet_names(workbook_xml: bytes) -> list[str]:
    root = parse_xml(workbook_xml)
    return [
        el.get("name")
        for el in root.xpath("./x:sheets/x:sheet", namespaces={"x": NS_MAIN})
    ]


def _defined_names(workbook_xml: bytes) -> dict[str, etree._Element]:
    root = parse_xml(workbook_xml)
    out: dict[str, etree._Element] = {}
    for dn in root.xpath("./x:definedNames/x:definedName", namespaces={"x": NS_MAIN}):
        name = dn.get("name")
        if name:
            out[name] = dn
    return out


def test_end_to_end_level1_hardcode() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        input_xlsx = tmp_dir / "input.xlsx"
        output_xlsx = tmp_dir / "output.xlsx"
        _write_template_workbook(input_xlsx)

        result = transform_workbook(
            input_xlsx=input_xlsx,
            output_xlsx=output_xlsx,
            target_sheet="Target",
            fail_on_target_errors=True,
        )

        assert result["predecessors"] == ["Calc2", "Input 1"]
        assert result["target_error_cells"] == []

        with zipfile.ZipFile(output_xlsx, "r") as zf:
            names = set(zf.namelist())

            assert "xl/worksheets/sheet4.xml" not in names
            assert "xl/worksheets/sheet1.xml" in names
            assert "xl/worksheets/sheet2.xml" in names
            assert "xl/worksheets/sheet3.xml" in names

            workbook_xml = zf.read("xl/workbook.xml")
            assert _sheet_names(workbook_xml) == ["Input 1", "Calc2", "Target"]

            dnames = _defined_names(workbook_xml)
            assert "RATE" in dnames
            assert "LOCAL_TMP" not in dnames

            sheet1_formula_count = _sheet_formula_count(zf.read("xl/worksheets/sheet1.xml"))
            sheet2_formula_count = _sheet_formula_count(zf.read("xl/worksheets/sheet2.xml"))
            sheet3_formula_count = _sheet_formula_count(zf.read("xl/worksheets/sheet3.xml"))

            assert sheet1_formula_count == 0
            assert sheet2_formula_count == 0
            assert sheet3_formula_count == 2

        errors = scan_target_cached_errors(output_xlsx, "Target")
        assert errors == []


def main() -> None:
    test_end_to_end_level1_hardcode()
    print("test_end_to_end_level1_hardcode: PASS")


if __name__ == "__main__":
    main()
