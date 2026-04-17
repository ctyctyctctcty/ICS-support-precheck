from __future__ import annotations

import html
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Sequence
from xml.etree import ElementTree as ET

NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
XML_NS = {'m': NS_MAIN, 'r': NS_REL}


@dataclass
class SheetData:
    name: str
    rows: List[List[str]]


def _col_index(cell_ref: str) -> int:
    letters = ''.join(ch for ch in cell_ref if ch.isalpha()).upper()
    index = 0
    for ch in letters:
        index = index * 26 + (ord(ch) - ord('A') + 1)
    return index - 1


def _read_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    if 'xl/sharedStrings.xml' not in zf.namelist():
        return []
    root = ET.fromstring(zf.read('xl/sharedStrings.xml'))
    values: List[str] = []
    for si in root.findall('m:si', XML_NS):
        parts = []
        for t in si.findall('.//m:t', XML_NS):
            parts.append(t.text or '')
        values.append(''.join(parts))
    return values


def _workbook_sheets(zf: zipfile.ZipFile) -> List[tuple[str, str]]:
    workbook = ET.fromstring(zf.read('xl/workbook.xml'))
    rels_root = ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
    rels: Dict[str, str] = {}
    for rel in rels_root:
        rel_id = rel.attrib.get('Id')
        target = rel.attrib.get('Target', '')
        if rel_id:
            rels[rel_id] = 'xl/' + target.lstrip('/') if not target.startswith('xl/') else target

    sheets: List[tuple[str, str]] = []
    for sheet in workbook.findall('m:sheets/m:sheet', XML_NS):
        name = sheet.attrib.get('name', 'Sheet')
        rel_id = sheet.attrib.get(f'{{{NS_REL}}}id')
        target = rels.get(rel_id or '')
        if target:
            sheets.append((name, target))
    return sheets


def _cell_text(cell: ET.Element, shared_strings: Sequence[str]) -> str:
    cell_type = cell.attrib.get('t')
    if cell_type == 'inlineStr':
        parts = [t.text or '' for t in cell.findall('.//m:t', XML_NS)]
        return ''.join(parts).strip()

    value = cell.find('m:v', XML_NS)
    raw = value.text if value is not None else ''
    if raw is None:
        raw = ''

    if cell_type == 's':
        try:
            return shared_strings[int(raw)].strip()
        except (ValueError, IndexError):
            return ''
    return str(raw).strip()


def read_workbook(path: Path) -> List[SheetData]:
    sheets: List[SheetData] = []
    with zipfile.ZipFile(path) as zf:
        shared_strings = _read_shared_strings(zf)
        for sheet_name, sheet_path in _workbook_sheets(zf):
            if sheet_path not in zf.namelist():
                continue
            root = ET.fromstring(zf.read(sheet_path))
            rows: List[List[str]] = []
            for row in root.findall('.//m:sheetData/m:row', XML_NS):
                values: Dict[int, str] = {}
                max_index = -1
                for cell in row.findall('m:c', XML_NS):
                    ref = cell.attrib.get('r', '')
                    idx = _col_index(ref)
                    values[idx] = _cell_text(cell, shared_strings)
                    max_index = max(max_index, idx)
                if max_index >= 0:
                    rows.append([values.get(i, '') for i in range(max_index + 1)])
            sheets.append(SheetData(sheet_name, rows))
    return sheets


def workbook_as_text(path: Path) -> str:
    parts = []
    for sheet in read_workbook(path):
        parts.append(f'## {sheet.name}')
        for row_index, row in enumerate(sheet.rows, start=1):
            if any(cell.strip() for cell in row):
                parts.append(f'{row_index}: ' + '\t'.join(row))
    return '\n'.join(parts)


def _column_letter(index: int) -> str:
    index += 1
    letters = ''
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord('A') + remainder) + letters
    return letters


def _sheet_xml(headers: Sequence[str], rows: Sequence[Dict[str, str]]) -> str:
    all_rows = [list(headers)]
    for row in rows:
        all_rows.append([str(row.get(header, '') or '') for header in headers])

    row_xml = []
    for r_idx, row in enumerate(all_rows, start=1):
        cells = []
        for c_idx, value in enumerate(row):
            ref = f'{_column_letter(c_idx)}{r_idx}'
            safe = html.escape(value)
            cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{safe}</t></is></c>')
        row_xml.append(f'<row r="{r_idx}">{"".join(cells)}</row>')
    dimension = f'A1:{_column_letter(len(headers) - 1)}{len(all_rows)}'
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">
  <dimension ref="{dimension}"/>
  <sheetData>{''.join(row_xml)}</sheetData>
</worksheet>'''


def write_standard_workbook(path: Path, headers: Sequence[str], rows: Sequence[Dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    sheet_xml = _sheet_xml(headers, rows)
    with zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>''')
        zf.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>''')
        zf.writestr('xl/workbook.xml', f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>''')
        zf.writestr('xl/_rels/workbook.xml.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>''')
        zf.writestr('xl/worksheets/sheet1.xml', sheet_xml)