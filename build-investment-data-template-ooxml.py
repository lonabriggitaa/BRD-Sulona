from __future__ import annotations

from datetime import date
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape

ROOT = Path(r"C:\Users\LENOVO\Documents\Codex\2026-07-01\https-eldcm9-oss-github-io-brd")
OUTPUT = ROOT / "outputs" / "Investment_Data_Template.xlsx"
SCAFFOLD = ROOT / "outputs" / "laravel-excel-backend-scaffold" / "storage" / "app" / "templates" / "Investment_Data_Template.xlsx"

LOOKUPS = {
    "Division": ["Engineering", "Finance", "Operations", "Supply Chain", "Quality", "Logistics", "Procurement"],
    "Site": ["Bandung", "Jakarta", "Surabaya", "Medan", "Semarang", "Makassar"],
    "Status": ["Draft", "Submitted", "Approved", "Running", "Completed", "Closed", "Cancelled"],
    "Document Type": ["BRD", "Financial Analysis", "Quotation", "Contract", "Progress Report", "LPJ", "Final Report", "Others"],
    "Priority": ["Low", "Medium", "High"],
    "Activity Status": ["Planned", "Running", "Completed", "Delayed", "Cancelled"],
}

SHEETS = {
    "PROJECT_MASTER": {
        "headers": ["BRD Number", "Project Name", "Division", "Site", "Vendor", "PIC", "Budget Requested", "Budget Approved", "Actual Cost", "Planned Start Date", "Target Finish Date", "Progress (%)", "Status", "Remarks"],
        "rows": [["BRD-2026-001", "Machine Blowing", "Engineering", "Bandung", "PT ABC", "Andi", 850000000, 800000000, 350000000, date(2026, 7, 1), date(2026, 9, 30), 45, "Running", "Material already delivered."]],
    },
    "PROJECT_ACTIVITY": {
        "headers": ["Activity Date", "BRD Number", "Project Name", "Activity", "Description", "PIC", "Progress (%)", "Status", "Notes"],
        "rows": [
            [date(2026, 7, 1), "BRD-2026-001", "Machine Blowing", "Kick Off Meeting", "Project officially started.", "Andi", 5, "Running", "-"],
            [date(2026, 7, 5), "BRD-2026-001", "Machine Blowing", "Vendor Discussion", "Vendor submitted quotation.", "Andi", 10, "Running", "Quotation approved."],
            [date(2026, 7, 10), "BRD-2026-001", "Machine Blowing", "Material Delivery", "Materials delivered to site.", "Andi", 35, "Running", "Warehouse checked."],
        ],
    },
    "PROJECT_DOCUMENT": {
        "headers": ["BRD Number", "Project Name", "Document Type", "Document Name", "Version", "Upload Date", "Remarks"],
        "rows": [["BRD-2026-001", "Machine Blowing", "BRD", "BRD Rev 01", "1.0", date(2026, 7, 1), "Approved"]],
    },
}

README_ROWS = [
    ["Investment Data Template - Instructions"],
    ["1. Fill PROJECT_MASTER for new or updated projects."],
    ["2. Fill PROJECT_ACTIVITY for project activity updates."],
    ["3. Fill PROJECT_DOCUMENT for document metadata."],
    ["4. Do not change sheet names."],
    ["5. Do not change column names."],
    ["6. Upload this workbook in Import Update menu."],
    ["7. Dashboard will update automatically after successful import."],
]


def col_name(index: int) -> str:
    result = ""
    while index:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result


def excel_serial(value: date) -> int:
    return (value - date(1899, 12, 30)).days


def cell(ref: str, value, style: int = 0) -> str:
    style_attr = f' s="{style}"' if style else ""
    if value is None or value == "":
        return f'<c r="{ref}"{style_attr}/>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"{style_attr}><v>{value}</v></c>'
    if isinstance(value, date):
        return f'<c r="{ref}" s="4"><v>{excel_serial(value)}</v></c>'
    return f'<c r="{ref}" t="inlineStr"{style_attr}><is><t>{escape(str(value))}</t></is></c>'


def validations_for(name: str, headers: list[str]) -> str:
    validations = []
    lookup_col = {"Division": "A", "Site": "B", "Status": "C", "Document Type": "D", "Priority": "E", "Activity Status": "F"}
    list_targets = {}
    if name == "PROJECT_MASTER":
        list_targets = {"Division": "Division", "Site": "Site", "Status": "Status"}
    elif name == "PROJECT_ACTIVITY":
        list_targets = {"Status": "Activity Status"}
    elif name == "PROJECT_DOCUMENT":
        list_targets = {"Document Type": "Document Type"}

    for header, lookup_name in list_targets.items():
        if header in headers:
            col = col_name(headers.index(header) + 1)
            lookup_letter = lookup_col[lookup_name]
            count = len(LOOKUPS[lookup_name]) + 1
            validations.append(f'<dataValidation type="list" allowBlank="1" sqref="{col}2:{col}500"><formula1>LOOKUP!${lookup_letter}$2:${lookup_letter}${count}</formula1></dataValidation>')

    if "Progress (%)" in headers:
        col = col_name(headers.index("Progress (%)") + 1)
        validations.append(f'<dataValidation type="whole" operator="between" allowBlank="1" sqref="{col}2:{col}500"><formula1>0</formula1><formula2>100</formula2></dataValidation>')

    return f'<dataValidations count="{len(validations)}">{"".join(validations)}</dataValidations>' if validations else ""


def conditional_formatting(name: str, headers: list[str]) -> str:
    blocks = []
    if "Progress (%)" in headers:
        col = col_name(headers.index("Progress (%)") + 1)
        ref = f"{col}2:{col}500"
        blocks.append(f'''<conditionalFormatting sqref="{ref}">
<cfRule type="cellIs" priority="1" operator="between" dxfId="0"><formula>0</formula><formula>30</formula></cfRule>
<cfRule type="cellIs" priority="2" operator="between" dxfId="1"><formula>31</formula><formula>70</formula></cfRule>
<cfRule type="cellIs" priority="3" operator="between" dxfId="2"><formula>71</formula><formula>100</formula></cfRule>
</conditionalFormatting>''')
    if "Status" in headers:
        col = col_name(headers.index("Status") + 1)
        ref = f"{col}2:{col}500"
        first = f"{col}2"
        blocks.append(f'''<conditionalFormatting sqref="{ref}">
<cfRule type="expression" priority="4" dxfId="3"><formula>${first}="Running"</formula></cfRule>
<cfRule type="expression" priority="5" dxfId="2"><formula>${first}="Completed"</formula></cfRule>
<cfRule type="expression" priority="6" dxfId="0"><formula>OR(${first}="Delayed",${first}="Cancelled")</formula></cfRule>
<cfRule type="expression" priority="7" dxfId="4"><formula>OR(${first}="Draft",${first}="Submitted")</formula></cfRule>
</conditionalFormatting>''')
    return "".join(blocks)


def data_sheet_xml(name: str, config: dict) -> str:
    headers = config["headers"]
    last_col = col_name(len(headers))
    rows = []
    rows.append(f'<row r="1">{"".join(cell(f"{col_name(i)}1", h, 3) for i, h in enumerate(headers, start=1))}</row>')
    money_headers = {"Budget Requested", "Budget Approved", "Actual Cost"}
    for row_idx, data in enumerate(config["rows"], start=2):
        cells = []
        for col_idx, value in enumerate(data, start=1):
            style = 5 if headers[col_idx - 1] in money_headers else 0
            cells.append(cell(f"{col_name(col_idx)}{row_idx}", value, style))
        rows.append(f'<row r="{row_idx}">{"".join(cells)}</row>')
    widths = ''.join(f'<col min="{i}" max="{i}" width="{max(13, min(34, len(header) + 5))}" customWidth="1"/>' for i, header in enumerate(headers, start=1))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<cols>{widths}</cols>
<sheetData>{"".join(rows)}</sheetData>
{validations_for(name, headers)}
{conditional_formatting(name, headers)}
<autoFilter ref="A1:{last_col}{max(2, len(config["rows"]) + 1)}"/>
</worksheet>'''


def lookup_xml() -> str:
    headers = list(LOOKUPS)
    max_len = max(len(v) for v in LOOKUPS.values())
    rows = [f'<row r="1">{"".join(cell(f"{col_name(i)}1", h, 3) for i, h in enumerate(headers, start=1))}</row>']
    for row_idx in range(2, max_len + 2):
        values = []
        for col_idx, header in enumerate(headers, start=1):
            values.append(cell(f"{col_name(col_idx)}{row_idx}", LOOKUPS[header][row_idx - 2] if row_idx - 2 < len(LOOKUPS[header]) else ""))
        rows.append(f'<row r="{row_idx}">{"".join(values)}</row>')
    widths = ''.join(f'<col min="{i}" max="{i}" width="22" customWidth="1"/>' for i in range(1, len(headers) + 1))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<cols>{widths}</cols><sheetData>{"".join(rows)}</sheetData><autoFilter ref="A1:{col_name(len(headers))}{max_len + 1}"/></worksheet>'''


def readme_xml() -> str:
    rows = []
    for idx, row in enumerate(README_ROWS, start=1):
        style = 1 if idx == 1 else 0
        rows.append(f'<row r="{idx}">{cell(f"A{idx}", row[0], style)}</row>')
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>
<cols><col min="1" max="1" width="90" customWidth="1"/></cols><sheetData>{"".join(rows)}</sheetData></worksheet>'''


ALL_SHEETS = ["PROJECT_MASTER", "PROJECT_ACTIVITY", "PROJECT_DOCUMENT", "LOOKUP", "README"]
CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>''' + ''.join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, 6)) + '</Types>'
ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'''
WORKBOOK_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>''' + ''.join(f'<sheet name="{name}" sheetId="{i}" r:id="rId{i}"/>' for i, name in enumerate(ALL_SHEETS, start=1)) + '</sheets></workbook>'
WORKBOOK_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">''' + ''.join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, 6)) + '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/><numFmt numFmtId="165" formatCode="&quot;Rp&quot; #,##0"/></numFmts><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0F172A"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2563EB"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"><color rgb="FFCBD5E1"/></left><right style="thin"><color rgb="FFCBD5E1"/></right><top style="thin"><color rgb="FFCBD5E1"/></top><bottom style="thin"><color rgb="FFCBD5E1"/></bottom></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/></cellXfs><dxfs count="5"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFEE2E2"/></patternFill></fill><font><color rgb="FF991B1B"/></font></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFEF3C7"/></patternFill></fill><font><color rgb="FF92400E"/></font></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFDCFCE7"/></patternFill></fill><font><color rgb="FF166534"/></font></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFDBEAFE"/></patternFill></fill><font><color rgb="FF1D4ED8"/></font></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFE5E7EB"/></patternFill></fill><font><color rgb="FF374151"/></font></dxf></dxfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'''
APP = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Codex</Application></Properties>'''
CORE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Investment Data Template</dc:title><dc:creator>Codex</dc:creator></cp:coreProperties>'''

OUTPUT.parent.mkdir(parents=True, exist_ok=True)
SCAFFOLD.parent.mkdir(parents=True, exist_ok=True)
with ZipFile(OUTPUT, "w", ZIP_DEFLATED) as zf:
    zf.writestr("[Content_Types].xml", CONTENT_TYPES)
    zf.writestr("_rels/.rels", ROOT_RELS)
    zf.writestr("xl/workbook.xml", WORKBOOK_XML)
    zf.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
    zf.writestr("xl/styles.xml", STYLES)
    zf.writestr("docProps/app.xml", APP)
    zf.writestr("docProps/core.xml", CORE)
    zf.writestr("xl/worksheets/sheet1.xml", data_sheet_xml("PROJECT_MASTER", SHEETS["PROJECT_MASTER"]))
    zf.writestr("xl/worksheets/sheet2.xml", data_sheet_xml("PROJECT_ACTIVITY", SHEETS["PROJECT_ACTIVITY"]))
    zf.writestr("xl/worksheets/sheet3.xml", data_sheet_xml("PROJECT_DOCUMENT", SHEETS["PROJECT_DOCUMENT"]))
    zf.writestr("xl/worksheets/sheet4.xml", lookup_xml())
    zf.writestr("xl/worksheets/sheet5.xml", readme_xml())

SCAFFOLD.write_bytes(OUTPUT.read_bytes())
print(OUTPUT)
