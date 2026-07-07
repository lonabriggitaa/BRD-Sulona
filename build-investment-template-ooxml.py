from __future__ import annotations

from datetime import date
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape

ROOT = Path(r"C:\Users\LENOVO\Documents\Codex\2026-07-01\https-eldcm9-oss-github-io-brd")
OUTPUT = ROOT / "outputs" / "Investment_Backend_Template.xlsx"
SCAFFOLD = ROOT / "outputs" / "laravel-excel-backend-scaffold" / "storage" / "app" / "templates" / "Investment_Backend_Template.xlsx"

PROJECT_STATUS = "Draft,Submitted,Approved,Running,Completed,Closed,Cancelled"
ACTIVITY_STATUS = "Planned,Running,Completed,Delayed,Cancelled"
PRIORITY = "Low,Medium,High"

SHEETS = {
    "MASTER_PROJECT": {
        "headers": [
            "project_code", "brd_number", "project_name", "division", "department", "site", "category",
            "priority", "vendor", "pic", "budget_requested", "budget_approved", "actual_cost",
            "planned_start_date", "planned_finish_date", "actual_start_date", "actual_finish_date",
            "progress", "status", "remarks",
        ],
        "required": {"project_code", "project_name", "division", "category", "pic", "budget_approved", "planned_start_date", "planned_finish_date", "status"},
        "rows": [
            ["PRJ-ERP-001", "BRD-26002", "ERP Finance Extension", "Finance", "Accounting", "Bandung", "IT System", "High", "SAP Partner Indonesia", "Dina Putri", 9500000000, 9200000000, 2100000000, date(2026, 3, 1), date(2026, 12, 15), date(2026, 3, 5), None, 22, "Running", "Finance module extension for investment dashboard."],
            ["PRJ-WH-001", "BRD-26001", "Warehouse Automation Line", "Supply Chain", "Warehouse", "Jakarta", "Automation", "High", "Siemens", "Ari Wibowo", 20000000000, 18500000000, 13600000000, date(2026, 1, 10), date(2026, 9, 30), date(2026, 1, 12), None, 68, "Running", "Automation line rollout."],
        ],
    },
    "PROJECT_ACTIVITY": {
        "headers": ["activity_id", "project_code", "activity_date", "activity_type", "activity_title", "description", "pic", "progress", "status", "notes"],
        "required": {"project_code", "activity_date", "activity_type", "activity_title", "status"},
        "rows": [
            ["ACT-ERP-001", "PRJ-ERP-001", date(2026, 6, 3), "Milestone", "Procurement Review", "Committee notes closed.", "Dina Putri", 22, "Completed", "Validated in weekly review."],
            ["ACT-WH-001", "PRJ-WH-001", date(2026, 6, 18), "Review", "PLC Integration Retest", "Retest conveyor PLC integration.", "Ari Wibowo", 68, "Running", "Safety interlock check included."],
        ],
    },
    "PROJECT_DOCUMENT": {
        "headers": ["document_id", "project_code", "document_type", "document_name", "version", "upload_date", "file_name", "remarks"],
        "required": {"project_code", "document_type", "document_name"},
        "rows": [
            ["DOC-ERP-001", "PRJ-ERP-001", "Cost Benefit Analysis", "ERP Finance CBA", "v1.0", date(2026, 6, 5), "erp_finance_cba.pdf", "Metadata only; actual file upload is separate."],
            ["DOC-WH-001", "PRJ-WH-001", "Financial Summary", "Warehouse Automation Financial Summary", "v1.1", date(2026, 6, 20), "warehouse_financial_summary.xlsx", "Metadata only."],
        ],
    },
    "PROJECT_REVIEW": {
        "headers": ["review_id", "project_code", "executive_summary", "timeline_performance", "budget_performance", "activity_summary", "benefit_realization", "lessons_learned", "recommendation", "closing_summary"],
        "required": {"project_code"},
        "rows": [
            ["REV-ERP-001", "PRJ-ERP-001", "ERP extension remains on track.", "Timeline is aligned with planned finish.", "Budget usage is below approved value.", "Procurement review completed.", "Faster finance reporting cycle.", "Keep vendor milestone gates visible.", "Continue execution monitoring.", ""],
            ["REV-WH-001", "PRJ-WH-001", "Warehouse automation requires integration focus.", "PLC retest is in progress.", "Budget utilization is within tolerance.", "Retest and installation activities active.", "Throughput improvement expected.", "Retest windows should be planned earlier.", "Monitor vendor dependency weekly.", ""],
        ],
    },
}


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


def sheet_xml(name: str, config: dict) -> str:
    headers = config["headers"]
    required = config["required"]
    rows = []
    last_col = col_name(len(headers))
    rows.append(f'<row r="1">{cell("A1", f"{name} - Investment Backend Import Sheet", 1)}</row>')
    note_cells = [cell(f"{col_name(i)}2", "Required" if header in required else "Optional", 2) for i, header in enumerate(headers, start=1)]
    rows.append(f'<row r="2">{"".join(note_cells)}</row>')
    header_cells = [cell(f"{col_name(i)}3", header, 3) for i, header in enumerate(headers, start=1)]
    rows.append(f'<row r="3">{"".join(header_cells)}</row>')
    money_columns = {"budget_requested", "budget_approved", "actual_cost"}
    for row_idx, data in enumerate(config["rows"], start=4):
        cells = []
        for col_idx, value in enumerate(data, start=1):
            header = headers[col_idx - 1]
            style = 5 if header in money_columns else 0
            cells.append(cell(f"{col_name(col_idx)}{row_idx}", value, style))
        rows.append(f'<row r="{row_idx}">{"".join(cells)}</row>')

    validations = []
    if "status" in headers:
        col = col_name(headers.index("status") + 1)
        options = ACTIVITY_STATUS if name == "PROJECT_ACTIVITY" else PROJECT_STATUS
        validations.append(f'<dataValidation type="list" allowBlank="1" sqref="{col}4:{col}500"><formula1>"{options}"</formula1></dataValidation>')
    if "priority" in headers:
        col = col_name(headers.index("priority") + 1)
        validations.append(f'<dataValidation type="list" allowBlank="1" sqref="{col}4:{col}500"><formula1>"{PRIORITY}"</formula1></dataValidation>')
    if "progress" in headers:
        col = col_name(headers.index("progress") + 1)
        validations.append(f'<dataValidation type="whole" operator="between" allowBlank="1" sqref="{col}4:{col}500"><formula1>0</formula1><formula2>100</formula2></dataValidation>')
    validations_xml = f'<dataValidations count="{len(validations)}">{"".join(validations)}</dataValidations>' if validations else ""

    widths = ''.join(f'<col min="{i}" max="{i}" width="{max(12, min(34, len(header) + 4))}" customWidth="1"/>' for i, header in enumerate(headers, start=1))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<cols>{widths}</cols>
<sheetData>{"".join(rows)}</sheetData>
<mergeCells count="1"><mergeCell ref="A1:{last_col}1"/></mergeCells>
{validations_xml}
<autoFilter ref="A3:{last_col}{len(config["rows"]) + 3}"/>
</worksheet>'''


CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
''' + ''.join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, len(SHEETS) + 1)) + '</Types>'

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

WORKBOOK_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
''' + ''.join(f'<sheet name="{name}" sheetId="{i}" r:id="rId{i}"/>' for i, name in enumerate(SHEETS, start=1)) + '''
</sheets>
</workbook>'''

WORKBOOK_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
''' + ''.join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(SHEETS) + 1)) + f'''
<Relationship Id="rId{len(SHEETS) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><numFmt numFmtId="165" formatCode="&quot;Rp&quot; #,##0"/></numFmts>
<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0F172A"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2563EB"/></patternFill></fill></fills>
<borders count="2"><border/><border><left style="thin"><color rgb="FFCBD5E1"/></left><right style="thin"><color rgb="FFCBD5E1"/></right><top style="thin"><color rgb="FFCBD5E1"/></top><bottom style="thin"><color rgb="FFCBD5E1"/></bottom></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="6">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

APP = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Codex</Application></Properties>'''

CORE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Investment Backend Template</dc:title><dc:creator>Codex</dc:creator></cp:coreProperties>'''

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
    for i, (sheet_name, config) in enumerate(SHEETS.items(), start=1):
        zf.writestr(f"xl/worksheets/sheet{i}.xml", sheet_xml(sheet_name, config))

SCAFFOLD.write_bytes(OUTPUT.read_bytes())
print(OUTPUT)
