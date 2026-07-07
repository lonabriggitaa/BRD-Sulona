from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(r"C:\Users\LENOVO\Documents\Codex\2026-07-01\https-eldcm9-oss-github-io-brd")
SCRIPT = ROOT / "work" / "app-script.js"
OUTPUT = ROOT / "outputs" / "Investment_Backend_Detail_From_BRD_Sulona.xlsx"

LOOKUPS = {
    "Division": ["Engineering", "Finance", "Operations", "Supply Chain", "Quality", "Logistics", "Procurement", "Manufacturing", "Human Capital", "Commercial", "General Affairs"],
    "Site": ["Bandung", "Jakarta", "Surabaya", "Medan", "Semarang", "Makassar"],
    "Status": ["Draft", "Submitted", "Approved", "Running", "Completed", "Closed", "Cancelled"],
    "Document Type": ["BRD", "Financial Analysis", "Quotation", "Contract", "Progress Report", "LPJ", "Final Report", "Others"],
    "Priority": ["Low", "Medium", "High"],
    "Activity Status": ["Planned", "Running", "Completed", "Delayed", "Cancelled"],
}

PROJECT_HEADERS = [
    "BRD Number", "Project Name", "Division", "Site", "Vendor", "PIC", "Budget Requested",
    "Budget Approved", "Actual Cost", "Planned Start Date", "Target Finish Date",
    "Progress (%)", "Status", "Remarks",
]
ACTIVITY_HEADERS = [
    "Activity Date", "BRD Number", "Project Name", "Activity", "Description", "PIC",
    "Progress (%)", "Status", "Notes",
]
DOCUMENT_HEADERS = ["BRD Number", "Project Name", "Document Type", "Document Name", "Version", "Upload Date", "Remarks"]


def col_name(index: int) -> str:
    out = ""
    while index:
        index, rem = divmod(index - 1, 26)
        out = chr(65 + rem) + out
    return out


def excel_serial(value: date) -> int:
    return (value - date(1899, 12, 30)).days


def parse_date(value: str | None):
    if not value:
        return ""
    return datetime.strptime(value, "%Y-%m-%d").date()


def parse_objects(var_name: str):
    text = SCRIPT.read_text(encoding="utf-8")
    match_var = re.search(rf"(?:let|const)\s+{re.escape(var_name)}\s*=\s*\[", text)
    if not match_var:
        return []
    start = match_var.start()
    cursor = text.index("[", start)
    depth = 0
    end = cursor
    while end < len(text):
        char = text[end]
        if char == "[":
            depth += 1
        elif char == "]":
            depth -= 1
            if depth == 0:
                break
        end += 1
    array_text = text[cursor + 1:end]
    objects = []
    for match in re.finditer(r"\{([^{}]+)\}", array_text, flags=re.S):
        row = {}
        for key, raw in re.findall(r"(\w+):\s*(\"[^\"]*\"|[0-9]+)", match.group(1)):
            if raw.startswith('"'):
                row[key] = raw[1:-1]
            else:
                row[key] = int(raw)
        if row:
            objects.append(row)
    return objects


def normalize_project_status(status: str) -> str:
    return {
        "Under Review": "Submitted",
        "Overdue": "Running",
        "Procurement": "Approved",
    }.get(status, status if status in LOOKUPS["Status"] else "Draft")


def normalize_activity_status(project_status: str) -> str:
    if project_status == "Completed":
        return "Completed"
    if project_status == "Overdue":
        return "Delayed"
    if project_status == "Upcoming":
        return "Planned"
    if project_status in LOOKUPS["Activity Status"]:
        return project_status
    return "Running"


projects = parse_objects("projects")
activities = parse_objects("activities")
calendar_items = parse_objects("calendarItems")
project_by_name = {row["name"]: row for row in projects}

project_rows = []
for project in projects:
    status = normalize_project_status(project.get("status", "Draft"))
    remarks = project.get("remarks", "")
    if project.get("status") != status:
        remarks = f"Original dashboard status: {project.get('status')}. {remarks}".strip()
    project_rows.append([
        project.get("brd", ""),
        project.get("name", ""),
        project.get("division", ""),
        project.get("site", ""),
        project.get("vendor", ""),
        project.get("pic", ""),
        project.get("budget", 0),
        project.get("approved", 0),
        project.get("actual", 0),
        parse_date(project.get("start")),
        parse_date(project.get("finish")),
        project.get("progress", 0),
        status,
        remarks,
    ])

activity_rows = []
for activity in activities:
    project = project_by_name.get(activity.get("project", ""), {})
    activity_rows.append([
        parse_date(activity.get("date")),
        project.get("brd", ""),
        project.get("name", activity.get("project", "")),
        activity.get("type", "Activity"),
        activity.get("text", ""),
        project.get("pic", ""),
        project.get("progress", 0),
        normalize_activity_status(project.get("status", "Running")),
        "Generated from current HTML preview data.",
    ])

for item in calendar_items:
    project = project_by_name.get(item.get("project", ""), {})
    activity_rows.append([
        parse_date(item.get("date")),
        project.get("brd", ""),
        project.get("name", item.get("project", "")),
        item.get("title", item.get("type", "Activity")),
        item.get("note", ""),
        project.get("pic", ""),
        item.get("progress", project.get("progress", 0)),
        normalize_activity_status(item.get("status", project.get("status", "Running"))),
        f"Calendar type: {item.get('type', 'Activity')}. Source: https://lonabriggitaa.github.io/BRD-Sulona/",
    ])

document_rows = []
document_plan = [
    ("BRD", "BRD Rev 01", 0, "Initial business requirement document."),
    ("Financial Analysis", "Financial Analysis Rev 01", 2, "Budget and benefit analysis metadata."),
    ("Quotation", "Vendor Quotation", 5, "Vendor quotation reference."),
    ("Contract", "Contract Package", 12, "Contract and commercial agreement metadata."),
    ("Progress Report", "Monthly Progress Report", 30, "Latest progress reporting metadata."),
    ("LPJ", "LPJ Draft", 45, "Accountability report metadata."),
    ("Final Report", "Final Report", 60, "Closeout document metadata when project is complete."),
]
for project in projects:
    start = parse_date(project.get("start")) or date(2026, 1, 1)
    for doc_type, name, offset, note in document_plan:
        if doc_type in {"LPJ", "Final Report"} and normalize_project_status(project.get("status", "")) not in {"Completed", "Closed"}:
            continue
        upload_date = date.fromordinal(start.toordinal() + offset)
        document_rows.append([
            project.get("brd", ""),
            project.get("name", ""),
            doc_type,
            f"{name} - {project.get('name', '')}",
            "1.0",
            upload_date,
            note,
        ])

README_ROWS = [
    ["Investment Backend From HTML"],
    ["1. Data was extracted from the current BRD-Sulona HTML preview."],
    ["2. Use PROJECT_MASTER for imported project master data."],
    ["3. Use PROJECT_ACTIVITY for project activity history."],
    ["4. PROJECT_DOCUMENT is included for document metadata."],
    ["5. Do not change sheet names."],
    ["6. Do not change column names."],
    ["7. Upload this workbook in Import Update menu to synchronize the dashboard."],
    ["Source URL: https://lonabriggitaa.github.io/BRD-Sulona/"],
]


def cell(ref: str, value, style: int = 2) -> str:
    style_attr = f' s="{style}"' if style else ""
    if value is None or value == "":
        return f'<c r="{ref}"{style_attr}/>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"{style_attr}><v>{value}</v></c>'
    if isinstance(value, date):
        return f'<c r="{ref}" s="4"><v>{excel_serial(value)}</v></c>'
    return f'<c r="{ref}" t="inlineStr"{style_attr}><is><t>{escape(str(value))}</t></is></c>'


def validation_xml(sheet_name: str, headers: list[str]) -> str:
    lookup_col = {"Division": "A", "Site": "B", "Status": "C", "Document Type": "D", "Priority": "E", "Activity Status": "F"}
    targets = {}
    if sheet_name == "PROJECT_MASTER":
        targets = {"Division": "Division", "Site": "Site", "Status": "Status"}
    elif sheet_name == "PROJECT_ACTIVITY":
        targets = {"Status": "Activity Status"}
    elif sheet_name == "PROJECT_DOCUMENT":
        targets = {"Document Type": "Document Type"}
    items = []
    for header, lookup_name in targets.items():
        if header in headers:
            idx = headers.index(header) + 1
            col = col_name(idx)
            lookup_letter = lookup_col[lookup_name]
            end_row = len(LOOKUPS[lookup_name]) + 1
            items.append(
                f'<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="{col}2:{col}500">'
                f"<formula1>LOOKUP!${lookup_letter}$2:${lookup_letter}${end_row}</formula1></dataValidation>"
            )
    for progress_header in ["Progress (%)"]:
        if progress_header in headers:
            col = col_name(headers.index(progress_header) + 1)
            items.append(f'<dataValidation type="whole" operator="between" allowBlank="1" sqref="{col}2:{col}500"><formula1>0</formula1><formula2>100</formula2></dataValidation>')
    if not items:
        return ""
    return f'<dataValidations count="{len(items)}">{"".join(items)}</dataValidations>'


def conditional_xml(headers: list[str]) -> str:
    parts = []
    if "Progress (%)" in headers:
        col = col_name(headers.index("Progress (%)") + 1)
        parts.append(
            f'<conditionalFormatting sqref="{col}2:{col}500">'
            '<cfRule type="cellIs" priority="1" operator="between" dxfId="0"><formula>0</formula><formula>30</formula></cfRule>'
            '<cfRule type="cellIs" priority="2" operator="between" dxfId="1"><formula>31</formula><formula>70</formula></cfRule>'
            '<cfRule type="cellIs" priority="3" operator="between" dxfId="2"><formula>71</formula><formula>100</formula></cfRule>'
            '</conditionalFormatting>'
        )
    if "Status" in headers:
        col = col_name(headers.index("Status") + 1)
        parts.append(
            f'<conditionalFormatting sqref="{col}2:{col}500">'
            '<cfRule type="containsText" priority="4" operator="containsText" text="Running" dxfId="3"><formula>NOT(ISERROR(SEARCH("Running",'
            f'{col}2)))</formula></cfRule>'
            '<cfRule type="containsText" priority="5" operator="containsText" text="Completed" dxfId="2"><formula>NOT(ISERROR(SEARCH("Completed",'
            f'{col}2)))</formula></cfRule>'
            '<cfRule type="containsText" priority="6" operator="containsText" text="Delayed" dxfId="0"><formula>NOT(ISERROR(SEARCH("Delayed",'
            f'{col}2)))</formula></cfRule>'
            '<cfRule type="containsText" priority="7" operator="containsText" text="Cancelled" dxfId="0"><formula>NOT(ISERROR(SEARCH("Cancelled",'
            f'{col}2)))</formula></cfRule>'
            '</conditionalFormatting>'
        )
    return "".join(parts)


def sheet_xml(sheet_name: str, headers: list[str], rows_data: list[list]) -> str:
    rows = [f'<row r="1">{"".join(cell(f"{col_name(i)}1", h, 3) for i, h in enumerate(headers, 1))}</row>']
    money = {"Budget Requested", "Budget Approved", "Actual Cost"}
    for row_idx, row_values in enumerate(rows_data, 2):
        cells = []
        for col_idx, value in enumerate(row_values, 1):
            header = headers[col_idx - 1]
            style = 5 if header in money else 2
            cells.append(cell(f"{col_name(col_idx)}{row_idx}", value, style))
        rows.append(f'<row r="{row_idx}">{"".join(cells)}</row>')
    widths = "".join(f'<col min="{i}" max="{i}" width="{max(13, min(34, len(header) + 5))}" customWidth="1"/>' for i, header in enumerate(headers, 1))
    last_col = col_name(len(headers))
    last_row = max(2, len(rows_data) + 1)
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<cols>{widths}</cols><sheetData>{"".join(rows)}</sheetData>
{validation_xml(sheet_name, headers)}
{conditional_xml(headers)}
<autoFilter ref="A1:{last_col}{last_row}"/>
</worksheet>'''


def lookup_xml() -> str:
    headers = list(LOOKUPS)
    max_len = max(len(v) for v in LOOKUPS.values())
    rows = [f'<row r="1">{"".join(cell(f"{col_name(i)}1", h, 3) for i, h in enumerate(headers, 1))}</row>']
    for row_idx in range(2, max_len + 2):
        rows.append(f'<row r="{row_idx}">{"".join(cell(f"{col_name(i)}{row_idx}", LOOKUPS[h][row_idx - 2] if row_idx - 2 < len(LOOKUPS[h]) else "") for i, h in enumerate(headers, 1))}</row>')
    widths = "".join(f'<col min="{i}" max="{i}" width="22" customWidth="1"/>' for i in range(1, len(headers) + 1))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>{widths}</cols><sheetData>{"".join(rows)}</sheetData><autoFilter ref="A1:{col_name(len(headers))}{max_len + 1}"/></worksheet>'''


def readme_xml() -> str:
    rows = [f'<row r="{i}">{cell(f"A{i}", row[0], 1 if i == 1 else 2)}</row>' for i, row in enumerate(README_ROWS, 1)]
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews><cols><col min="1" max="1" width="96" customWidth="1"/></cols><sheetData>{"".join(rows)}</sheetData></worksheet>'''


SHEET_NAMES = ["PROJECT_MASTER", "PROJECT_ACTIVITY", "PROJECT_DOCUMENT", "LOOKUP", "README"]
CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>''' + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, 6)) + "</Types>"
ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'''
WORKBOOK = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>''' + "".join(f'<sheet name="{name}" sheetId="{i}" r:id="rId{i}"/>' for i, name in enumerate(SHEET_NAMES, 1)) + "</sheets></workbook>"
WORKBOOK_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">''' + "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, 6)) + '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/><numFmt numFmtId="165" formatCode="&quot;Rp&quot; #,##0"/></numFmts><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0F172A"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2563EB"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"><color rgb="FFCBD5E1"/></left><right style="thin"><color rgb="FFCBD5E1"/></right><top style="thin"><color rgb="FFCBD5E1"/></top><bottom style="thin"><color rgb="FFCBD5E1"/></bottom></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/></cellXfs><dxfs count="4"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFEE2E2"/></patternFill></fill><font><color rgb="FF991B1B"/></font></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFEF3C7"/></patternFill></fill><font><color rgb="FF92400E"/></font></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFDCFCE7"/></patternFill></fill><font><color rgb="FF166534"/></font></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFDBEAFE"/></patternFill></fill><font><color rgb="FF1D4ED8"/></font></dxf></dxfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'''
APP = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Codex</Application></Properties>'''
CORE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Investment Backend From HTML</dc:title><dc:creator>Codex</dc:creator></cp:coreProperties>'''

OUTPUT.parent.mkdir(parents=True, exist_ok=True)
with ZipFile(OUTPUT, "w", ZIP_DEFLATED) as zf:
    zf.writestr("[Content_Types].xml", CONTENT_TYPES)
    zf.writestr("_rels/.rels", ROOT_RELS)
    zf.writestr("xl/workbook.xml", WORKBOOK)
    zf.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
    zf.writestr("xl/styles.xml", STYLES)
    zf.writestr("docProps/app.xml", APP)
    zf.writestr("docProps/core.xml", CORE)
    zf.writestr("xl/worksheets/sheet1.xml", sheet_xml("PROJECT_MASTER", PROJECT_HEADERS, project_rows))
    zf.writestr("xl/worksheets/sheet2.xml", sheet_xml("PROJECT_ACTIVITY", ACTIVITY_HEADERS, activity_rows))
    zf.writestr("xl/worksheets/sheet3.xml", sheet_xml("PROJECT_DOCUMENT", DOCUMENT_HEADERS, document_rows))
    zf.writestr("xl/worksheets/sheet4.xml", lookup_xml())
    zf.writestr("xl/worksheets/sheet5.xml", readme_xml())

print(OUTPUT)
