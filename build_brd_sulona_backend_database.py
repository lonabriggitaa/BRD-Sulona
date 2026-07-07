from __future__ import annotations

import json
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(r"C:\Users\LENOVO\Documents\Codex\2026-07-01\https-eldcm9-oss-github-io-brd")
APP_SCRIPT = ROOT / "work" / "app-script.js"
OUTPUT = ROOT / "outputs" / "BRD_Sulona_Backend_Database.xlsx"

MAX_PROJECT_ROWS = 500
MAX_ACTIVITY_ROWS = 1000
MAX_DOCUMENT_ROWS = 1000
MAX_REVIEW_ROWS = 500

SHEETS = [
    "DASHBOARD",
    "PROJECT_MASTER",
    "PROJECT_ACTIVITY",
    "PROJECT_DOCUMENT",
    "PROJECT_REVIEW",
    "LOOKUP",
    "DASHBOARD_SUMMARY",
    "CHART_DATA",
    "README",
]

PROJECT_HEADERS = [
    "BRD Number", "Project Name", "Division", "Site", "Category", "Vendor", "PIC",
    "Budget Requested", "Budget Approved", "Actual Cost", "Budget Variance",
    "Budget Utilization (%)", "Planned Start Date", "Target Finish Date",
    "Actual Finish Date", "Project Duration", "Remaining Days", "Delay Days",
    "Progress (%)", "Status", "Project Health", "Remarks",
]
ACTIVITY_HEADERS = [
    "Activity Date", "BRD Number", "Project Name", "Activity Type", "Activity Title",
    "Description", "PIC", "Progress (%)", "Status", "Notes",
]
DOCUMENT_HEADERS = [
    "BRD Number", "Project Name", "Document Type", "Document Name", "Version",
    "Upload Date", "Status", "Remarks",
]
REVIEW_HEADERS = [
    "BRD Number", "Project Name", "Executive Summary", "Timeline Performance",
    "Budget Performance", "Activity Summary", "Benefit Realization", "Lessons Learned",
    "Recommendation", "Closing Summary", "Review Status",
]

LOOKUPS = {
    "Division": ["Engineering", "Production", "IT", "GA", "Finance", "QA", "Warehouse", "Manufacture", "Corporate Secretary", "Creative & Innovation"],
    "Site": ["Cirebon", "Bandung", "Jakarta", "Yogyakarta", "Cilacap"],
    "Category": ["Machine", "Building", "IT System", "Vehicle", "Utility", "Warehouse", "Other", "Expansion Investment", "Replacement Investment", "Non-Measurable Profit Investment"],
    "Project Status": ["Draft", "Submitted", "Approved", "Running", "Completed", "Closed", "Cancelled"],
    "Priority": ["Low", "Medium", "High"],
    "Project Health": ["Green", "Yellow", "Red"],
    "Document Type": ["BRD", "Financial Analysis", "Quotation", "Contract", "Purchase Order", "Progress Report", "LPJ", "Final Report", "Others"],
    "Activity Type": ["BRD Submitted", "Investment Review", "Vendor Discussion", "Vendor Selection", "Budget Approval", "Procurement", "Material Delivery", "Installation", "Testing", "Go Live", "Monitoring", "Final Review", "Project Completed"],
    "Activity Status": ["Planned", "Running", "Completed", "Delayed", "Cancelled"],
    "Document Status": ["Not Available", "Draft", "Uploaded", "Approved", "Revised"],
    "Review Status": ["Not Started", "Draft", "Completed"],
}


def col_name(index: int) -> str:
    out = ""
    while index:
        index, rem = divmod(index - 1, 26)
        out = chr(65 + rem) + out
    return out


def excel_serial(value: date) -> int:
    return (value - date(1899, 12, 30)).days


def parse_iso_date(value: str | None) -> date | str:
    if not value:
        return ""
    try:
        return datetime.strptime(value[:10], "%Y-%m-%d").date()
    except ValueError:
        return value


def number(value, default=0):
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


def parse_array(name: str):
    text = APP_SCRIPT.read_text(encoding="utf-8")
    match = re.search(rf"(?:let|const)\s+{re.escape(name)}\s*=\s*\[", text)
    if not match:
        return []
    start = text.index("[", match.start())
    depth = 0
    end = start
    while end < len(text):
        if text[end] == "[":
            depth += 1
        elif text[end] == "]":
            depth -= 1
            if depth == 0:
                end += 1
                break
        end += 1
    return json.loads(text[start:end])


projects = parse_array("projects")
activities = parse_array("activities")
project_by_brd = {p.get("brd"): p for p in projects}
project_by_name = {p.get("name"): p for p in projects}


def project_status(status: str) -> str:
    if status in {"Completed", "Closed"}:
        return "Completed"
    if status == "Submitted":
        return "Submitted"
    if status == "Procurement":
        return "Approved"
    if status == "Overdue":
        return "Running"
    return status if status in LOOKUPS["Project Status"] else "Running"


def activity_status(status: str) -> str:
    if status == "Upcoming":
        return "Planned"
    if status == "Overdue":
        return "Delayed"
    return status if status in LOOKUPS["Activity Status"] else "Running"


project_rows = []
for p in projects:
    actual_finish = parse_iso_date(p.get("finish")) if project_status(p.get("status", "")) == "Completed" else ""
    project_rows.append([
        p.get("brd", ""),
        p.get("name", ""),
        p.get("division", ""),
        p.get("site", ""),
        p.get("category", ""),
        p.get("vendor", ""),
        p.get("pic", ""),
        number(p.get("budget")),
        number(p.get("approved")),
        number(p.get("actual")),
        None,
        None,
        parse_iso_date(p.get("start")),
        parse_iso_date(p.get("finish")),
        actual_finish,
        None,
        None,
        None,
        number(p.get("progress")),
        project_status(p.get("status", "")),
        None,
        p.get("remarks", ""),
    ])

activity_rows = []
for a in activities:
    p = project_by_brd.get(a.get("brd")) or project_by_name.get(a.get("project")) or {}
    activity_rows.append([
        parse_iso_date(a.get("date")),
        p.get("brd") or a.get("brd", ""),
        None,
        a.get("type", "Monitoring"),
        a.get("title") or a.get("type", "Activity"),
        a.get("text") or a.get("note", ""),
        a.get("pic") or p.get("pic", ""),
        number(a.get("progress"), number(p.get("progress"))),
        activity_status(a.get("status", "")),
        a.get("note") or "",
    ])

document_rows = []
doc_seed = [
    ("BRD", "BRD Rev 01", "Approved", 0),
    ("Financial Analysis", "Financial Analysis", "Uploaded", 2),
    ("Quotation", "Vendor Quotation", "Uploaded", 5),
    ("Contract", "Contract Package", "Draft", 12),
    ("Purchase Order", "Purchase Order", "Uploaded", 14),
    ("Progress Report", "Monthly Progress Report", "Uploaded", 30),
]
for p in projects:
    start = parse_iso_date(p.get("start"))
    if not isinstance(start, date):
        start = date(2026, 1, 1)
    rows_for_project = list(doc_seed)
    if project_status(p.get("status", "")) == "Completed":
        rows_for_project.extend([("LPJ", "LPJ Final", "Approved", 45), ("Final Report", "Final Report", "Approved", 60)])
    for doc_type, name, status, offset in rows_for_project:
        document_rows.append([
            p.get("brd", ""),
            None,
            doc_type,
            f"{name} - {p.get('name', '')}",
            "1.0",
            start + timedelta(days=offset),
            status,
            "Metadata dokumen untuk dashboard BRD Sulona.",
        ])

review_rows = []
for p in projects:
    completed = project_status(p.get("status", "")) == "Completed"
    review_rows.append([
        p.get("brd", ""),
        None,
        f"Ringkasan eksekutif untuk {p.get('name', '')}." if completed else "",
        "Sesuai target." if completed else "",
        "Terkendali terhadap budget." if completed else "",
        "Aktivitas proyek tercatat pada timeline." if completed else "",
        "Benefit akan dimonitor setelah implementasi." if completed else "",
        "Dokumentasi dan koordinasi perlu dijaga." if completed else "",
        "Tutup proyek dan lanjutkan monitoring benefit." if completed else "",
        "Project completed and ready for review archive." if completed else "",
        "Completed" if completed else "Not Started",
    ])


def cell(ref: str, value=None, style: int = 2, formula: str | None = None) -> str:
    style_attr = f' s="{style}"' if style else ""
    if formula is not None:
        return f'<c r="{ref}"{style_attr}><f>{escape(formula)}</f></c>'
    if value is None or value == "":
        return f'<c r="{ref}"{style_attr}/>'
    if isinstance(value, bool):
        return f'<c r="{ref}" t="b"{style_attr}><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"{style_attr}><v>{value}</v></c>'
    if isinstance(value, date):
        return f'<c r="{ref}" s="4"><v>{excel_serial(value)}</v></c>'
    return f'<c r="{ref}" t="inlineStr"{style_attr}><is><t>{escape(str(value))}</t></is></c>'


def row_xml(row_idx: int, values: list, headers: list[str], formulas: dict[int, str] | None = None, styles: dict[int, int] | None = None) -> str:
    formulas = formulas or {}
    styles = styles or {}
    cells = []
    for col_idx, value in enumerate(values, 1):
        header = headers[col_idx - 1] if col_idx - 1 < len(headers) else ""
        style = styles.get(col_idx)
        if style is None:
            if "Budget" in header or header == "Actual Cost":
                style = 5
            elif "Date" in header:
                style = 4
            elif "Utilization" in header:
                style = 6
            else:
                style = 2
        cells.append(cell(f"{col_name(col_idx)}{row_idx}", value, style=style, formula=formulas.get(col_idx)))
    return f'<row r="{row_idx}">{"".join(cells)}</row>'


def validations(sheet: str, headers: list[str], max_row: int) -> str:
    lookup_columns = {name: col_name(i) for i, name in enumerate(LOOKUPS, 1)}
    targets: dict[str, str] = {}
    if sheet == "PROJECT_MASTER":
        targets = {"Division": "Division", "Site": "Site", "Category": "Category", "Status": "Project Status"}
    elif sheet == "PROJECT_ACTIVITY":
        targets = {"Activity Type": "Activity Type", "Status": "Activity Status"}
    elif sheet == "PROJECT_DOCUMENT":
        targets = {"Document Type": "Document Type", "Status": "Document Status"}
    elif sheet == "PROJECT_REVIEW":
        targets = {"Review Status": "Review Status"}
    items = []
    for header, lookup_name in targets.items():
        if header in headers:
            col = col_name(headers.index(header) + 1)
            lookup_col = lookup_columns[lookup_name]
            last = len(LOOKUPS[lookup_name]) + 1
            items.append(f'<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="{col}2:{col}{max_row}"><formula1>LOOKUP!${lookup_col}$2:${lookup_col}${last}</formula1></dataValidation>')
    for header in ("Progress (%)",):
        if header in headers:
            col = col_name(headers.index(header) + 1)
            items.append(f'<dataValidation type="whole" operator="between" allowBlank="1" showErrorMessage="1" sqref="{col}2:{col}{max_row}"><formula1>0</formula1><formula2>100</formula2></dataValidation>')
    if not items:
        return ""
    return f'<dataValidations count="{len(items)}">{"".join(items)}</dataValidations>'


def cf_rules(sheet: str, headers: list[str], max_row: int) -> str:
    parts = []
    if "Progress (%)" in headers:
        col = col_name(headers.index("Progress (%)") + 1)
        parts.append(f'<conditionalFormatting sqref="{col}2:{col}{max_row}"><cfRule type="cellIs" priority="1" operator="between" dxfId="0"><formula>0</formula><formula>30</formula></cfRule><cfRule type="cellIs" priority="2" operator="between" dxfId="1"><formula>31</formula><formula>70</formula></cfRule><cfRule type="cellIs" priority="3" operator="between" dxfId="2"><formula>71</formula><formula>100</formula></cfRule></conditionalFormatting>')
    if "Status" in headers:
        col = col_name(headers.index("Status") + 1)
        parts.append(f'<conditionalFormatting sqref="{col}2:{col}{max_row}"><cfRule type="containsText" priority="4" operator="containsText" text="Running" dxfId="3"><formula>NOT(ISERROR(SEARCH("Running",{col}2)))</formula></cfRule><cfRule type="containsText" priority="5" operator="containsText" text="Completed" dxfId="2"><formula>NOT(ISERROR(SEARCH("Completed",{col}2)))</formula></cfRule><cfRule type="containsText" priority="6" operator="containsText" text="Delayed" dxfId="0"><formula>NOT(ISERROR(SEARCH("Delayed",{col}2)))</formula></cfRule><cfRule type="containsText" priority="7" operator="containsText" text="Cancelled" dxfId="0"><formula>NOT(ISERROR(SEARCH("Cancelled",{col}2)))</formula></cfRule><cfRule type="containsText" priority="8" operator="containsText" text="Draft" dxfId="4"><formula>NOT(ISERROR(SEARCH("Draft",{col}2)))</formula></cfRule><cfRule type="containsText" priority="9" operator="containsText" text="Submitted" dxfId="4"><formula>NOT(ISERROR(SEARCH("Submitted",{col}2)))</formula></cfRule><cfRule type="containsText" priority="10" operator="containsText" text="Approved" dxfId="4"><formula>NOT(ISERROR(SEARCH("Approved",{col}2)))</formula></cfRule></conditionalFormatting>')
    if "Project Health" in headers:
        col = col_name(headers.index("Project Health") + 1)
        parts.append(f'<conditionalFormatting sqref="{col}2:{col}{max_row}"><cfRule type="containsText" priority="11" operator="containsText" text="Green" dxfId="2"><formula>NOT(ISERROR(SEARCH("Green",{col}2)))</formula></cfRule><cfRule type="containsText" priority="12" operator="containsText" text="Yellow" dxfId="1"><formula>NOT(ISERROR(SEARCH("Yellow",{col}2)))</formula></cfRule><cfRule type="containsText" priority="13" operator="containsText" text="Red" dxfId="0"><formula>NOT(ISERROR(SEARCH("Red",{col}2)))</formula></cfRule></conditionalFormatting>')
    if "Budget Utilization (%)" in headers:
        col = col_name(headers.index("Budget Utilization (%)") + 1)
        parts.append(f'<conditionalFormatting sqref="{col}2:{col}{max_row}"><cfRule type="cellIs" priority="14" operator="greaterThan" dxfId="0"><formula>1</formula></cfRule><cfRule type="cellIs" priority="15" operator="between" dxfId="1"><formula>0.8</formula><formula>1</formula></cfRule><cfRule type="cellIs" priority="16" operator="lessThan" dxfId="2"><formula>0.8</formula></cfRule></conditionalFormatting>')
    if "Remaining Days" in headers:
        col = col_name(headers.index("Remaining Days") + 1)
        parts.append(f'<conditionalFormatting sqref="{col}2:{col}{max_row}"><cfRule type="cellIs" priority="17" operator="lessThanOrEqual" dxfId="0"><formula>7</formula></cfRule><cfRule type="cellIs" priority="18" operator="between" dxfId="1"><formula>8</formula><formula>14</formula></cfRule><cfRule type="cellIs" priority="19" operator="greaterThan" dxfId="2"><formula>14</formula></cfRule></conditionalFormatting>')
    return "".join(parts)


def table_xml(table_id: int, name: str, ref: str, headers: list[str]) -> str:
    cols = "".join(f'<tableColumn id="{i}" name="{escape(header)}"/>' for i, header in enumerate(headers, 1))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="{table_id}" name="{name}" displayName="{name}" ref="{ref}" totalsRowShown="0"><autoFilter ref="{ref}"/><tableColumns count="{len(headers)}">{cols}</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>'''


def worksheet_xml(sheet_name: str, headers: list[str], rows_data: list[list], max_row: int, formula_builder=None, table_id: int | None = None) -> str:
    header = f'<row r="1">{"".join(cell(f"{col_name(i)}1", h, style=3) for i, h in enumerate(headers, 1))}</row>'
    rows = [header]
    for idx in range(2, max_row + 1):
        data_idx = idx - 2
        values = rows_data[data_idx] if data_idx < len(rows_data) else [""] * len(headers)
        formulas = formula_builder(idx) if formula_builder else {}
        rows.append(row_xml(idx, values, headers, formulas=formulas))
    widths = "".join(f'<col min="{i}" max="{i}" width="{max(12, min(28, len(h) + 4))}" customWidth="1"/>' for i, h in enumerate(headers, 1))
    last_col = col_name(len(headers))
    table_part = '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>' if table_id else ""
    rel_ns = ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' if table_id else ""
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"{rel_ns}><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>{widths}</cols><sheetData>{"".join(rows)}</sheetData>{validations(sheet_name, headers, max_row)}{cf_rules(sheet_name, headers, max_row)}<autoFilter ref="A1:{last_col}{max_row}"/>{table_part}</worksheet>'''


def project_formulas(r: int) -> dict[int, str]:
    return {
        11: f'IFERROR(I{r}-J{r},0)',
        12: f'IFERROR(J{r}/I{r},0)',
        16: f'IFERROR(N{r}-M{r},0)',
        17: f'IF(N{r}="","",N{r}-TODAY())',
        18: f'IF(AND(T{r}<>"Completed",TODAY()>N{r}),TODAY()-N{r},0)',
        21: f'IF(OR(R{r}>0,L{r}>1),"Red",IF(OR(AND(S{r}>=40,S{r}<=79),Q{r}<=14),"Yellow",IF(AND(S{r}>=80,R{r}=0),"Green","Yellow")))',
    }


def project_name_formula(r: int, brd_col: str = "B") -> str:
    return f'IFERROR(XLOOKUP({brd_col}{r},PROJECT_MASTER!$A$2:$A${MAX_PROJECT_ROWS},PROJECT_MASTER!$B$2:$B${MAX_PROJECT_ROWS}),"")'


def lookup_sheet_xml() -> str:
    headers = list(LOOKUPS)
    max_len = max(len(v) for v in LOOKUPS.values())
    rows = [f'<row r="1">{"".join(cell(f"{col_name(i)}1", h, style=3) for i, h in enumerate(headers, 1))}</row>']
    for r in range(2, max_len + 2):
        values = [LOOKUPS[h][r - 2] if r - 2 < len(LOOKUPS[h]) else "" for h in headers]
        rows.append(row_xml(r, values, headers))
    widths = "".join(f'<col min="{i}" max="{i}" width="26" customWidth="1"/>' for i in range(1, len(headers) + 1))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>{widths}</cols><sheetData>{"".join(rows)}</sheetData><autoFilter ref="A1:{col_name(len(headers))}{max_len + 1}"/></worksheet>'''


def dashboard_summary_xml() -> str:
    rows = [["Metric", "Value", "Description"]]
    metrics = [
        ("Total Investment", f"SUM(PROJECT_MASTER!I2:I{MAX_PROJECT_ROWS})", "Total Budget Approved"),
        ("Total Actual Cost", f"SUM(PROJECT_MASTER!J2:J{MAX_PROJECT_ROWS})", "Total Actual Cost"),
        ("Total Projects", f"COUNTA(PROJECT_MASTER!B2:B{MAX_PROJECT_ROWS})", "Jumlah proyek"),
        ("Running Projects", f'COUNTIF(PROJECT_MASTER!T2:T{MAX_PROJECT_ROWS},"Running")', "Status Running"),
        ("Completed Projects", f'COUNTIF(PROJECT_MASTER!T2:T{MAX_PROJECT_ROWS},"Completed")', "Status Completed"),
        ("Pending Projects", f'COUNTIF(PROJECT_MASTER!T2:T{MAX_PROJECT_ROWS},"Draft")+COUNTIF(PROJECT_MASTER!T2:T{MAX_PROJECT_ROWS},"Submitted")+COUNTIF(PROJECT_MASTER!T2:T{MAX_PROJECT_ROWS},"Approved")', "Draft / Submitted / Approved"),
        ("Overdue Projects", f'COUNTIF(PROJECT_MASTER!R2:R{MAX_PROJECT_ROWS},">0")', "Delay Days > 0"),
        ("Average Progress", f'AVERAGEIF(PROJECT_MASTER!B2:B{MAX_PROJECT_ROWS},"<>",PROJECT_MASTER!S2:S{MAX_PROJECT_ROWS})', "Rata-rata progress"),
        ("Budget Utilization", "IFERROR(B3/B2,0)", "Actual Cost / Investment"),
        ("Healthy Projects", f'COUNTIF(PROJECT_MASTER!U2:U{MAX_PROJECT_ROWS},"Green")', "Health Green"),
        ("Need Attention Projects", f'COUNTIF(PROJECT_MASTER!U2:U{MAX_PROJECT_ROWS},"Yellow")', "Health Yellow"),
        ("Critical Projects", f'COUNTIF(PROJECT_MASTER!U2:U{MAX_PROJECT_ROWS},"Red")', "Health Red"),
        ("Total Activities", f"COUNTA(PROJECT_ACTIVITY!E2:E{MAX_ACTIVITY_ROWS})", "Jumlah aktivitas"),
        ("Total Documents", f"COUNTA(PROJECT_DOCUMENT!D2:D{MAX_DOCUMENT_ROWS})", "Jumlah dokumen"),
        ("Projects Without Recent Activity", f'SUMPRODUCT(--(PROJECT_MASTER!A2:A{MAX_PROJECT_ROWS}<>""),--(COUNTIFS(PROJECT_ACTIVITY!B2:B{MAX_ACTIVITY_ROWS},PROJECT_MASTER!A2:A{MAX_PROJECT_ROWS},PROJECT_ACTIVITY!A2:A{MAX_ACTIVITY_ROWS},">="&TODAY()-7)=0))', "Tidak ada aktivitas 7 hari terakhir"),
    ]
    sheet_rows = [f'<row r="1">{"".join(cell(f"{col_name(i)}1", h, style=3) for i, h in enumerate(rows[0], 1))}</row>']
    for r, (metric, formula, desc) in enumerate(metrics, 2):
        sheet_rows.append(f'<row r="{r}">{cell(f"A{r}", metric)}{cell(f"B{r}", style=6 if metric in {"Average Progress", "Budget Utilization"} else 2, formula=formula)}{cell(f"C{r}", desc)}</row>')
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols><col min="1" max="1" width="34" customWidth="1"/><col min="2" max="2" width="20" customWidth="1"/><col min="3" max="3" width="52" customWidth="1"/></cols><sheetData>{"".join(sheet_rows)}</sheetData><autoFilter ref="A1:C{len(metrics)+1}"/></worksheet>'''


def dashboard_xml() -> str:
    rows: list[str] = []

    def add(row_num: int, cells: list[tuple[str, object | None, int, str | None]]) -> None:
        rows.append(f'<row r="{row_num}">{"".join(cell(f"{col}{row_num}", value, style=style, formula=formula) for col, value, style, formula in cells)}</row>')

    add(1, [("A", "BRD SULONA INVESTMENT DASHBOARD", 1, None)])
    add(2, [("A", "Ringkasan Excel backend untuk dashboard investasi, portfolio, aktivitas, dan dokumen.", 2, None)])
    add(4, [("A", "Total Investment", 3, None), ("C", "Total Actual Cost", 3, None), ("E", "Total Projects", 3, None), ("G", "Average Progress", 3, None), ("I", "Budget Utilization", 3, None)])
    add(5, [
        ("A", None, 5, "DASHBOARD_SUMMARY!B2"),
        ("C", None, 5, "DASHBOARD_SUMMARY!B3"),
        ("E", None, 2, "DASHBOARD_SUMMARY!B4"),
        ("G", None, 6, "DASHBOARD_SUMMARY!B9"),
        ("I", None, 6, "DASHBOARD_SUMMARY!B10"),
    ])
    add(7, [("A", "Running Projects", 3, None), ("C", "Completed Projects", 3, None), ("E", "Overdue Projects", 3, None), ("G", "Total Activities", 3, None), ("I", "Total Documents", 3, None)])
    add(8, [
        ("A", None, 2, "DASHBOARD_SUMMARY!B5"),
        ("C", None, 2, "DASHBOARD_SUMMARY!B6"),
        ("E", None, 2, "DASHBOARD_SUMMARY!B8"),
        ("G", None, 2, "DASHBOARD_SUMMARY!B14"),
        ("I", None, 2, "DASHBOARD_SUMMARY!B15"),
    ])

    add(11, [("A", "Portfolio Health", 1, None), ("D", "Project Status", 1, None), ("G", "Top Projects by Budget", 1, None)])
    add(12, [("A", "Health", 3, None), ("B", "Projects", 3, None), ("D", "Status", 3, None), ("E", "Projects", 3, None), ("G", "BRD Number", 3, None), ("H", "Project Name", 3, None), ("I", "Budget Approved", 3, None), ("J", "Actual Cost", 3, None), ("K", "Progress", 3, None), ("L", "Health", 3, None)])
    health_refs = [("Green", "DASHBOARD_SUMMARY!B11"), ("Yellow", "DASHBOARD_SUMMARY!B12"), ("Red", "DASHBOARD_SUMMARY!B13")]
    statuses = LOOKUPS["Project Status"]
    for idx in range(10):
        r = 13 + idx
        cells = []
        if idx < len(health_refs):
            label, ref = health_refs[idx]
            cells.extend([("A", label, 2, None), ("B", None, 2, ref)])
        if idx < len(statuses):
            status = statuses[idx]
            cells.extend([("D", status, 2, None), ("E", None, 2, f'COUNTIF(PROJECT_MASTER!T2:T{MAX_PROJECT_ROWS},"{status}")')])
        source_row = 2 + idx
        cells.extend([
            ("G", None, 2, f'PROJECT_MASTER!A{source_row}'),
            ("H", None, 2, f'PROJECT_MASTER!B{source_row}'),
            ("I", None, 5, f'PROJECT_MASTER!I{source_row}'),
            ("J", None, 5, f'PROJECT_MASTER!J{source_row}'),
            ("K", None, 6, f'PROJECT_MASTER!S{source_row}/100'),
            ("L", None, 2, f'PROJECT_MASTER!U{source_row}'),
        ])
        add(r, cells)

    add(25, [("A", "Budget vs Actual by Month", 1, None), ("G", "Latest Activity Timeline", 1, None)])
    add(26, [("A", "Month", 3, None), ("B", "Investment", 3, None), ("C", "Actual Cost", 3, None), ("D", "Completed", 3, None), ("E", "Activities", 3, None), ("G", "Date", 3, None), ("H", "BRD Number", 3, None), ("I", "Project Name", 3, None), ("J", "Activity", 3, None), ("K", "Status", 3, None)])
    for idx in range(12):
        r = 27 + idx
        chart_row = 3 + idx
        activity_row = 2 + idx
        add(r, [
            ("A", None, 2, f'CHART_DATA!A{chart_row}'),
            ("B", None, 5, f'CHART_DATA!B{chart_row}'),
            ("C", None, 5, f'CHART_DATA!C{chart_row}'),
            ("D", None, 2, f'CHART_DATA!D{chart_row}'),
            ("E", None, 2, f'CHART_DATA!E{chart_row}'),
            ("G", None, 4, f'PROJECT_ACTIVITY!A{activity_row}'),
            ("H", None, 2, f'PROJECT_ACTIVITY!B{activity_row}'),
            ("I", None, 2, f'PROJECT_ACTIVITY!C{activity_row}'),
            ("J", None, 2, f'PROJECT_ACTIVITY!E{activity_row}'),
            ("K", None, 2, f'PROJECT_ACTIVITY!I{activity_row}'),
        ])

    add(41, [("A", "Investment by Division", 1, None), ("D", "Investment by Category", 1, None), ("G", "Investment by Site", 1, None)])
    add(42, [("A", "Division", 3, None), ("B", "Investment", 3, None), ("D", "Category", 3, None), ("E", "Investment", 3, None), ("G", "Site", 3, None), ("H", "Investment", 3, None)])
    for idx in range(10):
        r = 43 + idx
        division = LOOKUPS["Division"][idx] if idx < len(LOOKUPS["Division"]) else ""
        category = LOOKUPS["Category"][idx] if idx < len(LOOKUPS["Category"]) else ""
        site = LOOKUPS["Site"][idx] if idx < len(LOOKUPS["Site"]) else ""
        cells = []
        if division:
            cells.extend([("A", division, 2, None), ("B", None, 5, f'SUMIF(PROJECT_MASTER!C2:C{MAX_PROJECT_ROWS},"{division}",PROJECT_MASTER!I2:I{MAX_PROJECT_ROWS})')])
        if category:
            cells.extend([("D", category, 2, None), ("E", None, 5, f'SUMIF(PROJECT_MASTER!E2:E{MAX_PROJECT_ROWS},"{category}",PROJECT_MASTER!I2:I{MAX_PROJECT_ROWS})')])
        if site:
            cells.extend([("G", site, 2, None), ("H", None, 5, f'SUMIF(PROJECT_MASTER!D2:D{MAX_PROJECT_ROWS},"{site}",PROJECT_MASTER!I2:I{MAX_PROJECT_ROWS})')])
        add(r, cells)

    widths = "".join(
        f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>'
        for i, w in enumerate([20, 18, 20, 20, 18, 4, 18, 28, 20, 20, 16, 16], 1)
    )
    cf = (
        '<conditionalFormatting sqref="A13:A15 L13:L22">'
        '<cfRule type="containsText" priority="1" operator="containsText" text="Green" dxfId="2"><formula>NOT(ISERROR(SEARCH("Green",A13)))</formula></cfRule>'
        '<cfRule type="containsText" priority="2" operator="containsText" text="Yellow" dxfId="1"><formula>NOT(ISERROR(SEARCH("Yellow",A13)))</formula></cfRule>'
        '<cfRule type="containsText" priority="3" operator="containsText" text="Red" dxfId="0"><formula>NOT(ISERROR(SEARCH("Red",A13)))</formula></cfRule>'
        '</conditionalFormatting>'
        '<conditionalFormatting sqref="K13:K22"><cfRule type="cellIs" priority="4" operator="lessThanOrEqual" dxfId="0"><formula>0.3</formula></cfRule><cfRule type="cellIs" priority="5" operator="between" dxfId="1"><formula>0.31</formula><formula>0.7</formula></cfRule><cfRule type="cellIs" priority="6" operator="greaterThan" dxfId="2"><formula>0.7</formula></cfRule></conditionalFormatting>'
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>{widths}</cols><sheetData>{"".join(rows)}</sheetData>{cf}</worksheet>'''


def chart_data_xml() -> str:
    rows = []
    current_row = 1

    def add_title(title: str):
        nonlocal current_row
        rows.append(f'<row r="{current_row}">{cell(f"A{current_row}", title, style=1)}</row>')
        current_row += 1

    def add_table(headers: list[str], data_rows: list[list[tuple[str | None, object | None, int]]]):
        nonlocal current_row
        rows.append(f'<row r="{current_row}">{"".join(cell(f"{col_name(i)}{current_row}", h, style=3) for i, h in enumerate(headers, 1))}</row>')
        current_row += 1
        for data in data_rows:
            cells = []
            for c, (formula, value, style) in enumerate(data, 1):
                cells.append(cell(f"{col_name(c)}{current_row}", value, style=style, formula=formula))
            rows.append(f'<row r="{current_row}">{"".join(cells)}</row>')
            current_row += 1
        current_row += 1

    month_rows = []
    for m in range(1, 13):
        label = date(2026, m, 1).strftime("%b 2026")
        start = f"DATE(2026,{m},1)"
        end = f"EDATE(DATE(2026,{m},1),1)"
        month_rows.append([
            (None, label, 2),
            (f'SUMIFS(PROJECT_MASTER!I2:I{MAX_PROJECT_ROWS},PROJECT_MASTER!M2:M{MAX_PROJECT_ROWS},">="&{start},PROJECT_MASTER!M2:M{MAX_PROJECT_ROWS},"<"&{end})', None, 5),
            (f'SUMIFS(PROJECT_MASTER!J2:J{MAX_PROJECT_ROWS},PROJECT_MASTER!M2:M{MAX_PROJECT_ROWS},">="&{start},PROJECT_MASTER!M2:M{MAX_PROJECT_ROWS},"<"&{end})', None, 5),
            (f'COUNTIFS(PROJECT_MASTER!O2:O{MAX_PROJECT_ROWS},">="&{start},PROJECT_MASTER!O2:O{MAX_PROJECT_ROWS},"<"&{end},PROJECT_MASTER!T2:T{MAX_PROJECT_ROWS},"Completed")', None, 2),
            (f'COUNTIFS(PROJECT_ACTIVITY!A2:A{MAX_ACTIVITY_ROWS},">="&{start},PROJECT_ACTIVITY!A2:A{MAX_ACTIVITY_ROWS},"<"&{end})', None, 2),
        ])
    add_title("Investment Trend / Budget vs Actual / Activity by Month")
    add_table(["Month", "Investment Trend", "Actual Cost Trend", "Monthly Project Completion", "Activity Count"], month_rows)

    for title, lookup_name, source_col, value_col in [
        ("Investment by Division", "Division", "C", "I"),
        ("Investment by Category", "Category", "E", "I"),
        ("Investment by Site", "Site", "D", "I"),
        ("Project Status Distribution", "Project Status", "T", None),
        ("Project Health Distribution", "Project Health", "U", None),
    ]:
        lookup_values = LOOKUPS[lookup_name]
        data = []
        for item in lookup_values:
            if value_col:
                formula = f'SUMIF(PROJECT_MASTER!{source_col}2:{source_col}{MAX_PROJECT_ROWS},"{item}",PROJECT_MASTER!{value_col}2:{value_col}{MAX_PROJECT_ROWS})'
                style = 5
            else:
                formula = f'COUNTIF(PROJECT_MASTER!{source_col}2:{source_col}{MAX_PROJECT_ROWS},"{item}")'
                style = 2
            data.append([(None, item, 2), (formula, None, style)])
        add_title(title)
        add_table(["Label", "Value"], data)

    widths = "".join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate([34, 22, 22, 26, 18], 1))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews><cols>{widths}</cols><sheetData>{"".join(rows)}</sheetData></worksheet>'''


def readme_xml() -> str:
    rows = [
        ["BRD Sulona Backend Database"],
        ["1. Isi PROJECT_MASTER untuk setiap proyek investasi."],
        ["2. Isi PROJECT_ACTIVITY setiap ada progress atau aktivitas proyek."],
        ["3. Isi PROJECT_DOCUMENT saat dokumen di-upload atau diperbarui."],
        ["4. Isi PROJECT_REVIEW setelah status proyek Completed."],
        ["5. Jangan mengubah nama sheet."],
        ["6. Jangan mengubah nama kolom."],
        ["7. Gunakan pilihan dropdown yang tersedia."],
        ["8. Data dashboard diambil dari DASHBOARD_SUMMARY dan CHART_DATA."],
        ["9. Kolom formula dihitung otomatis, jangan ditimpa manual."],
    ]
    xml_rows = [f'<row r="{i}">{cell(f"A{i}", row[0], style=1 if i == 1 else 2)}</row>' for i, row in enumerate(rows, 1)]
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews><cols><col min="1" max="1" width="96" customWidth="1"/></cols><sheetData>{"".join(xml_rows)}</sheetData></worksheet>'''


def content_types(table_count: int) -> str:
    sheet_overrides = "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, len(SHEETS) + 1))
    table_overrides = "".join(f'<Override PartName="/xl/tables/table{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' for i in range(1, table_count + 1))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>{sheet_overrides}{table_overrides}</Types>'''


ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'''
WORKBOOK = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>''' + "".join(f'<sheet name="{name}" sheetId="{i}" r:id="rId{i}"/>' for i, name in enumerate(SHEETS, 1)) + '</sheets><calcPr calcMode="auto" fullCalcOnLoad="1"/></workbook>'
WORKBOOK_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">''' + "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(SHEETS) + 1)) + f'<Relationship Id="rId{len(SHEETS)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="3"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/><numFmt numFmtId="165" formatCode="&quot;Rp&quot; #,##0"/><numFmt numFmtId="166" formatCode="0%"/></numFmts><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0F172A"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2563EB"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"><color rgb="FFCBD5E1"/></left><right style="thin"><color rgb="FFCBD5E1"/></right><top style="thin"><color rgb="FFCBD5E1"/></top><bottom style="thin"><color rgb="FFCBD5E1"/></bottom></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="7"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/><xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/></cellXfs><dxfs count="5"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFEE2E2"/></patternFill></fill><font><color rgb="FF991B1B"/></font></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFEF3C7"/></patternFill></fill><font><color rgb="FF92400E"/></font></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFDCFCE7"/></patternFill></fill><font><color rgb="FF166534"/></font></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFDBEAFE"/></patternFill></fill><font><color rgb="FF1D4ED8"/></font></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFE5E7EB"/></patternFill></fill><font><color rgb="FF374151"/></font></dxf></dxfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'''
APP = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Codex</Application></Properties>'''
CORE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>BRD Sulona Backend Database</dc:title><dc:creator>Codex</dc:creator></cp:coreProperties>'''


def build() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    table_specs = [
        ("PROJECT_MASTER", PROJECT_HEADERS, MAX_PROJECT_ROWS, "ProjectMasterTable"),
        ("PROJECT_ACTIVITY", ACTIVITY_HEADERS, MAX_ACTIVITY_ROWS, "ProjectActivityTable"),
        ("PROJECT_DOCUMENT", DOCUMENT_HEADERS, MAX_DOCUMENT_ROWS, "ProjectDocumentTable"),
        ("PROJECT_REVIEW", REVIEW_HEADERS, MAX_REVIEW_ROWS, "ProjectReviewTable"),
    ]
    with ZipFile(OUTPUT, "w", ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types(len(table_specs)))
        zf.writestr("_rels/.rels", ROOT_RELS)
        zf.writestr("xl/workbook.xml", WORKBOOK)
        zf.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        zf.writestr("xl/styles.xml", STYLES)
        zf.writestr("docProps/app.xml", APP)
        zf.writestr("docProps/core.xml", CORE)

        zf.writestr("xl/worksheets/sheet1.xml", dashboard_xml())

        zf.writestr("xl/worksheets/sheet2.xml", worksheet_xml("PROJECT_MASTER", PROJECT_HEADERS, project_rows, MAX_PROJECT_ROWS, formula_builder=project_formulas, table_id=1))
        zf.writestr("xl/worksheets/_rels/sheet2.xml.rels", '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>''')
        zf.writestr("xl/tables/table1.xml", table_xml(1, "ProjectMasterTable", f"A1:{col_name(len(PROJECT_HEADERS))}{MAX_PROJECT_ROWS}", PROJECT_HEADERS))

        zf.writestr("xl/worksheets/sheet3.xml", worksheet_xml("PROJECT_ACTIVITY", ACTIVITY_HEADERS, activity_rows, MAX_ACTIVITY_ROWS, formula_builder=lambda r: {3: project_name_formula(r, "B")}, table_id=2))
        zf.writestr("xl/worksheets/_rels/sheet3.xml.rels", '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table2.xml"/></Relationships>''')
        zf.writestr("xl/tables/table2.xml", table_xml(2, "ProjectActivityTable", f"A1:{col_name(len(ACTIVITY_HEADERS))}{MAX_ACTIVITY_ROWS}", ACTIVITY_HEADERS))

        zf.writestr("xl/worksheets/sheet4.xml", worksheet_xml("PROJECT_DOCUMENT", DOCUMENT_HEADERS, document_rows, MAX_DOCUMENT_ROWS, formula_builder=lambda r: {2: project_name_formula(r, "A")}, table_id=3))
        zf.writestr("xl/worksheets/_rels/sheet4.xml.rels", '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table3.xml"/></Relationships>''')
        zf.writestr("xl/tables/table3.xml", table_xml(3, "ProjectDocumentTable", f"A1:{col_name(len(DOCUMENT_HEADERS))}{MAX_DOCUMENT_ROWS}", DOCUMENT_HEADERS))

        zf.writestr("xl/worksheets/sheet5.xml", worksheet_xml("PROJECT_REVIEW", REVIEW_HEADERS, review_rows, MAX_REVIEW_ROWS, formula_builder=lambda r: {2: project_name_formula(r, "A")}, table_id=4))
        zf.writestr("xl/worksheets/_rels/sheet5.xml.rels", '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table4.xml"/></Relationships>''')
        zf.writestr("xl/tables/table4.xml", table_xml(4, "ProjectReviewTable", f"A1:{col_name(len(REVIEW_HEADERS))}{MAX_REVIEW_ROWS}", REVIEW_HEADERS))

        zf.writestr("xl/worksheets/sheet6.xml", lookup_sheet_xml())
        zf.writestr("xl/worksheets/sheet7.xml", dashboard_summary_xml())
        zf.writestr("xl/worksheets/sheet8.xml", chart_data_xml())
        zf.writestr("xl/worksheets/sheet9.xml", readme_xml())
    print(OUTPUT)


if __name__ == "__main__":
    build()
