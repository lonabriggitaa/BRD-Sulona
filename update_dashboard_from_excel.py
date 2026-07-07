from __future__ import annotations

import json
import re
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(r"C:\Users\LENOVO\Documents\Codex\2026-07-01\https-eldcm9-oss-github-io-brd")
INPUT = Path(r"C:\Users\LENOVO\Downloads\brd_sulona_backend_simple_input.xlsx")
APP_SCRIPT = ROOT / "work" / "app-script.js"
WORK_INDEX = ROOT / "work" / "index.html"
OUTPUT_INDEX = ROOT / "outputs" / "index.html"
SUMMARY = ROOT / "outputs" / "dashboard_excel_update_summary.json"

NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def text_of(node: ET.Element | None) -> str:
    if node is None:
        return ""
    return "".join(node.itertext())


def column_index(ref: str) -> int:
    letters = re.match(r"[A-Z]+", ref or "")
    if not letters:
        return 0
    total = 0
    for char in letters.group(0):
        total = total * 26 + (ord(char) - 64)
    return total - 1


def excel_date(value) -> str:
    if value in ("", None):
        return ""
    if isinstance(value, str):
        value = value.strip()
        if not value:
            return ""
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(value, fmt).date().isoformat()
            except ValueError:
                pass
        if re.fullmatch(r"\d+(\.\d+)?", value):
            value = float(value)
        else:
            return value
    if isinstance(value, (int, float)):
        return (date(1899, 12, 30) + timedelta(days=int(value))).isoformat()
    return str(value)


def number(value, default=0):
    if value in ("", None):
        return default
    if isinstance(value, (int, float)):
        return value
    cleaned = str(value).replace("Rp", "").replace(".", "").replace(",", "").strip()
    try:
        return int(float(cleaned))
    except ValueError:
        return default


def first(row: dict, *keys: str, default=""):
    for key in keys:
        if row.get(key) not in ("", None):
            return row.get(key)
    return default


def progress_number(value) -> int:
    raw = number(value, 0)
    if 0 < raw <= 1:
        return int(round(raw * 100))
    return int(raw)


def read_workbook(path: Path) -> dict[str, list[dict]]:
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            shared = [text_of(si) for si in root.findall("x:si", NS)]

        workbook = ET.fromstring(z.read("xl/workbook.xml"))
        rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rels = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels_root
        }

        sheets = {}
        for sheet in workbook.find("x:sheets", NS):
            name = sheet.attrib["name"]
            rel_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = rels[rel_id]
            if not target.startswith("xl/"):
                target = "xl/" + target
            ws = ET.fromstring(z.read(target))
            rows = []
            for row in ws.findall(".//x:sheetData/x:row", NS):
                values = []
                for cell in row.findall("x:c", NS):
                    idx = column_index(cell.attrib.get("r", ""))
                    while len(values) <= idx:
                        values.append("")
                    cell_type = cell.attrib.get("t")
                    raw = text_of(cell.find("x:v", NS))
                    if cell_type == "s" and raw != "":
                        values[idx] = shared[int(raw)]
                    elif cell_type == "inlineStr":
                        values[idx] = text_of(cell.find("x:is", NS))
                    else:
                        values[idx] = raw
                rows.append(values)

            if not rows:
                sheets[name] = []
                continue

            header_row_index = 0
            for idx, row in enumerate(rows[:10]):
                lowered = [str(value).strip().lower() for value in row]
                if "brd_number" in lowered or "brd number" in lowered or "date" in lowered:
                    header_row_index = idx
                    break

            headers = [str(value).strip() for value in rows[header_row_index]]
            assoc_rows = []
            for row in rows[header_row_index + 1:]:
                if not any(str(value).strip() for value in row):
                    continue
                assoc_rows.append({
                    headers[i]: row[i] if i < len(row) else ""
                    for i in range(len(headers))
                    if headers[i]
                })
            sheets[name] = assoc_rows
        return sheets


def status_for_dashboard(status: str, finish: str = "") -> str:
    status = (status or "").strip()
    if status in {"Draft", "Submitted"}:
        return "Submitted"
    if status == "Approved":
        return "Procurement"
    if status in {"Closed", "Completed"}:
        return "Completed"
    if status == "Cancelled":
        return "Overdue"
    if finish:
        try:
            if datetime.fromisoformat(finish).date() < date(2026, 7, 2) and status != "Completed":
                return "Overdue"
        except ValueError:
            pass
    return status or "Running"


def health_for(project: dict) -> str:
    explicit = first(project, "health", "Health")
    if explicit:
        return explicit
    status = first(project, "status", "Status")
    progress = progress_number(first(project, "Progress (%)", "progress"))
    finish = excel_date(first(project, "Target Finish Date", "target_finish"))
    if status in {"Completed", "Closed"} or progress >= 90:
        return "Green"
    try:
        if finish and datetime.fromisoformat(finish).date() < date(2026, 7, 2):
            return "Red"
    except ValueError:
        pass
    if progress < 40:
        return "Yellow"
    return "Yellow"


def normalize_project(row: dict, idx: int) -> dict:
    start = excel_date(first(row, "Planned Start Date", "planned_start", "start"))
    finish = excel_date(first(row, "Target Finish Date", "target_finish", "planned_finish_date"))
    budget = number(first(row, "Budget Requested", "budget"))
    approved = number(first(row, "Budget Approved", "budget", default=budget))
    actual = number(first(row, "Actual Cost", "actual"))
    progress = progress_number(first(row, "Progress (%)", "progress"))
    status = first(row, "Status", "status")
    return {
        "id": idx + 1,
        "brd": first(row, "BRD Number", "brd_number") or f"BRD-EXCEL-{idx+1:03d}",
        "name": first(row, "Project Name", "project_name") or "-",
        "division": first(row, "Division", "division") or "Unassigned",
        "department": first(row, "Department", "department") or "-",
        "site": first(row, "Site", "site") or "Unassigned",
        "category": first(row, "Category", "category") or "Investment",
        "priority": first(row, "Priority", "priority") or "Medium",
        "status": status_for_dashboard(status, finish),
        "health": health_for(row),
        "pic": first(row, "PIC", "pic") or "-",
        "vendor": first(row, "Vendor", "vendor") or "-",
        "budget": budget,
        "approved": approved,
        "actual": actual,
        "progress": progress,
        "start": start,
        "finish": finish,
        "approvalDays": 0,
        "remarks": first(row, "Remarks", "notes") or "",
    }


def normalize_activity(row: dict, projects_by_brd: dict, idx: int) -> dict:
    brd = first(row, "BRD Number", "brd_number")
    project = projects_by_brd.get(brd, {})
    status = first(row, "Status", "status") or "Running"
    if status == "Planned":
        dashboard_status = "Upcoming"
    elif status == "Delayed":
        dashboard_status = "Overdue"
    else:
        dashboard_status = status
    return {
        "id": idx + 1,
        "projectId": project.get("id"),
        "brd": brd,
        "date": excel_date(first(row, "Activity Date", "date")),
        "type": first(row, "Activity", "activity_type") or "Activity",
        "title": first(row, "Activity", "activity_type") or "Activity",
        "project": project.get("name") or first(row, "Project Name", "project_name") or brd,
        "text": first(row, "Description", "note", "Notes") or "-",
        "status": dashboard_status,
        "progress": progress_number(first(row, "Progress (%)", "progress", default=project.get("progress", 0))),
        "pic": first(row, "PIC", "pic") or project.get("pic") or "-",
        "attachment": False,
        "note": first(row, "Notes", "note", "Description") or "",
    }


def normalize_document(row: dict, projects_by_brd: dict, idx: int) -> dict:
    brd = first(row, "BRD Number", "brd_number")
    project = projects_by_brd.get(brd, {})
    return {
        "id": idx + 1,
        "projectId": project.get("id"),
        "BRD Number": brd,
        "Project Name": project.get("name") or row.get("Project Name") or brd,
        "Document Type": first(row, "Document Type", "document_type") or "Others",
        "Document Name": first(row, "Document Name", "document_name") or "Project Document",
        "Upload Date": excel_date(first(row, "Upload Date", "upload_date")),
        "version": first(row, "Version", "version") or "1.0",
        "remarks": first(row, "Remarks", "remarks") or "",
        "project": project.get("name") or brd,
    }


def js_array(name: str, rows: list[dict], declaration: str = "let") -> str:
    return f"{declaration} {name} = {json.dumps(rows, ensure_ascii=False, indent=8)};"


def replace_array(text: str, name: str, replacement: str) -> str:
    match = re.search(rf"\s*(?:let|const)\s+{re.escape(name)}\s*=\s*\[", text)
    if not match:
        raise RuntimeError(f"Could not find array {name}")
    start = match.start()
    cursor = text.index("[", match.start())
    depth = 0
    end = cursor
    while end < len(text):
        char = text[end]
        if char == "[":
            depth += 1
        elif char == "]":
            depth -= 1
            if depth == 0:
                end += 1
                while end < len(text) and text[end] in " \t\r\n;":
                    end += 1
                break
        end += 1
    return text[:start] + "\n      " + replacement.replace("\n", "\n      ") + "\n" + text[end:]


def build_issues(projects: list[dict]) -> list[dict]:
    issues = []
    for project in projects:
        if project["health"] == "Red" or project["status"] == "Overdue":
            issues.append({
                "project": project["name"],
                "title": "Project requires attention",
                "severity": "Critical" if project["health"] == "Red" else "High",
                "due": project.get("finish") or "2026-07-02",
                "owner": project.get("pic") or "-",
                "status": "Open",
            })
        elif project["progress"] < 50 and project["status"] != "Completed":
            issues.append({
                "project": project["name"],
                "title": "Progress below plan",
                "severity": "Medium",
                "due": project.get("finish") or "2026-07-02",
                "owner": project.get("pic") or "-",
                "status": "In Progress",
            })
    return issues[:8]


def build_calendar_items(activities: list[dict]) -> list[dict]:
    items = []
    for activity in activities[:40]:
        items.append({
            "date": activity.get("date") or "2026-07-02",
            "type": activity.get("type") or "Activity",
            "title": activity.get("title") or activity.get("type") or "Activity",
            "project": activity.get("project") or "-",
            "status": activity.get("status") or "Running",
            "progress": activity.get("progress") or 0,
            "attachment": False,
            "note": activity.get("note") or activity.get("text") or "",
        })
    return items


def main() -> None:
    sheets = read_workbook(INPUT)
    project_rows = sheets.get("PROJECT_MASTER") or sheets.get("MASTER_PROJECT") or sheets.get("INPUT_PROJECT") or []
    activity_rows = sheets.get("PROJECT_ACTIVITY") or sheets.get("INPUT_ACTIVITY") or []
    document_rows = sheets.get("PROJECT_DOCUMENT") or sheets.get("INPUT_DOCUMENT") or []

    projects = [normalize_project(row, idx) for idx, row in enumerate(project_rows)]
    by_brd = {project["brd"]: project for project in projects}
    activities = [normalize_activity(row, by_brd, idx) for idx, row in enumerate(activity_rows)]
    documents = [normalize_document(row, by_brd, idx) for idx, row in enumerate(document_rows)]
    issues = build_issues(projects)
    calendar_items = build_calendar_items(activities)

    if not projects:
        raise RuntimeError("PROJECT_MASTER is empty or not found.")

    script = APP_SCRIPT.read_text(encoding="utf-8")
    script = replace_array(script, "projects", js_array("projects", projects))
    script = replace_array(script, "issues", js_array("issues", issues))
    script = replace_array(script, "activities", js_array("activities", activities))
    script = replace_array(script, "documents", js_array("documents", documents))
    script = replace_array(script, "calendarItems", js_array("calendarItems", calendar_items, "const"))
    APP_SCRIPT.write_text(script, encoding="utf-8")

    html = WORK_INDEX.read_text(encoding="utf-8")
    updated = re.sub(r"<script>[\s\S]*?</script>", lambda _match: "<script>\r\n" + script + "\r\n</script>", html, count=1)
    WORK_INDEX.write_text(updated, encoding="utf-8")
    OUTPUT_INDEX.write_text(updated, encoding="utf-8")

    SUMMARY.write_text(json.dumps({
        "source": str(INPUT),
        "projects": len(projects),
        "activities": len(activities),
        "documents": len(documents),
        "issues": len(issues),
        "calendarItems": len(calendar_items),
        "sheets": list(sheets),
    }, indent=2), encoding="utf-8")
    print(SUMMARY)


if __name__ == "__main__":
    main()
