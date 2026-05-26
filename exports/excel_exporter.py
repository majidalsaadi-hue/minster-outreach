
# Excel export — status export and enhanced template generator.

import io
from datetime import date
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from config.settings import (
    MISA_GREEN, MISA_GOLD, SECTORS, COUNTRIES, INVESTOR_TIERS,
    INVESTOR_STATUSES, JOURNEY_STAGES, MEETING_TYPES, MEETING_STATUSES,
    MEETING_OBJECTIVES, OPPORTUNITY_TYPES, OPPORTUNITY_SOURCES,
    OPPORTUNITY_STAGES, CONFIDENCE_LEVELS, ACTION_STATUSES,
    PROGRESS_OPTIONS, ESCALATION_FLAGS, ENGAGEMENT_TYPES, DEPARTMENTS,
    TASK_PRIORITIES,
)

# Colours (hex without #)
_GREEN  = "1B5C3F"
_GOLD   = "C9974A"
_WHITE  = "FFFFFF"
_LGRAY  = "F7F7F2"
_DGRAY  = "4A4A4A"
_RED    = "C0392B"


def export_status_excel(dfs: dict) -> bytes:
    """Export all current data as a formatted Excel workbook."""
    wb = Workbook()
    wb.remove(wb.active)

    sheet_order = [
        ("Investor Master",    dfs.get("Investor Master",    pd.DataFrame())),
        ("Meeting Log",        dfs.get("Meeting Log",        pd.DataFrame())),
        ("Opportunity Pipeline", dfs.get("Opportunity Pipeline", pd.DataFrame())),
        ("Action Items",       dfs.get("Action Items",       pd.DataFrame())),
        ("RM Tasks",           dfs.get("RM Tasks",           pd.DataFrame())),
    ]

    for sheet_name, df in sheet_order:
        if df.empty:
            ws = wb.create_sheet(sheet_name)
            ws.append([f"No data for {sheet_name}"])
        else:
            ws = wb.create_sheet(sheet_name)
            _write_data_sheet(ws, df, sheet_name)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def generate_template() -> bytes:
    """Generate the enhanced blank Excel template with all sheets, headers,
    dropdowns, and sample rows."""
    wb = Workbook()
    wb.remove(wb.active)

    _build_investor_sheet(wb)
    _build_meeting_sheet(wb)
    _build_opportunity_sheet(wb)
    _build_action_sheet(wb)
    _build_tasks_sheet(wb)
    _build_reference_sheet(wb)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Sheet builders for template ───────────────────────────────────────────────

def _build_investor_sheet(wb: Workbook):
    ws = wb.create_sheet("Investor Master")
    headers = [
        "Investor ID", "Company Name", "Country", "Sector",
        "Investor Tier", "Relationship Manager", "Account Manager",
        "Outreach Manager", "Journey Stage", "Relationship Status",
        "Est. Investment Value (SAR)", "Actual Commitment (SAR)",
        "Last Meeting Date", "Next Meeting Date", "Last Updated",
        "Escalation Flag", "Notes",
    ]
    _write_header_row(ws, headers, _GREEN, _WHITE)
    _set_col_widths(ws, [12, 25, 18, 22, 22, 22, 22, 18, 28, 20, 28, 25, 18, 18, 15, 18, 40])

    # Data validations
    _add_dv(ws, "C", COUNTRIES,       2, 200)
    _add_dv(ws, "D", SECTORS,         2, 200)
    _add_dv(ws, "E", INVESTOR_TIERS,  2, 200)
    _add_dv(ws, "J", INVESTOR_STATUSES, 2, 200)
    _add_dv(ws, "I", JOURNEY_STAGES,  2, 200)
    _add_dv(ws, "P", ESCALATION_FLAGS, 2, 200)

    # Sample row
    ws.append([
        "INV-001", "Siemens Energy", "Germany", "Energy & Industrial",
        "Tier 1 — Strategic", "Khalid Al Sheddi", "TBD", "Majed",
        "Technical Engagement", "Active",
        2_000_000_000, None,
        date(2026, 5, 15), date(2026, 6, 10), date(2026, 5, 26),
        "None", "Turbine factory expansion + new switchgear plant",
    ])
    ws.append([
        "INV-002", "BlackRock", "United States", "Finance & Investment",
        "Tier 1 — Strategic", "Dana", "TBD", "Majed",
        "Opportunity Matching", "Active",
        5_000_000_000, None,
        date(2026, 4, 27), date(2026, 6, 30), date(2026, 4, 27),
        "None", "Infrastructure & data center investment narrative in development",
    ])
    ws.append([
        "INV-003", "Barclays", "United Kingdom", "Finance & Investment",
        "Tier 2 — High Potential", "Dana", "TBD", "Majed",
        "Technical Engagement", "Active",
        None, None,
        date(2026, 5, 15), None, date(2026, 5, 15),
        "None", "Operational readiness + NIS roadshow coordination",
    ])
    ws.append([
        "INV-004", "Brookfield", "United States", "Real Estate",
        "Tier 2 — High Potential", "Dana", "TBD", "Majed",
        "Qualification", "Pending",
        None, None, None, None, date(2026, 5, 26),
        "None", "",
    ])
    _style_data_rows(ws, 2, 5)


def _build_meeting_sheet(wb: Workbook):
    ws = wb.create_sheet("Meeting Log")
    headers = [
        "Meeting ID", "Investor ID", "Company Name", "Meeting Date",
        "Meeting Type", "Location", "MISA Attendees", "Investor Attendees",
        "Meeting Objective", "Key Discussion Points", "Decisions Made",
        "Blockers Identified", "Next Steps", "Follow-Up Owner",
        "Follow-Up Due Date", "Meeting Status", "RM Reviewed", "Logged By",
    ]
    _write_header_row(ws, headers, _GREEN, _WHITE)
    _set_col_widths(ws, [14, 12, 22, 14, 16, 20, 25, 25, 22, 40, 35, 30, 30, 20, 18, 20, 14, 18])

    _add_dv(ws, "E", MEETING_TYPES,      2, 500)
    _add_dv(ws, "I", MEETING_OBJECTIVES, 2, 500)
    _add_dv(ws, "P", MEETING_STATUSES,   2, 500)
    _add_dv(ws, "Q", ["Yes", "No"],      2, 500)

    ws.append([
        "MTG-001", "INV-001", "Siemens Energy", date(2026, 5, 15),
        "In-Person", "Riyadh — MISA HQ",
        "Khalid Al Sheddi, Thamer Almoneef",
        "Dr. Roland Busch (CEO), Regional Director",
        "Technical Discussion",
        "Discussed turbine factory expansion in Dammam and new switchgear plant opportunity",
        "MISA to coordinate site visit with MODON; Siemens to submit investment plan",
        "Land allocation timeline not confirmed", "Arrange MODON coordination meeting",
        "Thamer Almoneef", date(2026, 5, 25), "Follow-Up Required", "Yes", "Thamer Almoneef",
    ])
    _style_data_rows(ws, 2, 2)


def _build_opportunity_sheet(wb: Workbook):
    ws = wb.create_sheet("Opportunity Pipeline")
    headers = [
        "Opportunity ID", "Investor ID", "Company Name", "Opportunity Name",
        "Opportunity Type", "Opportunity Source", "Sector",
        "Opportunity Stage", "Est. Value (SAR)", "Confidence Level",
        "Assigned AM", "Start Date", "Target Closure Date",
        "Opportunity Status", "Blockers", "Escalation Required",
        "Last Updated", "Notes",
    ]
    _write_header_row(ws, headers, _GOLD, _WHITE)
    _set_col_widths(ws, [14, 12, 22, 35, 22, 22, 22, 18, 22, 18, 20, 14, 18, 18, 35, 18, 14, 40])

    _add_dv(ws, "E", OPPORTUNITY_TYPES,   2, 500)
    _add_dv(ws, "F", OPPORTUNITY_SOURCES, 2, 500)
    _add_dv(ws, "G", SECTORS,             2, 500)
    _add_dv(ws, "H", OPPORTUNITY_STAGES,  2, 500)
    _add_dv(ws, "J", CONFIDENCE_LEVELS,   2, 500)
    _add_dv(ws, "N", ["Active", "Under Review", "Blocked", "Converted to Deal", "Dropped"], 2, 500)
    _add_dv(ws, "P", ["Yes", "No"],        2, 500)

    ws.append([
        "OPP-001", "INV-001", "Siemens Energy",
        "Turbine Factory Expansion — Dammam",
        "Mature Opportunity", "Sector Opportunity", "Energy & Industrial",
        "Ready", 2_000_000_000, "High",
        "Thamer Almoneef", date(2026, 5, 8), date(2026, 9, 30),
        "Active", "MODON land allocation pending",
        "Yes", date(2026, 5, 26), "Original factory operating — expansion ready to proceed",
    ])
    ws.append([
        "OPP-002", "INV-001", "Siemens Energy",
        "New Switchgear Manufacturing Plant",
        "Expansion Plan", "Company Expansion", "Energy & Industrial",
        "Development", 800_000_000, "Medium",
        "Thamer Almoneef", date(2026, 5, 12), date(2026, 12, 31),
        "Active", "",
        "No", date(2026, 5, 26), "Siemens submitted plan — MISA review underway",
    ])
    ws.append([
        "OPP-003", "INV-002", "BlackRock",
        "KSA Infrastructure Investment Fund",
        "Partner", "Strategic Partners", "Finance & Investment",
        "Development", 10_000_000_000, "Medium",
        "Naila Alfaifi", date(2026, 5, 3), date(2026, 12, 31),
        "Active", "Awaiting investor list review",
        "No", date(2026, 5, 20), "Investment narrative in progress",
    ])
    _style_data_rows(ws, 2, 4)


def _build_action_sheet(wb: Workbook):
    ws = wb.create_sheet("Action Items")
    headers = [
        "Action ID", "Investor ID", "Company Name", "Meeting ID",
        "Opportunity ID", "Action Description", "Assigned To", "Department",
        "Sector", "Type of Engagement", "Start Date", "Due Date",
        "Priority", "Progress", "Status", "Escalation Flag",
        "Remarks", "Outcome", "Next Action", "Next Action Date",
        "Last Updated", "Updated By",
    ]
    _write_header_row(ws, headers, _GREEN, _WHITE)
    _set_col_widths(ws, [14, 12, 22, 12, 14, 45, 22, 25, 22, 20, 12, 12, 10, 10, 14, 18, 40, 35, 35, 16, 14, 18])

    _add_dv(ws, "I", SECTORS,          2, 1000)
    _add_dv(ws, "J", ENGAGEMENT_TYPES, 2, 1000)
    _add_dv(ws, "M", ["High", "Medium", "Low"], 2, 1000)
    _add_dv(ws, "N", PROGRESS_OPTIONS, 2, 1000)
    _add_dv(ws, "O", ACTION_STATUSES,  2, 1000)
    _add_dv(ws, "P", ESCALATION_FLAGS, 2, 1000)
    _add_dv(ws, "H", DEPARTMENTS,      2, 1000)

    # Load from extracted GA data
    sample_actions = [
        ["ACT-INV-001-001", "INV-001", "Siemens Energy", "MTG-001", "OPP-001",
         "Meeting between MISA & Siemens to understand company interest in Health matchmaking",
         "Thamer Almoneef", "Business Development", "Energy & Industrial", "Opportunity",
         date(2026,5,8), date(2026,5,8), "Medium", "100%", "Completed", "None",
         "Agreed to introduce Siemens to companies matching their profile",
         "Introduction meeting scheduled", "Review Siemens feedback on matched companies", date(2026,5,25),
         date(2026,5,26), "Thamer Almoneef"],
        ["ACT-INV-001-002", "INV-001", "Siemens Energy", "", "OPP-001",
         "Siemens to submit investment plan for Ministry review",
         "Thamer Almoneef", "Business Development", "Energy & Industrial", "Opportunity",
         date(2026,5,12), date(2026,5,14), "Medium", "100%", "Completed", "None",
         "MISA team reviewing Siemens plan for matchmaking opportunities",
         "Plan received and under review", "Complete MISA review and respond", date(2026,5,20),
         date(2026,5,26), "Thamer Almoneef"],
        ["ACT-INV-002-001", "INV-002", "BlackRock", "", "OPP-003",
         "Investment Narrative Development (Top-Down Narrative)",
         "Naila Alfaifi", "HE Outreach Office", "Finance & Investment", "Opportunity",
         date(2026,5,3), date(2026,5,15), "Medium", "75%", "In Progress", "None",
         "Preliminary proposal attached", "", "Finalise narrative following opportunity outputs 1-4",
         date(2026,5,20), date(2026,5,26), "Naila Alfaifi"],
        ["ACT-INV-002-002", "INV-002", "BlackRock", "", "",
         "Water Sector Engagement — Ministry of Water coordination",
         "Ammar Altaf", "Business Development", "Finance & Investment", "Opportunity",
         date(2026,5,4), date(2026,5,25), "High", "25%", "Blocked", "Flag for RM",
         "MEWA requested to put this track on hold pending CoG ownership discussion",
         "", "Await CoG resolution; re-engage MEWA Q3 2026", date(2026,7,1),
         date(2026,5,26), "Ammar Altaf"],
        ["ACT-INV-003-001", "INV-003", "Barclays", "", "",
         "Continued support for Barclays completing operational readiness in KSA",
         "Dana Aljarbu", "Investor Services", "Finance & Investment", "Support",
         date(2026,5,15), date(2026,12,30), "Low", "0%", "Not Started", "None",
         "", "", "Schedule kickoff call with Barclays compliance team", date(2026,6,15),
         date(2026,5,26), "Dana Aljarbu"],
        ["ACT-INV-003-002", "INV-003", "Barclays", "", "",
         "Coordination for virtual meeting in July — investor confidence & NIS",
         "Sara Alsayed", "Business Development", "Finance & Investment", "Opportunity",
         date(2026,5,15), date(2026,5,24), "Medium", "0%", "Not Started", "None",
         "", "", "Confirm July date with Barclays counterpart", date(2026,5,31),
         date(2026,5,26), "Sara Alsayed"],
    ]
    for row in sample_actions:
        ws.append(row)
    _style_data_rows(ws, 2, len(sample_actions) + 1)


def _build_tasks_sheet(wb: Workbook):
    ws = wb.create_sheet("RM Tasks")
    headers = ["Task ID", "Task Title", "Linked Investor", "Priority", "Due Date", "Status", "Notes", "Created Date"]
    _write_header_row(ws, headers, _DGRAY, _WHITE)
    _set_col_widths(ws, [12, 40, 22, 12, 14, 16, 50, 14])
    _add_dv(ws, "D", TASK_PRIORITIES,      2, 200)
    _add_dv(ws, "F", ["Not Started", "In Progress", "Completed"], 2, 200)

    ws.append(["TASK-001", "Follow up on MODON land allocation for Siemens turbine factory",
               "Siemens Energy", "High", date(2026, 5, 30), "Not Started",
               "Coordinate with MODON after receiving their confirmation", date(2026, 5, 26)])
    ws.append(["TASK-002", "Review BlackRock investment narrative draft",
               "BlackRock", "High", date(2026, 5, 20), "In Progress",
               "Preliminary draft received from Naila — review and provide feedback", date(2026, 5, 10)])
    _style_data_rows(ws, 2, 3)


def _build_reference_sheet(wb: Workbook):
    ws = wb.create_sheet("Reference Lists")
    ws.sheet_state = "hidden"   # hide from end-users; used only for dropdowns

    lists = {
        "A": ("Sectors",           SECTORS),
        "B": ("Countries",         COUNTRIES),
        "C": ("Investor Tiers",    INVESTOR_TIERS),
        "D": ("Investor Statuses", INVESTOR_STATUSES),
        "E": ("Journey Stages",    JOURNEY_STAGES),
        "F": ("Meeting Types",     MEETING_TYPES),
        "G": ("Meeting Statuses",  MEETING_STATUSES),
        "H": ("Meeting Objectives",MEETING_OBJECTIVES),
        "I": ("Opp Types",         OPPORTUNITY_TYPES),
        "J": ("Opp Sources",       OPPORTUNITY_SOURCES),
        "K": ("Opp Stages",        OPPORTUNITY_STAGES),
        "L": ("Confidence",        CONFIDENCE_LEVELS),
        "M": ("Action Statuses",   ACTION_STATUSES),
        "N": ("Progress",          PROGRESS_OPTIONS),
        "O": ("Escalation Flags",  ESCALATION_FLAGS),
        "P": ("Engagement Types",  ENGAGEMENT_TYPES),
        "Q": ("Departments",       DEPARTMENTS),
        "R": ("Task Priorities",   TASK_PRIORITIES),
    }
    for col_letter, (header, values) in lists.items():
        col_idx = ord(col_letter) - ord("A") + 1
        ws.cell(row=1, column=col_idx, value=header).font = Font(bold=True)
        for i, v in enumerate(values, start=2):
            ws.cell(row=i, column=col_idx, value=v)


# ── Shared helpers ────────────────────────────────────────────────────────────

def _write_data_sheet(ws, df: pd.DataFrame, title: str):
    header_color = _GREEN if title != "Opportunity Pipeline" else _GOLD
    _write_header_row(ws, list(df.columns), header_color, _WHITE)
    for _, row in df.iterrows():
        ws.append([_clean_val(v) for v in row])
    for col_cells in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col_cells), default=8)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max_len + 4, 50)


def _write_header_row(ws, headers: list, bg_color: str, fg_color: str):
    ws.append(headers)
    header_fill = PatternFill("solid", fgColor=bg_color)
    header_font = Font(bold=True, color=fg_color, size=11)
    for cell in ws[1]:
        cell.fill   = header_fill
        cell.font   = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"


def _style_data_rows(ws, start: int, end: int):
    fill_alt = PatternFill("solid", fgColor="F7F7F2")
    for row_idx in range(start, end + 1):
        for cell in ws[row_idx]:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if row_idx % 2 == 0:
                cell.fill = fill_alt
        ws.row_dimensions[row_idx].height = 18


def _set_col_widths(ws, widths: list):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _add_dv(ws, col: str, options: list, min_row: int, max_row: int):
    formula = '"' + ",".join(str(o) for o in options[:30]) + '"'
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    dv.sqref = f"{col}{min_row}:{col}{max_row}"
    ws.add_data_validation(dv)


def _clean_val(v):
    if isinstance(v, float) and pd.isna(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.date()
    return v
