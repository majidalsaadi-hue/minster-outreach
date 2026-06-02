
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
    TASK_PRIORITIES, DEAL_CLASSIFICATIONS, VISION_2030_PILLARS,
    BLOCKER_LEVELS, MINISTER_ACTION_TYPES, SAUDI_CONTENT_OPTIONS,
    TECH_TRANSFER_OPTIONS,
)

# Colours (hex without #)
_GREEN  = "1B5C3F"
_GOLD   = "C9974A"
_WHITE  = "FFFFFF"
_LGRAY  = "F7F7F2"
_DGRAY  = "4A4A4A"
_RED    = "C0392B"


def export_status_excel(dfs: dict) -> bytes:
    """
    Export Action Items in the exact tracker format:
    - One sheet per company: 'Action Items [Company]'
    - Rows 1-12: empty (logo area)
    - Rows 13-16: header block — Outreach / Manager / A-M / RM / Last Updated / Next Meeting
    - Row 20: column headers (green, white)
    - Row 21+: data, color-coded by status and priority
    - Columns A-F empty; data in G–Q
    """
    wb = Workbook()
    wb.remove(wb.active)

    actions   = dfs.get("Action Items",    pd.DataFrame())
    investors = dfs.get("Investor Master", pd.DataFrame())

    if actions.empty:
        ws = wb.create_sheet("Action Items")
        ws["G21"] = "No action items yet."
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    companies = sorted(actions["Company Name"].dropna().unique().tolist()) if "Company Name" in actions.columns else ["All"]

    for company in companies:
        co_acts = actions[actions["Company Name"] == company] if "Company Name" in actions.columns else actions

        # Look up investor metadata
        inv_info = {"manager": "", "am": "", "rm": "", "next_meeting": None}
        if not investors.empty and "Company Name" in investors.columns:
            rows = investors[investors["Company Name"] == company]
            if not rows.empty:
                inv = rows.iloc[0]
                inv_info = {
                    "manager":      str(inv.get("Outreach Manager", "") or ""),
                    "am":           str(inv.get("Account Manager",  "") or ""),
                    "rm":           str(inv.get("Relationship Manager", "") or ""),
                    "next_meeting": _clean_val(inv.get("Next Meeting Date")),
                }

        sheet_name = f"Action Items {company}"[:31]
        ws = wb.create_sheet(sheet_name)
        _write_tracker_sheet(ws, co_acts, inv_info)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _write_tracker_sheet(ws, actions_df: pd.DataFrame, inv_info: dict):
    """Write one company's action items in the exact V5 tracker layout."""
    from openpyxl.styles import numbers as xl_numbers

    GREEN_FILL  = PatternFill("solid", fgColor=_GREEN)
    GOLD_FILL   = PatternFill("solid", fgColor=_GOLD)
    WHITE_BOLD  = Font(bold=True,  color=_WHITE, size=10)
    WHITE_NORM  = Font(bold=False, color=_WHITE, size=10)
    DARK_BOLD   = Font(bold=True,  color=_DGRAY, size=10)
    NORM_FONT   = Font(size=10)
    CENTER      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT        = Alignment(horizontal="left",   vertical="top",    wrap_text=True)
    THIN        = Border(left=Side(style="thin"), right=Side(style="thin"),
                         top=Side(style="thin"),  bottom=Side(style="thin"))

    # ── Rows 1-12: logo area (empty, set row height) ─────────────────────────
    for r in range(1, 13):
        ws.row_dimensions[r].height = 14

    # ── Rows 13-16: header metadata block ────────────────────────────────────
    # Layout (1-indexed cols):
    #   L(12) = label   M(13) = value       P(16) = right-label   Q(17) = right-value
    def _hdr(row, label_col, label_val, val_col=None, val_val=None,
             right_col=None, right_label=None, right_val_col=None, right_val=None):
        c = ws.cell(row=row, column=label_col, value=label_val)
        c.fill, c.font, c.alignment = GREEN_FILL, WHITE_BOLD, CENTER
        if val_col and val_val is not None:
            vc = ws.cell(row=row, column=val_col, value=val_val)
            vc.font, vc.alignment = DARK_BOLD, LEFT
        if right_col and right_label:
            rc = ws.cell(row=row, column=right_col, value=right_label)
            rc.fill, rc.font, rc.alignment = GREEN_FILL, WHITE_BOLD, CENTER
        if right_val_col and right_val is not None:
            rvc = ws.cell(row=row, column=right_val_col, value=right_val)
            rvc.fill, rvc.font, rvc.alignment = GOLD_FILL, WHITE_NORM, CENTER

    _hdr(13, 12, "Outreach",
         right_col=16, right_label="¦ Last Updated",
         right_val_col=17, right_val=date.today())
    _hdr(14, 12, "4 Manager",  13, inv_info["manager"])
    _hdr(15, 12, "4 A-M",      13, inv_info["am"])
    _hdr(16, 12, "4 RM",       13, inv_info["rm"],
         right_col=16, right_label="¹ Next Meeting",
         right_val_col=17, right_val=inv_info["next_meeting"])

    for r in range(13, 20):
        ws.row_dimensions[r].height = 18

    # ── Row 20: column headers ────────────────────────────────────────────────
    COLS = ["ID", "Action Item", "Assigned to", "Sector",
            "Type of Engagement", "Start Date", "Due Date",
            "Priority", "Progress", "Status", "Remarks"]
    for j, h in enumerate(COLS):
        c = ws.cell(row=20, column=7 + j, value=h)
        c.fill, c.font, c.alignment = GREEN_FILL, WHITE_BOLD, CENTER
    ws.row_dimensions[20].height = 22

    # Column widths (A-F narrow; G-Q wide)
    for col in range(1, 7):
        ws.column_dimensions[get_column_letter(col)].width = 2
    for col, w in zip(range(7, 18), [6, 52, 20, 18, 20, 12, 12, 12, 10, 14, 38]):
        ws.column_dimensions[get_column_letter(col)].width = w

    # Freeze panes below header
    ws.freeze_panes = "H21"

    # ── Status / Priority fill maps ───────────────────────────────────────────
    STATUS_FILL = {
        "Completed":   PatternFill("solid", fgColor="E2EFDA"),
        "In Progress": PatternFill("solid", fgColor="FFF2CC"),
        "Inprogress":  PatternFill("solid", fgColor="FFF2CC"),
        "Not Started": PatternFill("solid", fgColor="F2F2F2"),
        "Blocked":     PatternFill("solid", fgColor="FCE4D6"),
        "Cancelled":   PatternFill("solid", fgColor="EDEDED"),
    }
    PRIO_COLOR = {
        "Very High": "C00000",
        "High":      _RED,
        "Medium":    _GOLD,
        "Low":       _GREEN,
    }

    # ── Data rows ─────────────────────────────────────────────────────────────
    for i, (_, row) in enumerate(actions_df.iterrows()):
        r       = 21 + i
        status  = str(row.get("Status", "") or "")
        prio    = str(row.get("Priority", "") or "")
        rfill   = STATUS_FILL.get(status, PatternFill("solid", fgColor="FFFFFF"))

        # Convert progress string ("25%", "0%") → float 0-1
        raw_prog = row.get("Progress", 0)
        if isinstance(raw_prog, str):
            try:
                prog = float(raw_prog.replace("%", "")) / 100
            except ValueError:
                prog = 0.0
        else:
            prog = float(raw_prog) if raw_prog not in (None, "") else 0.0
        # If already stored as percentage float > 1 (e.g. 25 not 0.25) normalise
        if prog > 1.0:
            prog = prog / 100

        data = [
            i + 1,
            str(row.get("Action Description", "") or ""),
            str(row.get("Assigned To", "")        or ""),
            str(row.get("Sector", "")             or ""),
            str(row.get("Type of Engagement", "") or ""),
            _clean_val(row.get("Start Date")),
            _clean_val(row.get("Due Date")),
            prio,
            prog,
            status,
            str(row.get("Remarks", "") or ""),
        ]

        for j, val in enumerate(data):
            c = ws.cell(row=r, column=7 + j, value=val)
            c.fill   = rfill
            c.border = THIN
            c.font   = NORM_FONT

            # Progress as % format
            if j == 8:
                c.number_format = "0%"
                c.alignment = CENTER
            # Priority — bold + colour
            elif j == 7:
                c.font = Font(size=10, bold=True, color=PRIO_COLOR.get(prio, _DGRAY))
                c.alignment = CENTER
            # ID centred
            elif j == 0:
                c.alignment = CENTER
            # Dates centred
            elif j in (5, 6):
                c.alignment = CENTER
            else:
                c.alignment = LEFT

        ws.row_dimensions[r].height = 40


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
        "Est. Jobs Created", "Saudi Content %", "Technology Transfer",
        "Deal Classification", "Vision 2030 Pillar",
        "Strategic Priority Score",
        "Minister Action Required", "Decision Required By", "Blocker Level",
        "Last Meeting Date", "Next Meeting Date", "Last Updated",
        "Escalation Flag", "Notes",
    ]
    _write_header_row(ws, headers, _GREEN, _WHITE)
    _set_col_widths(ws, [
        12, 25, 18, 22, 22, 22, 22, 18, 28, 20,
        28, 25, 18, 18, 18,
        28, 30, 10,
        30, 18, 38,
        18, 18, 15, 18, 40,
    ])

    # Core dropdowns
    _add_dv(ws, "C", COUNTRIES,            2, 200)
    _add_dv(ws, "D", SECTORS,              2, 200)
    _add_dv(ws, "E", INVESTOR_TIERS,       2, 200)
    _add_dv(ws, "I", JOURNEY_STAGES,       2, 200)
    _add_dv(ws, "J", INVESTOR_STATUSES,    2, 200)
    # Minister decision-support dropdowns
    _add_dv(ws, "N", SAUDI_CONTENT_OPTIONS, 2, 200)
    _add_dv(ws, "O", TECH_TRANSFER_OPTIONS, 2, 200)
    _add_dv(ws, "P", DEAL_CLASSIFICATIONS,  2, 200)
    _add_dv(ws, "Q", VISION_2030_PILLARS,   2, 200)
    _add_dv(ws, "S", MINISTER_ACTION_TYPES, 2, 200)
    _add_dv(ws, "U", BLOCKER_LEVELS,        2, 200)
    _add_dv(ws, "Y", ESCALATION_FLAGS,      2, 200)

    # Sample rows with minister decision fields
    ws.append([
        "INV-001", "Siemens Energy", "Germany", "Energy & Industrial",
        "Tier 1 — Strategic", "Khalid Al Sheddi", "TBD", "Majed",
        "Technical Engagement", "Active",
        2_000_000_000, None,
        8500, "40–60%", "Yes",
        "Greenfield", "Thriving Economy",
        4,
        "Site Visit", date(2026, 6, 15), "Ministerial — requires HE intervention",
        date(2026, 5, 15), date(2026, 6, 10), date(2026, 5, 26),
        "Flag for RM", "Turbine factory expansion + new switchgear plant",
    ])
    ws.append([
        "INV-002", "BlackRock", "United States", "Finance & Investment",
        "Tier 1 — Strategic", "Dana", "TBD", "Majed",
        "Opportunity Matching", "Active",
        5_000_000_000, None,
        3200, "20–40%", "No",
        "Fund / FDI", "Thriving Economy",
        5,
        "Decision Required", date(2026, 7, 31), "None",
        date(2026, 4, 27), date(2026, 6, 30), date(2026, 4, 27),
        "None", "Infrastructure & data center investment narrative in development",
    ])
    ws.append([
        "INV-003", "Barclays", "United Kingdom", "Finance & Investment",
        "Tier 2 — High Potential", "Dana", "TBD", "Majed",
        "Technical Engagement", "Active",
        None, None,
        1200, "< 20%", "No",
        "Strategic Partnership", "Thriving Economy",
        3,
        "None Required", None, "None",
        date(2026, 5, 15), None, date(2026, 5, 15),
        "None", "Operational readiness + NIS roadshow coordination",
    ])
    ws.append([
        "INV-004", "Brookfield", "United States", "Real Estate",
        "Tier 2 — High Potential", "Dana", "TBD", "Majed",
        "Qualification", "Pending",
        None, None,
        None, "TBD", "TBD",
        "Greenfield", "Vibrant Society",
        2,
        "None Required", None, "None",
        None, None, date(2026, 5, 26),
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
        "A": ("Sectors",             SECTORS),
        "B": ("Countries",           COUNTRIES),
        "C": ("Investor Tiers",      INVESTOR_TIERS),
        "D": ("Investor Statuses",   INVESTOR_STATUSES),
        "E": ("Journey Stages",      JOURNEY_STAGES),
        "F": ("Meeting Types",       MEETING_TYPES),
        "G": ("Meeting Statuses",    MEETING_STATUSES),
        "H": ("Meeting Objectives",  MEETING_OBJECTIVES),
        "I": ("Opp Types",           OPPORTUNITY_TYPES),
        "J": ("Opp Sources",         OPPORTUNITY_SOURCES),
        "K": ("Opp Stages",          OPPORTUNITY_STAGES),
        "L": ("Confidence",          CONFIDENCE_LEVELS),
        "M": ("Action Statuses",     ACTION_STATUSES),
        "N": ("Progress",            PROGRESS_OPTIONS),
        "O": ("Escalation Flags",    ESCALATION_FLAGS),
        "P": ("Engagement Types",    ENGAGEMENT_TYPES),
        "Q": ("Departments",         DEPARTMENTS),
        "R": ("Task Priorities",     TASK_PRIORITIES),
        "S": ("Deal Classifications",DEAL_CLASSIFICATIONS),
        "T": ("Vision 2030 Pillars", VISION_2030_PILLARS),
        "U": ("Blocker Levels",      BLOCKER_LEVELS),
        "V": ("Minister Actions",    MINISTER_ACTION_TYPES),
        "W": ("Saudi Content",       SAUDI_CONTENT_OPTIONS),
        "X": ("Tech Transfer",       TECH_TRANSFER_OPTIONS),
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
