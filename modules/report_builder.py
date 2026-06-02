
# Report Builder
# Upload: Word meeting minutes (Arabic) + company Excel tracker
# Output: updated Excel tracker, external email draft, CRM action items sync

import io
import re
import pandas as pd
import streamlit as st
from datetime import date, datetime
from pathlib import Path

import docx
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config.settings import MISA_GREEN, MISA_GOLD, MISA_GREEN_DARK, ENGAGEMENT_TYPES
from config.translations import t
from modules.persistence import save_session


_PRIORITIES  = ["Very High", "High", "Medium", "Low"]
_ENGAGE_OPTS = ["Support", "Opportunity", "Challenge", "Follow-up", "Action", "Administrative"]
_PRIORITY_AR = {"Very High": "مهم جدا", "High": "مهم", "Medium": "متوسط", "Low": "عادي"}
_PRIORITY_EN = {"مهم جدا": "Very High", "مهم": "High", "متوسط": "Medium", "عادي": "Low",
                "مستمر": "Ongoing"}

_EMPTY_ACTIONS = pd.DataFrame(columns=[
    "Action (AR)", "Action (EN)", "Assigned To", "Type", "Priority", "Due Date", "Remarks"
])

_COL_HEADER_AR = ["م", "التوجيه / المهمة", "المسؤول", "الأولوية", "تاريخ الإنجاز المتوقع"]
_COL_HEADER_EN = ["ID", "Action Item", "Assigned to", "Type of Engagement",
                  "Start Date", "Due Date", "Priority", "Progress", "Status", "Remarks"]


# ── Session state ─────────────────────────────────────────────────────────────

def _init():
    defaults = {
        "rb2_parsed":      None,   # dict with extracted meeting data
        "rb2_actions":     _EMPTY_ACTIONS.copy(),
        "rb2_company":     "",
        "rb2_date":        date.today(),
        "rb2_subject_ar":  "",
        "rb2_subject_en":  "",
        "rb2_location":    "المقر الرئيسي – وزارة الاستثمار",
        "rb2_chair":       "معالي الوزير",
        "rb2_next_mtg":    None,
        "rb2_attendees":   "",
        "rb2_disc_ar":     "",
        "rb2_disc_en":     "",
        "rb2_excel_bytes": None,   # raw bytes of uploaded tracker Excel
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ── Main render ───────────────────────────────────────────────────────────────

def render(dfs: dict, lang: str):
    _init()

    st.markdown("### 📝 Report Builder")
    st.caption(
        "Upload meeting minutes (Word) + company Excel tracker → "
        "review extracted data → download updated Excel, copy email draft, sync to CRM."
    )

    # ── Upload row ────────────────────────────────────────────────────────────
    up1, up2 = st.columns(2)
    with up1:
        st.markdown("**1 — Meeting Minutes (.docx)**")
        word_file = st.file_uploader("Word file", type=["docx"], key="_rb2_word",
                                     label_visibility="collapsed")
    with up2:
        st.markdown("**2 — Company Excel Tracker (.xlsx)** — optional")
        excel_file = st.file_uploader("Excel file", type=["xlsx"], key="_rb2_excel",
                                      label_visibility="collapsed")

    # Parse Word on upload
    if word_file is not None:
        raw = word_file.read()
        parsed = _parse_word(raw)
        if parsed:
            st.session_state["rb2_parsed"]     = parsed
            st.session_state["rb2_company"]    = parsed.get("company", "")
            st.session_state["rb2_date"]       = parsed.get("date") or date.today()
            st.session_state["rb2_subject_ar"] = parsed.get("subject_ar", "")
            st.session_state["rb2_subject_en"] = parsed.get("subject_en", "")
            st.session_state["rb2_location"]   = parsed.get("location", "المقر الرئيسي – وزارة الاستثمار")
            st.session_state["rb2_chair"]      = parsed.get("chair", "معالي الوزير")
            st.session_state["rb2_next_mtg"]   = parsed.get("next_meeting")
            st.session_state["rb2_attendees"]  = parsed.get("attendees", "")
            st.session_state["rb2_disc_ar"]    = parsed.get("discussion_ar", "")
            if parsed.get("action_items"):
                st.session_state["rb2_actions"] = pd.DataFrame(parsed["action_items"])
            st.success(f"✅ Parsed: {len(parsed.get('action_items', []))} action items extracted.")

    if excel_file is not None:
        st.session_state["rb2_excel_bytes"] = excel_file.read()
        sheets = _get_excel_sheets(st.session_state["rb2_excel_bytes"])
        st.success(f"✅ Excel loaded — sheets: {', '.join(sheets)}")

    st.markdown("---")

    # ── Review / edit section ─────────────────────────────────────────────────
    with st.expander("📋 Review & Edit Meeting Details", expanded=True):
        _meta_form(dfs)

    with st.expander("💬 Discussion Points", expanded=False):
        _discussion_form()

    with st.expander("✅ Review & Edit Action Items", expanded=True):
        _action_items_form()

    st.markdown("---")

    # ── Output tabs ────────────────────────────────────────────────────────────
    tab1, tab2, tab3 = st.tabs([
        "📊 Updated Excel Tracker",
        "✉️ Email Draft (English)",
        "🔄 Sync to CRM",
    ])

    with tab1:
        _excel_output_tab()

    with tab2:
        _email_tab()

    with tab3:
        _sync_tab(dfs, lang)


# ── Review forms ──────────────────────────────────────────────────────────────

def _meta_form(dfs: dict):
    investors = dfs.get("Investor Master", pd.DataFrame())
    companies = [""] + (
        sorted(investors["Company Name"].dropna().unique().tolist())
        if not investors.empty and "Company Name" in investors.columns else []
    )
    s = st.session_state

    c1, c2, c3 = st.columns(3)
    s["rb2_company"]  = c1.selectbox("Company", companies,
        index=companies.index(s["rb2_company"]) if s["rb2_company"] in companies else 0,
        key="_rb2_co")
    s["rb2_date"]     = c2.date_input("Meeting Date", value=s["rb2_date"], key="_rb2_date")
    s["rb2_chair"]    = c3.text_input("Chaired By / برئاسة", value=s["rb2_chair"], key="_rb2_chair")

    c4, c5, c6 = st.columns(3)
    s["rb2_location"] = c4.text_input("Location", value=s["rb2_location"], key="_rb2_loc")
    s["rb2_next_mtg"] = c5.date_input("Next Meeting", value=s["rb2_next_mtg"], key="_rb2_next")
    s["rb2_subject_ar"] = c6.text_input("Subject (Arabic)", value=s["rb2_subject_ar"], key="_rb2_sub_ar")
    s["rb2_subject_en"] = st.text_input("Subject (English) — used in the email subject line",
                                        value=s["rb2_subject_en"], key="_rb2_sub_en")
    st.markdown("**Attendees** — Name | Title, one per line")
    s["rb2_attendees"] = st.text_area("Attendees", value=s["rb2_attendees"], height=80,
                                      key="_rb2_att", label_visibility="collapsed")


def _discussion_form():
    s = st.session_state
    c1, c2 = st.columns(2)
    s["rb2_disc_ar"] = c1.text_area("Arabic (from document)", value=s["rb2_disc_ar"],
                                     height=200, key="_rb2_dar")
    s["rb2_disc_en"] = c2.text_area("English (optional — leave blank to use Arabic)",
                                     value=s["rb2_disc_en"], height=200, key="_rb2_den")


def _action_items_form():
    st.caption("Auto-extracted from Word. Edit as needed before generating outputs.")
    edited = st.data_editor(
        st.session_state["rb2_actions"],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Action (AR)":  st.column_config.TextColumn("Action Item (Arabic)", width="large"),
            "Action (EN)":  st.column_config.TextColumn("Action Item (English)", width="large"),
            "Assigned To":  st.column_config.TextColumn(width="medium"),
            "Type":         st.column_config.SelectboxColumn(options=_ENGAGE_OPTS, width="medium"),
            "Priority":     st.column_config.SelectboxColumn(options=_PRIORITIES, width="small"),
            "Due Date":     st.column_config.DateColumn(width="small"),
            "Remarks":      st.column_config.TextColumn(width="large"),
        },
        key="_rb2_actions_ed",
        hide_index=True,
    )
    st.session_state["rb2_actions"] = edited


# ── Tab 1: Updated Excel output ────────────────────────────────────────────────

def _excel_output_tab():
    s       = st.session_state
    company = s["rb2_company"]
    actions = s["rb2_actions"]

    valid = _valid_actions(actions)
    if valid.empty:
        st.info("Add action items in the review section above.")
        return

    if not company:
        st.warning("Select a company in the meeting details section.")
        return

    st.markdown(f"**{len(valid)} action item(s)** will be written to sheet **'{company}'** in the Excel tracker.")

    existing_bytes = s.get("rb2_excel_bytes")
    if existing_bytes:
        sheets = _get_excel_sheets(existing_bytes)
        if company in sheets:
            st.info(f"Sheet '{company}' found in uploaded Excel — new rows will be appended after existing ones.")
        else:
            st.info(f"Sheet '{company}' not found — a new sheet will be created with the standard format.")
    else:
        st.info("No Excel uploaded — a new tracker file will be created.")

    if st.button("📊 Generate Updated Excel", type="primary", key="_rb2_gen_xl"):
        xl_bytes = _build_excel(
            existing_bytes=existing_bytes,
            company=company,
            meeting_date=s["rb2_date"],
            next_meeting=s["rb2_next_mtg"],
            chair=s["rb2_chair"],
            actions_df=valid,
        )
        fname = f"ActionTracker_{company.replace(' ', '_')}_{s['rb2_date']}.xlsx"
        st.download_button(
            f"⬇️ Download {fname}",
            data=xl_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.success("Excel ready — click the download button above.")


# ── Tab 2: Email draft ────────────────────────────────────────────────────────

def _email_tab():
    st.markdown("#### External Email — Professional English Draft")
    s = st.session_state
    email_text = _build_email(s)
    st.text_area("Email (copy or edit before sending)", value=email_text,
                 height=500, key="_rb2_email_prev")
    st.download_button(
        "⬇️ Download as .txt",
        data=email_text.encode("utf-8"),
        file_name=f"Email_{s['rb2_company']}_{s['rb2_date']}.txt",
        mime="text/plain",
    )


# ── Tab 3: CRM sync ────────────────────────────────────────────────────────────

def _sync_tab(dfs: dict, lang: str):
    st.markdown("#### Add Action Items to CRM Tracker")
    s       = st.session_state
    company = s["rb2_company"]
    valid   = _valid_actions(s["rb2_actions"])

    if valid.empty:
        st.info("No action items to sync.")
        return
    if not company:
        st.warning("Select a company first.")
        return

    st.markdown(f"**{len(valid)} item(s)** → Company: **{company}**")
    st.dataframe(
        valid[["Action (EN)", "Assigned To", "Priority", "Due Date"]],
        use_container_width=True, hide_index=True,
    )

    if st.button("🔄 Add to CRM Action Items", type="primary", key="_rb2_sync"):
        investors = dfs.get("Investor Master", pd.DataFrame())
        inv_id    = _investor_id(investors, company)
        existing  = dfs.get("Action Items", pd.DataFrame())
        new_rows  = []

        for i, (_, row) in enumerate(valid.iterrows()):
            desc   = row.get("Action (EN)") or row.get("Action (AR)", "")
            new_id = _next_act_id(existing, len(new_rows))
            nr = {
                "Action ID":          new_id,
                "Investor ID":        inv_id,
                "Company Name":       company,
                "Meeting ID":         "",
                "Opportunity ID":     "",
                "Action Description": desc,
                "Assigned To":        row.get("Assigned To", ""),
                "Department":         "",
                "Sector":             "",
                "Type of Engagement": row.get("Type", "Action"),
                "Start Date":         s["rb2_date"],
                "Due Date":           row.get("Due Date") if pd.notna(row.get("Due Date")) else None,
                "Priority":           _map_prio(row.get("Priority", "Medium")),
                "Progress":           "0%",
                "Status":             "Not Started",
                "Escalation Flag":    "None",
                "Remarks":            row.get("Remarks", ""),
                "Outcome":            "",
                "Next Action":        "",
                "Next Action Date":   None,
                "Last Updated":       date.today(),
                "Updated By":         "",
            }
            new_rows.append(nr)
            existing = pd.concat([existing, pd.DataFrame([nr])], ignore_index=True)

        dfs["Action Items"] = existing
        save_session(dfs)
        st.success(f"✅ {len(new_rows)} action item(s) added for {company}.")
        st.balloons()


# ── Word parser ───────────────────────────────────────────────────────────────

def _parse_word(file_bytes: bytes) -> dict:
    try:
        doc = docx.Document(io.BytesIO(file_bytes))
    except Exception:
        return {}

    result = {
        "company": "", "date": None, "subject_ar": "", "subject_en": "",
        "location": "", "chair": "معالي الوزير", "priority": "Very High",
        "next_meeting": None, "attendees": "", "discussion_ar": "", "action_items": [],
    }

    # ── Pull all table data ───────────────────────────────────────────────────
    tables = []
    for tbl in doc.tables:
        rows = []
        for row in tbl.rows:
            cells = [c.text.strip() for c in row.cells]
            rows.append(cells)
        tables.append(rows)

    # ── Pull all paragraph text ───────────────────────────────────────────────
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    # ── Try to find meeting metadata from first table(s) ─────────────────────
    for tbl in tables[:3]:
        flat = " ".join(" ".join(r) for r in tbl)
        # Date
        m = re.search(r"(\d{4}[-/]\d{1,2}[-/]\d{1,2})", flat)
        if m and not result["date"]:
            try:
                result["date"] = datetime.strptime(m.group(1).replace("/", "-"), "%Y-%m-%d").date()
            except ValueError:
                pass
        # Subject
        for row in tbl:
            joined = " ".join(row)
            if any(kw in joined for kw in ["مستجدات", "اجتماع", "تحديث", "متابعة"]):
                # Grab any Arabic text that isn't a label
                for cell in row:
                    if cell and not any(lbl in cell for lbl in ["الموضوع", "التاريخ", "اليوم", "الموقع"]):
                        if len(cell) > 4:
                            result["subject_ar"] = cell
                            break
            if "المقر" in joined or "وزارة" in joined:
                for cell in row:
                    if "المقر" in cell or "وزارة" in cell:
                        result["location"] = cell
            if "معالي" in joined:
                for cell in row:
                    if "معالي" in cell:
                        result["chair"] = cell
            if "مهم جدا" in joined:
                result["priority"] = "Very High"
            elif "مهم" in joined:
                result["priority"] = "High"
            elif "متوسط" in joined:
                result["priority"] = "Medium"

    # Next meeting from paragraphs or tables
    for p in paras + [" ".join(c for r in t for c in r) for t in tables]:
        m = re.search(r"يوليو|أغسطس|سبتمبر|أكتوبر|نوفمبر|ديسمبر|يناير|فبراير|مارس|إبريل|مايو|يونيو", p)
        if m and not result["next_meeting"]:
            result["next_meeting"] = None  # keep as None; the month text is in next_mtg string

    # ── Company from subject or filename ─────────────────────────────────────
    known = ["Barclays", "BlackRock", "Brookfield", "Goldman", "HSBC",
             "JPMorgan", "Morgan Stanley", "UBS", "Citi", "Deutsche"]
    full_text = " ".join(paras)
    for name in known:
        if name.lower() in full_text.lower():
            result["company"] = name
            if not result["subject_en"]:
                result["subject_en"] = f"Latest Updates — {name}"
            break

    # ── Discussion points ─────────────────────────────────────────────────────
    disc_lines = []
    capturing  = False
    for p in paras:
        if "أبرز ما تم مناقشته" in p or "أبرز نقاط" in p:
            capturing = True
            continue
        if capturing:
            if any(h in p for h in ["توجيهات معاليه", "الحضور", "بنود العمل"]):
                capturing = False
                continue
            if p:
                disc_lines.append(p.lstrip("•● -"))
    result["discussion_ar"] = "\n".join(disc_lines)

    # ── Action items from tables ──────────────────────────────────────────────
    for tbl in tables:
        if not tbl:
            continue
        # Check if this looks like an action items table
        header_flat = " ".join(tbl[0])
        is_actions = any(kw in header_flat for kw in ["المهمة", "التوجيه", "Action Item", "ID"])
        if not is_actions and len(tbl) >= 2:
            second_flat = " ".join(tbl[1]) if len(tbl) > 1 else ""
            is_actions  = any(kw in second_flat for kw in ["المهمة", "التوجيه"])
        if not is_actions:
            continue

        for row in tbl[1:]:  # skip header
            cells = [c.strip() for c in row if c.strip()]
            if len(cells) < 2:
                continue
            # Skip if first cell is a number index label
            text_cells = [c for c in cells if not re.match(r"^\d+$", c)]
            if not text_cells:
                continue
            action_ar = text_cells[0] if text_cells else ""
            # Try to find owner, priority, due date from remaining cells
            owner = ""
            prio  = "High"
            due   = None
            for cell in text_cells[1:]:
                if re.search(r"\d{4}", cell):
                    try:
                        m = re.search(r"(\d{4}[-/]\d{1,2}[-/]\d{1,2})", cell)
                        if m:
                            due = datetime.strptime(m.group(1).replace("/", "-"), "%Y-%m-%d").date()
                    except Exception:
                        pass
                elif cell in _PRIORITY_EN:
                    prio = _PRIORITY_EN[cell]
                elif cell in ["مستمر", "Ongoing"]:
                    prio = "High"
                elif len(cell) > 1 and not any(cell == k for k in _PRIORITY_AR.values()):
                    if not owner:
                        owner = cell
            if action_ar and action_ar not in ("م", "التوجيه / المهمة", "Action Item"):
                result["action_items"].append({
                    "Action (AR)":  action_ar,
                    "Action (EN)":  "",
                    "Assigned To":  owner,
                    "Type":         "Action",
                    "Priority":     prio,
                    "Due Date":     due,
                    "Remarks":      "",
                })

    # ── Attendees from last table ──────────────────────────────────────────────
    for tbl in reversed(tables):
        header_flat = " ".join(tbl[0]) if tbl else ""
        if "الاسم" in header_flat or "الوظيفة" in header_flat or "الحضور" in header_flat:
            att_lines = []
            for row in tbl[1:]:
                cells = [c.strip() for c in row if c.strip()]
                if len(cells) >= 2:
                    att_lines.append(f"{cells[0]} | {cells[1]}")
                elif cells:
                    att_lines.append(cells[0])
            result["attendees"] = "\n".join(att_lines)
            break

    return result


# ── Excel builder ─────────────────────────────────────────────────────────────

def _build_excel(existing_bytes, company, meeting_date, next_meeting, chair, actions_df) -> bytes:
    if existing_bytes:
        wb = openpyxl.load_workbook(io.BytesIO(existing_bytes))
    else:
        wb = openpyxl.Workbook()
        # Remove default sheet
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    sheet_name = company[:31]  # Excel sheet name max 31 chars

    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # Find the last data row (after the header at row 20)
        last_row = 20
        for row in ws.iter_rows(min_row=21, values_only=True):
            if any(v is not None for v in row):
                last_row += 1
        start_row = last_row + 1
    else:
        ws = wb.create_sheet(sheet_name)
        _write_sheet_header(ws, company, meeting_date, next_meeting, chair)
        start_row = 21

    green_fill  = PatternFill("solid", fgColor="1B5C3F")
    white_font  = Font(color="FFFFFF", bold=True, size=10)
    normal_font = Font(size=10)
    center_al   = Alignment(horizontal="center", vertical="center")
    right_al    = Alignment(horizontal="right", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    status_colors = {
        "Not Started": "D9D9D9", "In Progress": "FFD966",
        "Completed": "70AD47", "Blocked": "FF0000",
    }

    for i, (_, row) in enumerate(actions_df.iterrows()):
        r = start_row + i
        due = row.get("Due Date")
        due_val = due if (due and str(due) not in ("NaT", "None", "")) else None
        prio_val = _map_prio(row.get("Priority", "Medium"))

        ws.cell(r, 7, i + 1)              # ID
        ws.cell(r, 8, row.get("Action (EN)") or row.get("Action (AR)", ""))
        ws.cell(r, 9, row.get("Assigned To", ""))
        ws.cell(r, 10, row.get("Type", "Action"))
        ws.cell(r, 11, meeting_date)       # Start Date
        ws.cell(r, 12, due_val)            # Due Date
        ws.cell(r, 13, prio_val)           # Priority
        ws.cell(r, 14, 0)                  # Progress (0%)
        ws.cell(r, 15, "Not Started")      # Status
        ws.cell(r, 16, row.get("Remarks", ""))

        for col in range(7, 17):
            cell = ws.cell(r, col)
            cell.font   = normal_font
            cell.border = thin_border
            cell.alignment = right_al if col == 8 else center_al

        # Colour-code status
        status_hex = status_colors.get("Not Started", "D9D9D9")
        ws.cell(r, 15).fill = PatternFill("solid", fgColor=status_hex)

        ws.row_dimensions[r].height = 40

    # Column widths (if new sheet)
    if start_row == 21:
        widths = {7: 6, 8: 50, 9: 18, 10: 18, 11: 12, 12: 12, 13: 12, 14: 10, 15: 14, 16: 30}
        for col, w in widths.items():
            ws.column_dimensions[get_column_letter(col)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _write_sheet_header(ws, company, meeting_date, next_meeting, chair):
    green_fill = PatternFill("solid", fgColor="1B5C3F")
    gold_fill  = PatternFill("solid", fgColor="C9974A")
    white_font = Font(color="FFFFFF", bold=True, size=11)
    gold_font  = Font(color="FFFFFF", bold=True, size=10)
    center_al  = Alignment(horizontal="center", vertical="center")
    right_al   = Alignment(horizontal="right",  vertical="center")

    # Row 13: Company + Last Updated
    ws.cell(13, 11, "Outreach").font = white_font
    ws.cell(13, 11).fill = green_fill
    ws.cell(13, 15, "Last Updated").font = white_font
    ws.cell(13, 15).fill = green_fill
    ws.cell(13, 16, meeting_date).font = white_font
    ws.cell(13, 16).fill = gold_fill

    # Row 14: Manager
    ws.cell(14, 11, "Manager").font = white_font
    ws.cell(14, 11).fill = green_fill
    ws.cell(14, 12, chair).font = Font(bold=True, size=10)

    # Row 16: Next Meeting
    ws.cell(16, 15, "Next Meeting").font = white_font
    ws.cell(16, 15).fill = green_fill
    ws.cell(16, 16, next_meeting or "TBD").font = gold_font
    ws.cell(16, 16).fill = gold_fill

    # Rows 13-19: set height
    for r in range(13, 20):
        ws.row_dimensions[r].height = 18

    # Row 20: Column headers
    headers = ["ID", "Action Item", "Assigned to", "Type of Engagement",
               "Start Date", "Due Date", "Priority", "Progress", "Status", "Remarks"]
    for j, h in enumerate(headers):
        cell = ws.cell(20, 7 + j, h)
        cell.fill      = green_fill
        cell.font      = white_font
        cell.alignment = center_al
    ws.row_dimensions[20].height = 22

    # Freeze panes below header
    ws.freeze_panes = "G21"


# ── Email builder ─────────────────────────────────────────────────────────────

def _build_email(s: dict) -> str:
    co     = s["rb2_company"] or "[Company]"
    mtg    = s["rb2_date"]
    sub_en = s["rb2_subject_en"] or s["rb2_subject_ar"] or "Meeting Follow-up"
    chair  = s["rb2_chair"]
    next_m = s["rb2_next_mtg"].strftime("%d %B %Y") if s["rb2_next_mtg"] else "TBD"
    disc   = s["rb2_disc_en"] or s["rb2_disc_ar"] or ""

    subject = f"SUBJECT: Follow-up on {sub_en} — {co} / Ministry of Investment | {mtg}"

    intro = (
        f"Dear {co} Team,\n\n"
        f"Thank you for the productive meeting held on {mtg}, chaired by {chair}. "
        "We appreciate your continued partnership and valued the open dialogue. "
        "Please find below a summary of the key discussion points and agreed action items."
    )

    disc_block = "KEY DISCUSSION POINTS\n" + "-" * 50
    for ln in _parse_lines(disc):
        disc_block += f"\n  • {ln}"
    if not _parse_lines(disc):
        disc_block += "\n  (See attached meeting minutes for full details)"

    actions = s["rb2_actions"]
    valid   = _valid_actions(actions)
    act_block = "\nAGREED ACTION ITEMS\n" + "-" * 50
    if valid.empty:
        act_block += "\n  (No action items recorded)"
    else:
        act_block += f"\n{'#':<4} {'Action':<55} {'Owner':<20} {'Due':<14} {'Priority'}"
        act_block += f"\n{'-'*4} {'-'*55} {'-'*20} {'-'*14} {'-'*8}"
        for i, (_, row) in enumerate(valid.iterrows(), 1):
            en = row.get("Action (EN)") or row.get("Action (AR)", "")
            due = str(row.get("Due Date", "TBD")) if row.get("Due Date") else "TBD"
            act_block += f"\n{i:<4} {en[:54]:<55} {row.get('Assigned To','TBD'):<20} {due:<14} {row.get('Priority','')}"

    closing = (
        f"\nNEXT STEPS\n{'-'*50}\n"
        f"Our next meeting is tentatively scheduled for {next_m}. "
        "We look forward to continued progress and remain committed to supporting your operations in the Kingdom.\n\n"
        "Please do not hesitate to reach out should you require any further information.\n\n"
        "Best regards,\n\n"
        f"{chair}\n"
        "Ministry of Investment of Saudi Arabia\n"
        "وزارة الاستثمار — المملكة العربية السعودية"
    )

    return "\n\n".join([subject, intro, disc_block, act_block, closing])


# ── Helpers ───────────────────────────────────────────────────────────────────

def _valid_actions(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    mask = (
        df["Action (EN)"].notna() & (df["Action (EN)"] != "") |
        df["Action (AR)"].notna() & (df["Action (AR)"] != "")
    )
    return df[mask]


def _parse_lines(text: str) -> list:
    if not text:
        return []
    return [ln.strip().lstrip("•●-– ") for ln in text.split("\n") if ln.strip()]


def _map_prio(p: str) -> str:
    return {"Very High": "High", "High": "High", "Medium": "Medium", "Low": "Low"}.get(p, "Medium")


def _get_excel_sheets(raw: bytes) -> list:
    try:
        wb = openpyxl.load_workbook(io.BytesIO(raw))
        return wb.sheetnames
    except Exception:
        return []


def _investor_id(investors: pd.DataFrame, company: str) -> str:
    if investors.empty or "Company Name" not in investors.columns:
        return ""
    m = investors[investors["Company Name"] == company]
    return str(m.iloc[0].get("Investor ID", "")) if not m.empty else ""


def _next_act_id(df: pd.DataFrame, offset: int = 0) -> str:
    if df.empty or "Action ID" not in df.columns:
        return f"ACT-{(1 + offset):03d}"
    nums = []
    for v in df["Action ID"].dropna():
        try:
            nums.append(int(str(v).split("-")[-1]))
        except ValueError:
            pass
    return f"ACT-{(max(nums) + 1 + offset):03d}"
