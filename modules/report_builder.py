
# Report Builder — Ministry of Investment
# Upload Arabic Word + Excel tracker → extract actions → update Excel + generate letter + sync CRM

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

# ─── Constants ────────────────────────────────────────────────────────────────

_PRIORITIES  = ["Very High", "High", "Medium", "Low"]
_ENGAGE_OPTS = ["Support", "Opportunity", "Challenge", "Follow-up", "Action", "Administrative"]
_PRIORITY_AR = {"مهم جدا": "Very High", "مهم": "High", "متوسط": "Medium", "عادي": "Low", "مستمر": "Ongoing"}
_OWNER_AR    = {
    "ساره السيد":   "Sara Al-Sayed",
    "القطاع المالي": "Financial Sector Dept.",
    "SIPA":         "SIPA",
    "زباد الجهيمان": "Zabad Al-Juhaiman",
    "دانا الجاربو":  "Dana Aljarbu",
    "خالد الخطاف":   "Khaled Alkhattaf",
}

_EMPTY_ACTIONS = pd.DataFrame(columns=[
    "Action (AR)", "Action (EN)", "Assigned To", "Type",
    "Priority", "Due Date", "Due Text", "Remarks",
])

_STATUS_FILLS = {
    "Completed":    "E2EFDA",
    "Inprogress":   "FFF2CC",
    "In Progress":  "FFF2CC",
    "Not Started":  "F2F2F2",
    "Blocked":      "FCE4D6",
    "Cancelled":    "EDEDED",
}
_BADGE_CSS = {
    "Very High":  "background:#fee2e2;color:#dc2626",
    "High":       "background:#fee2e2;color:#dc2626",
    "Medium":     "background:#fef3c7;color:#d97706",
    "Low":        "background:#f3f4f6;color:#6b7280",
    "Completed":  "background:#d1fae5;color:#065f46",
    "Not Started":"background:#f3f4f6;color:#6b7280",
    "Inprogress": "background:#fef3c7;color:#d97706",
    "In Progress":"background:#fef3c7;color:#d97706",
    "Blocked":    "background:#fee2e2;color:#dc2626",
}

# ─── Translation helpers ──────────────────────────────────────────────────────

def _translate_to_en(text: str) -> str:
    if not text or not text.strip():
        return text
    latin  = sum(1 for c in text if c.isascii() and c.isalpha())
    arabic = sum(1 for c in text if '؀' <= c <= 'ۿ')
    if latin > arabic:
        return text
    try:
        from deep_translator import GoogleTranslator
        return GoogleTranslator(source="ar", target="en").translate(text) or text
    except Exception:
        return text


def _translate_list(texts: list) -> list:
    if not texts:
        return texts
    try:
        from deep_translator import GoogleTranslator
        tr = GoogleTranslator(source="ar", target="en")
        return [tr.translate(s) or s for s in texts]
    except Exception:
        return texts


# ─── Session state ────────────────────────────────────────────────────────────

def _init():
    defaults = {
        "rb_parsed":            None,
        "rb_actions":           _EMPTY_ACTIONS.copy(),
        "rb_company":           "",
        "rb_recipient":         "",
        "rb_arm":               "",
        "rb_exec_rm":           "",
        "rb_date":              date.today().strftime("%d %B %Y"),
        "rb_subject_ar":        "",
        "rb_subject_en":        "",
        "rb_location":          "المقر الرئيسي – وزارة الاستثمار",
        "rb_chair":             "معالي الوزير",
        "rb_next_meeting_text": "",
        "rb_attendees":         "",
        "rb_disc_ar":           "",
        "rb_output_lang":       "English",
        "rb_excel_bytes":       None,
        "rb_result":            None,
        "rb_manual_opps":       "",
        "rb_detected_opps":     [],
        "rb_last_opp_company":  "",
        # Excel-derived per-company data
        "rb_excel_companies":   [],
        "rb_excel_header_data": {},
        "rb_last_co_fill":      "",
        "rb_rep_position":      "",
        "rb_rep_email":         "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ─── CSS ──────────────────────────────────────────────────────────────────────

def _inject_css():
    st.markdown("""
    <style>
    .rb-section{font-size:13px;font-weight:500;color:#6b7280;text-transform:uppercase;
                letter-spacing:.05em;margin-bottom:1rem}
    .rb-num{display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;
            border-radius:50%;background:#217141;color:#fff;font-size:12px;font-weight:500;
            margin-right:8px;vertical-align:middle;flex-shrink:0}
    .rb-badge{display:inline-flex;align-items:center;padding:2px 10px;border-radius:20px;
              font-size:11px;font-weight:500;background:#f3f4f6;color:#6b7280}
    .rb-action-table{width:100%;border-collapse:collapse;font-size:12px;margin-top:.5rem}
    .rb-action-table th{background:#217141;color:#fff;padding:7px 10px;text-align:left;
                        font-size:11px;font-weight:500}
    .rb-action-table td{padding:6px 10px;border:0.5px solid #e5e7eb;vertical-align:top}
    .rb-action-table tr:nth-child(even) td{background:#f9fafb}
    .rb-pk{display:inline-flex;align-items:center;padding:2px 8px;border-radius:12px;
           font-size:10px;font-weight:500}
    </style>
    """, unsafe_allow_html=True)


# ─── Main render ──────────────────────────────────────────────────────────────

def render(dfs: dict, lang: str):
    _inject_css()
    _init()

    # Header
    hc, bc = st.columns([5, 1])
    hc.markdown("**Report Builder — Ministry of Investment**")
    bc.markdown('<div style="text-align:right"><span class="rb-badge">Reusable workflow</span></div>',
                unsafe_allow_html=True)
    st.markdown('<hr style="margin:.25rem 0 1.25rem;border:none;border-top:0.5px solid #e5e7eb">',
                unsafe_allow_html=True)

    # ── Step 1: Upload ────────────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown('<p class="rb-section">Step 1 — Upload your files</p>', unsafe_allow_html=True)
        u1, u2 = st.columns(2)
        with u1:
            st.markdown(
                '<span class="rb-num">1</span>'
                '<strong style="font-size:12px">Arabic meeting minutes (.docx)</strong>',
                unsafe_allow_html=True)
            st.caption("Standard Ministry of Investment meeting template")
            word_file = st.file_uploader("word", type=["docx","doc"],
                                         key="rb_word_up", label_visibility="collapsed")
        with u2:
            st.markdown(
                '<span class="rb-num">2</span>'
                '<strong style="font-size:12px">Action Item Tracker (.xlsx)</strong>',
                unsafe_allow_html=True)
            st.caption("V5 tracker — only relevant sheet will be updated")
            excel_file = st.file_uploader("excel", type=["xlsx","xls"],
                                          key="rb_excel_up", label_visibility="collapsed")

    # Auto-parse Word
    if word_file is not None:
        raw = word_file.read()
        with st.spinner("Parsing Arabic document and translating action items…"):
            parsed = _parse_word(raw)
        if parsed:
            s = st.session_state
            s["rb_parsed"]            = parsed
            s["rb_company"]           = parsed.get("company", "")
            s["rb_date"]              = (
                parsed["date"].strftime("%d %B %Y") if parsed.get("date")
                else date.today().strftime("%d %B %Y")
            )
            s["rb_subject_ar"]        = parsed.get("subject_ar", "")
            s["rb_subject_en"]        = parsed.get("subject_en", "")
            s["rb_location"]          = parsed.get("location", "المقر الرئيسي – وزارة الاستثمار")
            s["rb_chair"]             = parsed.get("chair", "معالي الوزير")
            s["rb_next_meeting_text"] = parsed.get("next_meeting_text", "")
            s["rb_attendees"]         = parsed.get("attendees", "")
            s["rb_disc_ar"]           = parsed.get("discussion_ar", "")
            if parsed.get("action_items"):
                items    = parsed["action_items"]
                ar_texts = [i.get("Action (AR)", "") for i in items]
                en_texts = _translate_list(ar_texts)
                for i, en in enumerate(en_texts):
                    if not items[i].get("Action (EN)"):
                        items[i]["Action (EN)"] = en
                # Also translate "Due Text" fields
                due_texts = [i.get("Due Text", "") for i in items]
                due_en    = _translate_list([d for d in due_texts if d])
                due_idx = 0
                for i in items:
                    if i.get("Due Text"):
                        i["Due Text EN"] = due_en[due_idx] if due_idx < len(due_en) else i["Due Text"]
                        due_idx += 1
                s["rb_actions"] = pd.DataFrame(items)
            n = len(parsed.get("action_items", []))
            st.success(f"✅ Parsed — {n} action item(s) found and translated to English.")

    if excel_file is not None:
        st.session_state["rb_excel_bytes"] = excel_file.read()
        # Parse ALL per-company header data in one pass
        hdr_data = _read_all_header_data(st.session_state["rb_excel_bytes"])
        st.session_state["rb_excel_header_data"] = hdr_data
        xl_companies = list(hdr_data.keys())
        st.session_state["rb_excel_companies"] = xl_companies
        st.info(
            f"Excel loaded — {len(xl_companies)} company sheet(s): "
            f"{', '.join(xl_companies[:6])}"
        )
        # Auto-select + auto-fill first company if none yet selected
        if xl_companies:
            first_co = xl_companies[0]
            if not st.session_state.get("rb_company"):
                st.session_state["rb_company"] = first_co
            cur_co = st.session_state["rb_company"]
            # Fill fields for the currently-selected (or first) company
            _fill_company_fields(hdr_data, cur_co)
            st.session_state["rb_last_co_fill"] = cur_co
        # Detect potential opportunities from action items (keeps company per item)
        _det = _detect_opps_from_actions(st.session_state["rb_excel_bytes"])
        st.session_state["rb_detected_opps"] = _det
        # Don't pre-fill text area here — we filter by company in Step 2

    # ── Step 2: Configure ─────────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown('<p class="rb-section">Step 2 — Configure output</p>', unsafe_allow_html=True)
        investors    = dfs.get("Investor Master", pd.DataFrame())
        crm_companies = sorted(investors["Company Name"].dropna().unique().tolist()) \
            if not investors.empty and "Company Name" in investors.columns else []

        # Excel companies take precedence; CRM companies fill the rest
        xl_companies   = st.session_state.get("rb_excel_companies", [])
        company_list   = xl_companies + [c for c in crm_companies if c not in xl_companies]

        r1c1, r1c2 = st.columns(2)
        with r1c1:
            cur_co = st.session_state["rb_company"]
            if company_list:
                idx     = company_list.index(cur_co) if cur_co in company_list else 0
                company = st.selectbox("Company", company_list, index=idx, key="rb_co_sel")
            else:
                company = st.text_input("Company", value=cur_co,
                                        placeholder="e.g. Barclays", key="rb_co_txt")
            st.session_state["rb_company"] = company

            # Auto-fill AM / RM / Rep when company changes
            hdr_data = st.session_state.get("rb_excel_header_data", {})
            if company and company != st.session_state.get("rb_last_co_fill", "") and company in hdr_data:
                _fill_company_fields(hdr_data, company)
                st.session_state["rb_last_co_fill"] = company
                st.rerun()

        with r1c2:
            rep_pos   = st.session_state.get("rb_rep_position", "")
            rep_email = st.session_state.get("rb_rep_email", "")
            st.session_state["rb_recipient"] = st.text_input(
                "Representative name",
                value=st.session_state["rb_recipient"],
                placeholder="e.g. Khalid Al-Dabbagh", key="rb_recip")
            if rep_pos or rep_email:
                st.caption(f"{rep_pos}{'  ·  ' + rep_email if rep_email else ''}")

        r2c1, r2c2 = st.columns(2)
        with r2c1:
            st.session_state["rb_arm"] = st.text_input(
                "ARM (Account Relationship Manager)",
                value=st.session_state["rb_arm"],
                placeholder="e.g. Dana Aljarbu", key="rb_arm_inp")
        with r2c2:
            st.session_state["rb_exec_rm"] = st.text_input(
                "Executive RM (Minister's Office)",
                value=st.session_state["rb_exec_rm"],
                placeholder="e.g. Sara Al-Sayed", key="rb_exec_inp")

        r3c1, r3c2 = st.columns(2)
        with r3c1:
            st.session_state["rb_date"] = st.text_input(
                "Meeting date",
                value=st.session_state["rb_date"],
                placeholder=date.today().strftime("%d %B %Y"), key="rb_date_inp")
        with r3c2:
            st.session_state["rb_output_lang"] = st.selectbox(
                "Output language",
                ["English", "Arabic", "Bilingual (EN + AR)"],
                key="rb_lang_sel")

        # ── Opportunities ─────────────────────────────────────────────────────
        st.markdown("**Opportunities**")
        _cur_company = st.session_state.get("rb_company", "")
        _all_det     = st.session_state.get("rb_detected_opps", [])

        # Filter detected opps to the currently selected company only
        _co_det = [o for o in _all_det if o.get("company", "").lower() in _cur_company.lower()
                   or _cur_company.lower() in o.get("company", "").lower()] if _cur_company else []

        # Show all companies' detected opps grouped (informational)
        if _all_det:
            from collections import defaultdict as _dd
            _by_co: dict = _dd(list)
            for _o in _all_det:
                _by_co[_o["company"]].append(_o["name"])
            with st.expander(f"ℹ️ {len(_all_det)} potential opportunit{'y' if len(_all_det)==1 else 'ies'} detected across all companies — click to review", expanded=False):
                for _co, _names in _by_co.items():
                    st.markdown(f"**{_co}**")
                    for _n in _names:
                        st.markdown(f"  - {_n}")
                st.caption("Action items classified as 'Opportunity' type. The text area below is pre-filled for the selected company only.")

        # Pre-fill text area when company changes and it's currently empty
        _co_names = [o["name"] for o in _co_det]
        _prev_company = st.session_state.get("rb_last_opp_company", "")
        if _cur_company != _prev_company:
            # Company changed — reset text area to this company's detected opps
            prefill = "\n".join(_co_names)
            st.session_state["rb_manual_opps"]       = prefill
            st.session_state["rb_manual_opps_ta"]    = prefill
            st.session_state["rb_last_opp_company"]  = _cur_company

        st.session_state["rb_manual_opps"] = st.text_area(
            f"Opportunities for {_cur_company or 'selected company'} (one per line)",
            value=st.session_state["rb_manual_opps"],
            placeholder="e.g.\nInvestment in fintech ecosystem\nHealthcare matchmaking with InterHealth",
            height=120,
            key="rb_manual_opps_ta",
            help="These will appear in the letter and be synced to the Opportunity Pipeline for this company only.",
        )

    # ── Step 3: Generate ──────────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown('<p class="rb-section">Step 3 — Generate</p>', unsafe_allow_html=True)

        if st.button("▶  Run pipeline — extract actions, update Excel, generate letter & sync CRM",
                     type="primary", use_container_width=True, key="rb_run"):

            s       = st.session_state
            actions = _valid_actions(s["rb_actions"])
            company = s["rb_company"] or "Company"
            arm     = s["rb_arm"] or "Dana Aljarbu"
            exec_rm = s["rb_exec_rm"] or "Sara Al-Sayed"
            recipient = s["rb_recipient"] or f"{company} Team"
            mtg_date  = s["rb_date"] or date.today().strftime("%d %B %Y")

            prog   = st.progress(0)
            status = st.empty()

            def _log(msg: str, pct: int = None):
                status.markdown(f"→ {msg}")
                if pct is not None:
                    prog.progress(pct)

            try:
                _log("Reading parsed data…", 10)

                manual_opps = [
                    ln.strip() for ln in s.get("rb_manual_opps", "").splitlines()
                    if ln.strip()
                ]
                cfg = {
                    "company":           company,
                    "recipient":         recipient,
                    "arm":               arm,
                    "exec_rm":           exec_rm,
                    "meeting_date":      mtg_date,
                    "chair":             s.get("rb_chair", "معالي الوزير"),
                    "next_meeting_text": s.get("rb_next_meeting_text", ""),
                    "disc_en":           s.get("rb_disc_ar", ""),
                    "opportunities":     manual_opps,
                }

                # 1 — Updated Excel tracker
                updated_xl = None
                if s.get("rb_excel_bytes") and not actions.empty:
                    _log(f"Updating Excel tracker for {company}…", 30)
                    try:
                        mtg_d = datetime.strptime(mtg_date, "%d %B %Y").date()
                    except ValueError:
                        mtg_d = date.today()
                    updated_xl = _build_excel(
                        existing_bytes=s["rb_excel_bytes"],
                        company=company,
                        meeting_date=mtg_d,
                        next_meeting=s.get("rb_next_meeting_text") or None,
                        chair=cfg["chair"],
                        actions_df=actions,
                    )

                # 2 — Word letter
                _log("Generating company letter (.docx)…", 55)
                letter_bytes = _build_letter_docx(cfg, actions)

                # 3 — CRM sync (actions)
                _log("Syncing action items to CRM…", 75)
                _sync_actions_to_crm(dfs, actions, company)

                # 4 — Opportunities: Excel header block + detected (per-company) + manual
                opp_count = 0
                opp_items: list[dict] = []
                if s.get("rb_excel_bytes"):
                    _log("Syncing opportunities…", 90)
                    try:
                        _wb_ops   = openpyxl.load_workbook(io.BytesIO(s["rb_excel_bytes"]))
                        # Header-block opps (already carry correct company from sheet name)
                        opp_items = _parse_opps_from_excel(_wb_ops)
                        # Detected action-item opps (carry correct company per sheet)
                        for _do in s.get("rb_detected_opps", []):
                            if _do["name"] not in {o["name"] for o in opp_items}:
                                opp_items.append(_do)
                    except Exception:
                        pass
                # Manual text-area entries → current company only
                existing_names = {o["name"] for o in opp_items}
                for mo in manual_opps:
                    if mo not in existing_names:
                        opp_items.append({"company": company, "name": mo})
                        existing_names.add(mo)
                if opp_items:
                    _sync_opps_to_crm(dfs, opp_items, company)
                    opp_count = len(opp_items)

                prog.progress(100)
                status.success(
                    f"✓ Pipeline complete — 3 outputs ready"
                    + (f"  |  {opp_count} opportunit{'y' if opp_count == 1 else 'ies'} synced" if opp_count else "")
                )

                st.session_state["rb_result"] = {
                    "actions":      actions,
                    "xl_bytes":     updated_xl,
                    "letter":       letter_bytes,
                    "company":      company,
                    "meeting_date": mtg_date,
                    "opp_count":    opp_count,
                }

            except Exception as e:
                status.error(f"Error: {e}")
                import traceback
                st.code(traceback.format_exc())

    # ── Output section ────────────────────────────────────────────────────────
    result = st.session_state.get("rb_result")
    if result:
        _render_output(result)


# ─── Output renderer ──────────────────────────────────────────────────────────

def _render_output(result: dict):
    st.markdown('<hr style="margin:1.5rem 0;border:none;border-top:0.5px solid #e5e7eb">',
                unsafe_allow_html=True)

    actions    = result["actions"]
    xl_bytes   = result["xl_bytes"]
    letter     = result["letter"]
    company    = result["company"]
    mtg_date   = result["meeting_date"]
    fname_base = f"{company.replace(' ','_')}_{mtg_date.replace(' ','_')}"

    # ── Action items table ────────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown("#### ✅ Extracted action items")
        if not actions.empty:
            _render_action_table(actions)
        else:
            st.info("No action items extracted from the uploaded document.")

    # ── Downloads ─────────────────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown("#### ⬇ Downloads")
        dc1, dc2 = st.columns(2)

        with dc1:
            if xl_bytes:
                st.download_button(
                    "📊  Download updated Excel tracker",
                    data=xl_bytes,
                    file_name=f"ActionTracker_{fname_base}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, key="rb_dl_xl",
                )
            else:
                st.info("Excel tracker not generated (upload a tracker file in Step 1)")

        with dc2:
            if letter:
                st.download_button(
                    "📄  Download company letter (.docx)",
                    data=letter,
                    file_name=f"Letter_{fname_base}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True, key="rb_dl_letter",
                )

        opp_count = result.get("opp_count", 0)
        if xl_bytes and letter:
            msg = "✓ Both files are ready. Action items have been synced to CRM."
            if opp_count:
                msg += f" **{opp_count} opportunit{'y' if opp_count == 1 else 'ies'}** added to Opportunity Pipeline."
            st.success(msg)


def _render_action_table(df: pd.DataFrame):
    rows_html = ""
    for i, (_, row) in enumerate(df.iterrows()):
        en     = (row.get("Action (EN)") or row.get("Action (AR)", "")).strip()
        owner  = str(row.get("Assigned To", "") or "")
        prio   = str(row.get("Priority", "Medium") or "Medium")
        due    = str(row.get("Due Text EN", "") or row.get("Due Text", "") or
                     row.get("Due Date", "") or "TBD")
        status = str(row.get("Status", "Not Started") or "Not Started")
        rmk    = str(row.get("Remarks", "") or "")

        p_css  = _BADGE_CSS.get(prio,   "background:#f3f4f6;color:#6b7280")
        s_css  = _BADGE_CSS.get(status, "background:#f3f4f6;color:#6b7280")

        rows_html += (
            f"<tr>"
            f"<td style='font-weight:500;text-align:center;width:36px'>{i+1}</td>"
            f"<td>{en}</td><td>{owner}</td>"
            f"<td><span class='rb-pk' style='{p_css}'>{prio}</span></td>"
            f"<td>{due}</td>"
            f"<td><span class='rb-pk' style='{s_css}'>{status}</span></td>"
            f"<td style='font-size:11px;color:#6b7280'>{rmk}</td>"
            f"</tr>"
        )

    st.markdown(f"""
    <div style="overflow-x:auto">
    <table class="rb-action-table">
      <thead><tr><th>#</th><th>Action Item</th><th>Owner</th>
      <th>Priority</th><th>Timeline</th><th>Status</th><th>Remarks</th></tr></thead>
      <tbody>{rows_html}</tbody>
    </table></div>
    """, unsafe_allow_html=True)


# ─── Word parser ──────────────────────────────────────────────────────────────

def _parse_word(file_bytes: bytes) -> dict:
    """Parse Arabic Ministry meeting minutes. Handles the standard 4-table format."""
    try:
        doc = docx.Document(io.BytesIO(file_bytes))
    except Exception:
        return {}

    result = {
        "company": "", "date": None, "subject_ar": "", "subject_en": "",
        "location": "المقر الرئيسي – وزارة الاستثمار",
        "chair": "معالي الوزير", "priority": "Very High",
        "next_meeting_text": "",
        "attendees": "", "discussion_ar": "", "action_items": [],
    }

    tables = []
    for tbl in doc.tables:
        rows = []
        for row in tbl.rows:
            cells = [c.text.strip() for c in row.cells]
            rows.append(cells)
        tables.append(rows)

    def _classify_table(rows: list) -> str:
        """Identify table type by header content, not by position."""
        header_text = " ".join(c for row in rows[:2] for c in row)
        if any(m in header_text for m in ["التوجيه", "المهمة", "مسؤول", "الأولوية", "الموعد النهائي", "الإجراء"]):
            return "actions"
        if any(m in header_text for m in ["الاسم", "المسمى", "الجهة", "حضر", "المشاركون", "التوقيع"]):
            return "attendees"
        if len(rows) <= 5:
            for row in rows:
                for cell in row:
                    if len(cell) > 120:
                        return "discussion"
        if any(m in header_text for m in ["أبرز ما تم مناقشته", "نقاط النقاش", "مناقشة"]):
            return "discussion"
        return "metadata"

    # Classify all tables by content (order-independent)
    _classified: dict[str, list] = {}
    for _t in tables:
        _kind = _classify_table(_t)
        if _kind not in _classified:
            _classified[_kind] = _t

    def _dedup(row: list) -> list:
        seen, out = set(), []
        for c in row:
            if c not in seen:
                seen.add(c)
                out.append(c)
        return out

    def _parse_date(s: str):
        for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(s.strip(), fmt).date()
            except ValueError:
                pass
        m = re.match(r"(\d{1,2})[-/](\d{1,2})[-/](\d{4})", s.strip())
        if m:
            try:
                return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except Exception:
                pass
        return None

    # ── Table 0: Meeting metadata ─────────────────────────────────────────────
    # Row 0: headers (الموضوع, الموقع, اليوم, التاريخ)
    # Row 1: values  (subject, location, day, date)
    # Row 2: headers (وقت, برئاسة, الأولوية, الاجتماع القادم)
    # Row 3: values  (time, chair, priority, next_meeting)
    if tables:
        meta = tables[0]
        if len(meta) >= 2:
            row1 = _dedup(meta[1])
            if row1:
                result["subject_ar"] = row1[0]
            if len(row1) >= 2:
                result["location"] = row1[1]
            # Date: last unique cell that looks like a date
            for cell in reversed(row1):
                d = _parse_date(cell)
                if d:
                    result["date"] = d
                    break
            # Extract company from subject
            known = ["Barclays", "BlackRock", "Brookfield", "Goldman", "HSBC",
                     "JPMorgan", "Morgan Stanley", "UBS", "Citi", "Deutsche",
                     "Allianz", "Lazard", "Blackstone", "Carlyle", "KKR"]
            for name in known:
                if name.lower() in row1[0].lower():
                    result["company"] = name
                    result["subject_en"] = f"Latest Updates — {name}"
                    break

        if len(meta) >= 4:
            row3 = _dedup(meta[3])
            if len(row3) >= 2:
                result["chair"] = row3[1]
            if len(row3) >= 3:
                prio_map = {"مهم جدا": "Very High", "مهم": "High", "متوسط": "Medium", "عادي": "Low"}
                result["priority"] = prio_map.get(row3[2], "High")
            if len(row3) >= 4:
                result["next_meeting_text"] = row3[3]

    # ── Discussion table ──────────────────────────────────────────────────────
    _disc_rows = _classified.get("discussion")
    if _disc_rows is None:
        # Positional fallback: table 1, only if it wasn't claimed as actions/attendees
        _claimed = {id(_t) for _t in _classified.values()}
        for _t in tables[1:]:
            if id(_t) not in _claimed:
                _disc_rows = _t
                break
    if _disc_rows:
        full_cell = _disc_rows[0][0] if _disc_rows[0] else ""
        marker = "أبرز ما تم مناقشته:"
        if marker in full_cell:
            disc = full_cell.split(marker, 1)[-1].strip()
        else:
            disc = full_cell
        result["discussion_ar"] = disc

    # ── Action items table ────────────────────────────────────────────────────
    _act_rows = _classified.get("actions")
    if _act_rows is None and len(tables) >= 3:
        _act_rows = tables[2]
    if _act_rows:
        act_tbl = _act_rows
        for row in act_tbl[1:]:
            cells = _dedup(row)
            # Skip row if empty or only numbers
            non_empty = [c for c in cells if c and not re.match(r"^\d+$", c)]
            if not non_empty:
                continue
            action_ar = non_empty[0]
            if action_ar in ("م", "التوجيه / المهمة", "Action Item", ""):
                continue

            owner_ar  = non_empty[1] if len(non_empty) > 1 else ""
            prio_ar   = non_empty[2] if len(non_empty) > 2 else ""
            due_ar    = non_empty[3] if len(non_empty) > 3 else ""

            owner_en = _OWNER_AR.get(owner_ar, owner_ar)
            prio_en  = _PRIORITY_AR.get(prio_ar, "High")
            if prio_en == "Ongoing":
                prio_en = "High"

            result["action_items"].append({
                "Action (AR)": action_ar,
                "Action (EN)": "",
                "Assigned To": owner_en,
                "Type":        "Action",
                "Priority":    prio_en,
                "Due Date":    None,
                "Due Text":    due_ar,
                "Due Text EN": "",
                "Remarks":     "",
                "Status":      "Not Started",
            })

    # ── Attendees table ───────────────────────────────────────────────────────
    _att_rows = _classified.get("attendees")
    if _att_rows is None and len(tables) >= 4:
        _att_rows = tables[3]
    if _att_rows:
        att_lines = []
        for row in _att_rows[1:]:
            cells = [c for c in _dedup(row) if c and not re.match(r"^\d+$", c)]
            if len(cells) >= 2:
                att_lines.append(f"{cells[0]} | {cells[1]}")
            elif cells:
                att_lines.append(cells[0])
        result["attendees"] = "\n".join(att_lines)

    return result


# ─── Letter builder (Word doc) ────────────────────────────────────────────────

def _build_letter_docx(cfg: dict, actions_df: pd.DataFrame) -> bytes:
    """Generate a Word letter matching the Barclays_Email_Letter.docx template format."""
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    co         = cfg["company"]
    mtg        = cfg["meeting_date"]
    arm        = cfg["arm"] or "Dana Aljarbu"
    exec_rm    = cfg["exec_rm"] or "Sara Al-Sayed"
    recipient  = cfg["recipient"] or f"{co} Team"
    next_text  = cfg.get("next_meeting_text", "")
    yr_mon     = date.today().strftime("%Y-%m")

    arm_email  = (f"{arm.split()[0].lower()}."
                  f"{''.join(arm.split()[1:]).lower()}@misa.gov.sa")
    exec_email = (f"{exec_rm.split()[0].lower()}."
                  f"{''.join(exec_rm.split()[1:]).lower()}@misa.gov.sa")

    doc = docx.Document()
    sec = doc.sections[0]
    sec.left_margin   = Inches(1.0)
    sec.right_margin  = Inches(1.0)
    sec.top_margin    = Inches(0.75)
    sec.bottom_margin = Inches(0.75)

    def _rgb(h: str) -> RGBColor:
        h = h.lstrip("#")
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

    def _cell_bg(cell, hex_color: str):
        tcPr = cell._tc.get_or_add_tcPr()
        shd  = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), hex_color.lstrip("#").upper())
        tcPr.append(shd)

    def _run(para, text: str, bold=False, size_pt=11.0, color="1C1C1C"):
        r = para.add_run(text)
        r.bold            = bold
        r.font.size       = Pt(size_pt)
        r.font.color.rgb  = _rgb(color)
        return r

    def _para(text="", bold=False, size_pt=11.0, color="1C1C1C", align=WD_ALIGN_PARAGRAPH.LEFT):
        p = doc.add_paragraph()
        p.alignment = align
        if text:
            _run(p, text, bold=bold, size_pt=size_pt, color=color)
        return p

    # ── Ministry header table ─────────────────────────────────────────────────
    hdr = doc.add_table(rows=2, cols=1)
    hdr.style = "Table Grid"

    c0 = hdr.rows[0].cells[0]
    _cell_bg(c0, "217141")
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p0, "Ministry of Investment  |  وزارة الاستثمار",
         bold=True, size_pt=12, color="FFFFFF")

    c1 = hdr.rows[1].cells[0]
    _cell_bg(c1, "1A5C3F")
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p1, "Minister's Office  —  مكتب الوزير", size_pt=10, color="C9974A")

    _para()  # blank line

    # ── Date + Reference ──────────────────────────────────────────────────────
    ref_p = doc.add_paragraph()
    _run(ref_p, "Date: ",          bold=True, size_pt=10, color="555555")
    _run(ref_p, f"{mtg}",          bold=False, size_pt=10, color="555555")
    _run(ref_p, f"     |     Ref: MISA / {co} / ARM / {yr_mon}",
         bold=False, size_pt=10, color="999999")

    _para()

    # ── Greeting ──────────────────────────────────────────────────────────────
    greet = doc.add_paragraph()
    _run(greet, f"Dear {recipient},", bold=True, size_pt=12, color="1C1C1C")
    _para()

    # ── Body ──────────────────────────────────────────────────────────────────
    _para("I hope this message finds you well.", size_pt=11)
    _para()

    p3 = doc.add_paragraph()
    _run(p3, "On behalf of the ")
    _run(p3, "Ministry of Investment", bold=True)
    _run(p3, ", we would like to express our sincere appreciation for our strong and "
         "growing partnership with ")
    _run(p3, co, bold=True)
    _run(p3, ", and reaffirm our commitment to supporting your continued growth and "
         "success in the Kingdom following our meeting on ")
    _run(p3, mtg, bold=True)
    _run(p3, ".")
    _para()

    p4 = doc.add_paragraph()
    _run(p4, "In this regard, we are pleased to confirm that ")
    _run(p4, arm, bold=True)
    _run(p4, " will be leading the account team as the ")
    _run(p4, f"Account Relationship Manager (ARM) for {co}", bold=True)
    _run(p4, " — serving as your primary point of contact for all operational matters "
         "and coordination. ")
    _run(p4, exec_rm, bold=True)
    _run(p4, " from the Minister's Office will act as the ")
    _run(p4, "Executive Relationship Manager", bold=True)
    _run(p4, " for any topics related to the Minister.")
    _para()

    _para("As a progress update, please find below the current status of our agreed "
          "action items and key workstreams:", size_pt=11)
    _para()

    # ── Action items table ────────────────────────────────────────────────────
    if not actions_df.empty:
        tbl = doc.add_table(rows=1, cols=7)
        tbl.style = "Table Grid"

        # Header row
        hdrs = ["#", "Action Item", "Description / Update",
                "MISA Owner", "Priority", "Timeline", "Status"]
        widths = [0.25, 1.3, 2.7, 1.2, 0.65, 1.05, 0.85]
        for j, (cell, hdr_txt) in enumerate(zip(tbl.rows[0].cells, hdrs)):
            _cell_bg(cell, "217141")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _run(p, hdr_txt, bold=True, size_pt=9, color="FFFFFF")
            try:
                from docx.oxml import OxmlElement as OE
                tc   = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcW  = OE("w:tcW")
                tcW.set(qn("w:w"), str(int(widths[j] * 1440)))
                tcW.set(qn("w:type"), "dxa")
                tcPr.append(tcW)
            except Exception:
                pass

        for i, (_, row) in enumerate(actions_df.iterrows()):
            en     = (row.get("Action (EN)") or row.get("Action (AR)", "")).strip()
            owner  = str(row.get("Assigned To", "") or "")
            prio   = str(row.get("Priority", "Medium") or "Medium")
            due    = str(row.get("Due Text EN", "") or row.get("Due Text", "") or
                         row.get("Due Date", "") or "TBD")
            status = str(row.get("Status", "Not Started") or "Not Started")

            short  = (en[:55].rsplit(" ", 1)[0] + "…") if len(en) > 55 else en
            bg     = "FFFFFF" if i % 2 == 0 else "F5F5F5"
            st_bg  = _STATUS_FILLS.get(status, "F2F2F2")

            data_row = tbl.add_row()
            vals_bgs = [
                (str(i + 1), bg), (short, bg), (en, bg),
                (owner, bg), (prio, bg), (due, bg), (status, st_bg),
            ]
            for j, (val, cell_bg) in enumerate(vals_bgs):
                cell = data_row.cells[j]
                _cell_bg(cell, cell_bg)
                p = cell.paragraphs[0]
                if j == 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _run(p, val, size_pt=9, color="1C1C1C")

    _para()

    # ── Opportunities section ─────────────────────────────────────────────────
    opps_list = cfg.get("opportunities", [])
    if opps_list:
        p_opp_hdr = doc.add_paragraph()
        _run(p_opp_hdr, "Opportunities Identified:", bold=True, size_pt=11, color="1B5C3F")
        _para()
        for opp_name in opps_list:
            p_opp = doc.add_paragraph(style="List Bullet")
            _run(p_opp, opp_name, size_pt=10, color="1C1C1C")
        _para()

    # ── Closing ───────────────────────────────────────────────────────────────
    p7 = doc.add_paragraph()
    _run(p7, "We remain fully committed to supporting ")
    _run(p7, co, bold=True)
    _run(p7, " in growing its presence in the Kingdom and to progressing all "
         "workstreams with the attention they deserve.")
    _para()

    upcoming = f"our upcoming meeting in {next_text}" if next_text else "our upcoming meeting"
    _para(f"We look forward to our continued collaboration and to {upcoming}.")
    _para()

    p9 = doc.add_paragraph()
    _run(p9, "Please feel free to contact ")
    _run(p9, arm, bold=True)
    _run(p9, " directly for all day-to-day coordination (CC: ")
    _run(p9, exec_rm, bold=True)
    _run(p9, " for any ministerial matters).")
    _para()

    _para("Yours sincerely,")
    _para()
    _para()

    # ── Signature ─────────────────────────────────────────────────────────────
    sig_n = doc.add_paragraph()
    _run(sig_n, exec_rm, bold=True, size_pt=11.5, color="217141")

    sig_t = doc.add_paragraph()
    _run(sig_t, "Executive Relationship Manager  |  Minister's Office",
         size_pt=10.5, color="555555")

    sig_m = doc.add_paragraph()
    _run(sig_m, "Ministry of Investment  |  Kingdom of Saudi Arabia",
         size_pt=10.5, color="555555")

    sig_e = doc.add_paragraph()
    _run(sig_e, "E: ",         size_pt=10, color="888888")
    _run(sig_e, exec_email,    size_pt=10, color="1A5276")
    _run(sig_e, "   |   ARM: ", size_pt=10, color="888888")
    _run(sig_e, arm_email,     size_pt=10, color="1A5276")

    _para()

    # ── Footer table ──────────────────────────────────────────────────────────
    ftr = doc.add_table(rows=1, cols=1)
    ftr.style = "Table Grid"
    fc = ftr.rows[0].cells[0]
    _cell_bg(fc, "217141")
    fp = fc.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(fp, "Ministry of Investment  |  Kingdom of Saudi Arabia  |  www.misa.gov.sa",
         size_pt=9, color="FFFFFF")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ─── Excel builder ────────────────────────────────────────────────────────────

def _build_excel(existing_bytes, company, meeting_date, next_meeting, chair, actions_df) -> bytes:
    if existing_bytes:
        wb = openpyxl.load_workbook(io.BytesIO(existing_bytes))
    else:
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    # Match existing sheet with fuzzy name (handles typos like "Barclyes")
    target_sheet = None
    for sname in wb.sheetnames:
        if company.lower() in sname.lower() or sname.lower() in company.lower():
            target_sheet = sname
            break
    sheet_name = target_sheet or f"Action Items {company}"[:31]

    if sheet_name in wb.sheetnames:
        ws       = wb[sheet_name]
        last_row = 20
        for row in ws.iter_rows(min_row=21, values_only=True):
            if any(v is not None for v in row):
                last_row += 1
        start_row = last_row + 1
    else:
        ws = wb.create_sheet(sheet_name)
        _write_sheet_header(ws, company, meeting_date, next_meeting, chair)
        start_row = 21

    normal_font = Font(size=10)
    center_al   = Alignment(horizontal="center", vertical="center")
    right_al    = Alignment(horizontal="right",  vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )

    for i, (_, row) in enumerate(actions_df.iterrows()):
        r       = start_row + i
        due_val = None
        due_raw = row.get("Due Date")
        if due_raw and str(due_raw) not in ("NaT", "None", ""):
            due_val = due_raw

        prio_map = {"Very High": "Very High", "High": "High", "Medium": "Medium", "Low": "Low"}
        prio_val = prio_map.get(str(row.get("Priority", "Medium")), "Medium")
        due_text = str(row.get("Due Text EN", "") or row.get("Due Text", "") or "")

        ws.cell(r, 7,  i + 1)
        ws.cell(r, 8,  row.get("Action (EN)") or row.get("Action (AR)", ""))
        ws.cell(r, 9,  row.get("Assigned To", ""))
        ws.cell(r, 10, row.get("Type", "Action"))
        ws.cell(r, 11, meeting_date)
        ws.cell(r, 12, due_val or due_text)
        ws.cell(r, 13, prio_val)
        ws.cell(r, 14, 0)
        ws.cell(r, 15, "Not Started")
        ws.cell(r, 16, row.get("Remarks", ""))

        for col in range(7, 17):
            cell = ws.cell(r, col)
            cell.font      = normal_font
            cell.border    = thin_border
            cell.alignment = right_al if col == 8 else center_al

        ws.cell(r, 15).fill = PatternFill("solid", fgColor="D9D9D9")
        ws.row_dimensions[r].height = 40

    if start_row == 21:
        widths = {7: 6, 8: 50, 9: 18, 10: 18, 11: 12, 12: 16, 13: 12, 14: 10, 15: 14, 16: 30}
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

    ws.cell(13, 11, "Outreach").font = white_font
    ws.cell(13, 11).fill = green_fill
    ws.cell(13, 15, "Last Updated").font = white_font
    ws.cell(13, 15).fill = green_fill
    ws.cell(13, 16, str(meeting_date)).font = white_font
    ws.cell(13, 16).fill = gold_fill
    ws.cell(14, 11, "Manager").font = white_font
    ws.cell(14, 11).fill = green_fill
    ws.cell(14, 12, str(chair)).font = Font(bold=True, size=10)
    ws.cell(16, 15, "Next Meeting").font = white_font
    ws.cell(16, 15).fill = green_fill
    ws.cell(16, 16, str(next_meeting) if next_meeting else "TBD").font = gold_font
    ws.cell(16, 16).fill = gold_fill

    for r in range(13, 20):
        ws.row_dimensions[r].height = 18

    headers = ["ID", "Action Item", "Assigned to", "Type of Engagement",
               "Start Date", "Due Date", "Priority", "Progress", "Status", "Remarks"]
    for j, h in enumerate(headers):
        cell = ws.cell(20, 7 + j, h)
        cell.fill      = green_fill
        cell.font      = white_font
        cell.alignment = center_al
    ws.row_dimensions[20].height = 22
    ws.freeze_panes = "G21"


# ─── Opportunity extraction & sync ───────────────────────────────────────────

def _parse_opps_from_excel(wb) -> list:
    """
    Read all sheets in the workbook and extract opportunity names from the
    header block rows 14-19, column H — the cells the user fills in to list
    major opportunities arising from each meeting.
    Returns a list of dicts: {company, name}.
    """
    _SKIP_VALS = {"opportunity", "opportunities", "type", "n/a", "none", ""}
    opps = []
    for sname in wb.sheetnames:
        # Derive company name by stripping "Action Items" prefix
        co = sname.strip()
        for prefix in ("Action Items", "Actions", "Action Item"):
            if co.lower().startswith(prefix.lower()):
                co = co[len(prefix):].strip()
                break

        ws = wb[sname]
        for row_idx in range(14, 20):       # rows 14-19 (header block area)
            cell_val = ws.cell(row_idx, 8).value   # column H
            if cell_val is None:
                continue
            val = str(cell_val).strip()
            if not val:
                continue
            if val.lower() in _SKIP_VALS:
                continue
            # Skip row-label artefacts like "4 Manager", "¦ Last Updated"
            if val[:2] in ("4 ", "¦ ", "¹ ") or val[0] in "¦¹":
                continue
            opps.append({"company": co or sname, "name": val})
    return opps


def _next_opp_id(df: pd.DataFrame) -> str:
    if df.empty or "Opportunity ID" not in df.columns:
        return "OPP-001"
    nums = []
    for v in df["Opportunity ID"].dropna():
        try:
            nums.append(int(str(v).split("-")[-1]))
        except ValueError:
            pass
    return f"OPP-{(max(nums) + 1 if nums else 1):03d}"


def _sync_opps_to_crm(dfs: dict, opp_items: list, default_company: str):
    """Add extracted opportunities to Opportunity Pipeline (deduplicates by name)."""
    if not opp_items:
        return
    investors = dfs.get("Investor Master", pd.DataFrame())
    existing  = dfs.get("Opportunity Pipeline", pd.DataFrame())
    new_rows  = []

    for item in opp_items:
        company = item.get("company") or default_company
        name    = item["name"]
        # Skip if already in pipeline
        if not existing.empty and "Opportunity Name" in existing.columns:
            if (existing["Opportunity Name"].astype(str).str.strip() == name).any():
                continue
        tmp    = pd.concat([existing, pd.DataFrame(new_rows)], ignore_index=True) if new_rows else existing
        new_id = _next_opp_id(tmp)
        inv_id = _investor_id(investors, company)
        new_row = {
            "Opportunity ID":     new_id,
            "Investor ID":        inv_id,
            "Company Name":       company,
            "Opportunity Name":   name,
            "Sector":             "",
            "Opportunity Stage":  "Exploration",
            "Opportunity Status": "Active",
            "Opportunity Type":   "Opportunity",
            "Opportunity Source": "Excel Tracker Import",
            "Last Updated":       date.today(),
        }
        new_rows.append(new_row)
        existing = pd.concat([existing, pd.DataFrame([new_row])], ignore_index=True)

    if new_rows:
        dfs["Opportunity Pipeline"] = existing
        save_session(dfs)


# ─── CRM sync ────────────────────────────────────────────────────────────────

def _sync_actions_to_crm(dfs: dict, actions: pd.DataFrame, company: str):
    if actions.empty or not company:
        return
    investors = dfs.get("Investor Master", pd.DataFrame())
    inv_id    = _investor_id(investors, company)
    existing  = dfs.get("Action Items", pd.DataFrame())
    new_rows  = []

    for _, row in actions.iterrows():
        desc = (row.get("Action (EN)") or row.get("Action (AR)", "")).strip()
        if not desc:
            continue
        if not existing.empty and "Action Description" in existing.columns:
            if (existing["Action Description"].astype(str).str.strip() == desc).any():
                continue
        tmp  = pd.concat([existing, pd.DataFrame(new_rows)], ignore_index=True) if new_rows else existing
        new_id = _next_act_id(tmp)
        new_rows.append({
            "Action ID":          new_id,
            "Investor ID":        inv_id,
            "Company Name":       company,
            "Action Description": desc,
            "Assigned To":        str(row.get("Assigned To", "") or ""),
            "Sector":             "",
            "Type of Engagement": str(row.get("Type", "Action") or "Action"),
            "Start Date":         date.today(),
            "Due Date":           row.get("Due Date") if row.get("Due Date") and pd.notna(row.get("Due Date")) else None,
            "Priority":           str(row.get("Priority", "Medium") or "Medium"),
            "Progress":           "0%",
            "Status":             "Not Started",
            "Escalation Flag":    "None",
            "Remarks":            str(row.get("Remarks", "") or ""),
            "Last Updated":       date.today(),
        })
        existing = pd.concat([existing, pd.DataFrame([new_rows[-1]])], ignore_index=True)

    if new_rows:
        dfs["Action Items"] = existing
        save_session(dfs)


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _valid_actions(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    en_col = df.get("Action (EN)", pd.Series(dtype=str))
    ar_col = df.get("Action (AR)", pd.Series(dtype=str))
    mask   = (en_col.notna() & (en_col != "")) | (ar_col.notna() & (ar_col != ""))
    return df[mask].reset_index(drop=True)


def _get_excel_sheets(raw: bytes) -> list:
    try:
        wb = openpyxl.load_workbook(io.BytesIO(raw))
        return wb.sheetnames
    except Exception:
        return []


def _read_all_header_data(raw_bytes: bytes) -> dict:
    """
    Parse every 'Action Items [Company]' sheet and extract the header block fields.
    Returns: {company: {"am", "rm", "rep", "position", "email", "last_updated", "next_meeting"}}
    The header block uses adjacent label→value pairs anywhere in rows 12-18.
    """
    _skip = {"", "nan", "none", "n/a"}
    result: dict = {}
    try:
        wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
        for ws in wb.worksheets:
            co = ws.title.strip()
            for prefix in ("Action Items", "Actions", "Action Item"):
                if co.lower().startswith(prefix.lower()):
                    co = co[len(prefix):].strip()
                    break
            if not co:
                continue
            info: dict = {
                "am": "", "rm": "", "rep": "", "position": "",
                "email": "", "last_updated": "", "next_meeting": "",
            }
            for r in range(12, 19):
                for c in range(1, 20):
                    raw_lbl = str(ws.cell(row=r, column=c).value or "").strip()
                    lbl     = raw_lbl.lstrip("4¦¹1234567890 ").strip()
                    val_cell = ws.cell(row=r, column=c + 1).value
                    val      = str(val_cell or "").strip()
                    if not lbl or val.lower() in _skip:
                        continue
                    lu = lbl.upper()
                    if re.search(r"^A[-.]?M$", lbl, re.IGNORECASE):
                        info["am"] = val
                    elif lu == "RM":
                        info["rm"] = val
                    elif lu == "REP":
                        info["rep"] = val
                    elif "POSIT" in lu:
                        info["position"] = val
                    elif lu == "EMAIL":
                        info["email"] = val
                    elif "LAST" in lu and ("UPDATE" in lu or "UPDAT" in lu):
                        try:
                            info["last_updated"] = (
                                val_cell.strftime("%d %B %Y")
                                if hasattr(val_cell, "strftime") else val[:10]
                            )
                        except Exception:
                            info["last_updated"] = val[:10]
                    elif "NEXT" in lu and "MEET" in lu:
                        try:
                            info["next_meeting"] = (
                                val_cell.strftime("%d %B %Y")
                                if hasattr(val_cell, "strftime") else val[:10]
                            )
                        except Exception:
                            info["next_meeting"] = val[:10]
            result[co] = info
    except Exception:
        pass
    return result


def _fill_company_fields(hdr_data: dict, company: str):
    """Push parsed header fields for `company` into Streamlit session state."""
    info = hdr_data.get(company, {})
    if info.get("am"):
        st.session_state["rb_arm"]     = info["am"]
        st.session_state["rb_arm_inp"] = info["am"]
    if info.get("rm"):
        st.session_state["rb_exec_rm"]  = info["rm"]
        st.session_state["rb_exec_inp"] = info["rm"]
    if info.get("rep"):
        st.session_state["rb_recipient"] = info["rep"]
        st.session_state["rb_recip"]     = info["rep"]
    if info.get("next_meeting"):
        st.session_state["rb_next_meeting_text"] = info["next_meeting"]
    st.session_state["rb_rep_position"] = info.get("position", "")
    st.session_state["rb_rep_email"]    = info.get("email", "")


def _read_am_rm_from_excel(raw_bytes: bytes) -> tuple[str, str]:
    """Legacy helper — kept for backward compatibility."""
    data = _read_all_header_data(raw_bytes)
    for info in data.values():
        if info.get("am") or info.get("rm"):
            return info.get("am", ""), info.get("rm", "")
    return "", ""


def _detect_opps_from_actions(raw_bytes: bytes) -> list[dict]:
    """
    Scan all sheets for action item rows whose Type of Engagement = 'Opportunity'.
    Returns list of {company, name} dicts — one per unique action item.
    """
    results: list[dict] = []
    _seen: set[str] = set()
    try:
        wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
        for ws in wb.worksheets:
            # Derive company name from sheet name
            co = ws.title.strip()
            for prefix in ("Action Items", "Actions", "Action Item"):
                if co.lower().startswith(prefix.lower()):
                    co = co[len(prefix):].strip()
                    break

            hdr_row = None
            for r in range(15, 25):
                for c in range(1, 20):
                    if str(ws.cell(row=r, column=c).value or "").strip() == "Action Item":
                        hdr_row = r
                        break
                if hdr_row:
                    break
            if not hdr_row:
                continue

            col_map: dict[str, int] = {}
            for c in range(1, 20):
                v = str(ws.cell(row=hdr_row, column=c).value or "").strip()
                if v:
                    col_map[v] = c

            action_col = col_map.get("Action Item", 8)
            type_col   = next(
                (col_map[k] for k in col_map if "Type" in k and "Engagement" in k),
                None,
            )
            if not type_col:
                continue

            for r in range(hdr_row + 1, hdr_row + 200):
                ai = str(ws.cell(row=r, column=action_col).value or "").strip()
                if not ai:
                    break
                eng = str(ws.cell(row=r, column=type_col).value or "").strip().lower()
                if eng == "opportunity" and ai not in _seen:
                    _seen.add(ai)
                    results.append({"company": co or ws.title, "name": ai})
    except Exception:
        pass
    return results


def _investor_id(investors: pd.DataFrame, company: str) -> str:
    if investors.empty or "Company Name" not in investors.columns:
        return ""
    m = investors[investors["Company Name"] == company]
    return str(m.iloc[0].get("Investor ID", "")) if not m.empty else ""


def _next_act_id(df: pd.DataFrame) -> str:
    if df.empty or "Action ID" not in df.columns:
        return "ACT-001"
    nums = []
    for v in df["Action ID"].dropna():
        try:
            nums.append(int(str(v).split("-")[-1]))
        except ValueError:
            pass
    return f"ACT-{(max(nums) + 1 if nums else 1):03d}"
