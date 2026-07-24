
# Report Builder — Ministry of Investment
# Upload Arabic Word + Excel tracker → extract actions → update Excel + generate letter + sync CRM

import io
import re
import numpy as np
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
        "rb_ppt_bytes":         None,
        "rb_ppt_company_info":  {},
        # Arabic minutes generator
        "rb_ar_summary":        "",
        "rb_ar_company":        "",
        "rb_ar_date":           date.today().strftime("%d/%m/%Y"),
        "rb_ar_location":       "المقر الرئيسي – وزارة الاستثمار",
        "rb_ar_chair":          "معالي الوزير",
        "rb_ar_priority":       "مهم جدا",
        "rb_ar_next_mtg":       "",
        "rb_ar_arm":            "",
        "rb_ar_exec_rm":        "",
        "rb_ar_api_key":        "",
        "rb_ar_content":        None,
        "rb_ar_docx_bytes":     None,
        "rb_ar_doc_key":        "",
        "rb_en_docx_bytes":     None,
        "rb_am_letter_bytes":   None,
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
        u1, u2, u3 = st.columns(3)
        with u1:
            st.markdown(
                '<span class="rb-num">1</span>'
                '<strong style="font-size:12px">Arabic meeting minutes (any format)</strong>',
                unsafe_allow_html=True)
            st.caption("Standard Ministry meeting template — optional if using PPT brief · accepts .docx, .pdf, .png, .jpg")
            word_file = st.file_uploader("word",
                                         type=["docx", "doc", "pdf", "png", "jpg", "jpeg", "webp"],
                                         key="rb_word_up", label_visibility="collapsed")
        with u2:
            st.markdown(
                '<span class="rb-num">2</span>'
                '<strong style="font-size:12px">Company Brief (any format)</strong>',
                unsafe_allow_html=True)
            st.caption("Account manager brief — extracts all action items by sector · accepts .pptx, .pdf, .png, .jpg")
            ppt_file = st.file_uploader("ppt",
                                        type=["pptx", "ppt", "pdf", "png", "jpg", "jpeg", "webp"],
                                        key="rb_ppt_up", label_visibility="collapsed")
        with u3:
            st.markdown(
                '<span class="rb-num">3</span>'
                '<strong style="font-size:12px">Action Item Tracker (any format)</strong>',
                unsafe_allow_html=True)
            st.caption("V5 tracker — existing rows updated, new rows appended · accepts .xlsx, .pdf, .png, .jpg")
            excel_file = st.file_uploader("excel",
                                          type=["xlsx", "xls", "pdf", "png", "jpg", "jpeg", "webp"],
                                          key="rb_excel_up", label_visibility="collapsed")

    # Auto-parse minutes file (only when a new file is uploaded)
    if word_file is not None:
        _wd_key = f"{word_file.name}_{word_file.size}"
        if _wd_key == st.session_state.get("rb_word_file_key", ""):
            word_file = None  # already parsed — skip
        else:
            st.session_state["rb_word_file_key"] = _wd_key
    if word_file is not None:
        raw = word_file.read()
        _ext = word_file.name.rsplit(".", 1)[-1].lower() if "." in word_file.name else ""
        _MIME_MAP = {
            "pdf": "application/pdf",
            "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "webp": "image/webp",
        }
        _is_docx = _ext in ("docx", "doc")
        with st.spinner("Parsing document and extracting action items…"):
            try:
                if _is_docx:
                    parsed = _parse_word(raw)
                else:
                    _api_key = st.session_state.get("rb_ar_api_key", "").strip() or \
                               __import__("os").environ.get("ANTHROPIC_API_KEY", "")
                    _mime = _MIME_MAP.get(_ext, "application/pdf")
                    parsed = _parse_minutes_via_claude(raw, _mime, _api_key) if _api_key else {}
                    if not _api_key:
                        st.warning("Add your Anthropic API key in the Arabic Minutes Generator section to parse non-DOCX files.")
            except Exception as _pe:
                parsed = {}
                st.error(f"Could not parse document: {_pe}")
        if not parsed:
            st.warning(
                "⚠️ No recognisable meeting-minutes structure found in the uploaded document. "
                "For DOCX files, the standard 4-table Arabic Ministry format is expected. "
                "You can still run the pipeline using action items loaded via the "
                "**Arabic Minutes Generator** below."
            )
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

    # Auto-parse company brief (any format)
    if ppt_file is not None:
        raw_ppt = ppt_file.read()
        _ppt_ext = ppt_file.name.rsplit(".", 1)[-1].lower() if "." in ppt_file.name else ""
        _MIME_MAP_PPT = {
            "pdf": "application/pdf",
            "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "webp": "image/webp",
        }
        _is_pptx = _ppt_ext in ("pptx", "ppt")
        with st.spinner("Parsing company brief…"):
            if _is_pptx:
                pptx_data = _parse_company_brief_pptx(raw_ppt)
            else:
                _ppt_api = st.session_state.get("rb_ar_api_key", "").strip() or \
                           __import__("os").environ.get("ANTHROPIC_API_KEY", "")
                if _ppt_api:
                    _ppt_mime = _MIME_MAP_PPT.get(_ppt_ext, "application/pdf")
                    pptx_data = _parse_brief_via_claude(raw_ppt, _ppt_mime, _ppt_api)
                else:
                    pptx_data = {"error": "API key required — add it in the Arabic Minutes Generator section"}
        if pptx_data.get("error"):
            st.error(f"Brief parse error: {pptx_data['error']}")
        else:
            st.session_state["rb_ppt_bytes"]        = raw_ppt
            st.session_state["rb_ppt_company_info"] = pptx_data.get("company_info", {})
            # Auto-set company name from PPT if not already set
            if pptx_data.get("company") and not st.session_state.get("rb_company"):
                st.session_state["rb_company"] = pptx_data["company"]
            # Populate action items from PPT slides
            if pptx_data.get("action_items"):
                act_rows = []
                for item in pptx_data["action_items"]:
                    act_rows.append({
                        "Action (AR)": "",
                        "Action (EN)": item.get("Action (EN)", ""),
                        "Sector":      item.get("Sector", ""),
                        "Assigned To": item.get("Assigned To", ""),
                        "Type":        item.get("Type", "Action"),
                        "Priority":    item.get("Priority", "Medium"),
                        "Start Date":  item.get("Start Date", ""),
                        "Due Date":    item.get("Due Date", ""),
                        "Progress":    item.get("Progress", 0),
                        "Status":      item.get("Status", "Not Started"),
                        "Remarks":     item.get("Remarks", ""),
                        "Due Text":    "",
                        "Due Text EN": "",
                    })
                st.session_state["rb_actions"] = pd.DataFrame(act_rows)
            n_acts   = len(pptx_data.get("action_items", []))
            sectors  = pptx_data.get("sectors", [])
            sec_str  = ", ".join(sectors) if sectors else "—"
            co_info  = pptx_data.get("company_info", {})
            info_parts = []
            if co_info.get("aum"):
                info_parts.append(f"AUM {co_info['aum']}")
            if co_info.get("sector"):
                info_parts.append(f"Sector: {co_info['sector']}")
            info_note = "  ·  " + "  ·  ".join(info_parts) if info_parts else ""
            st.success(
                f"✅ PPT parsed — **{n_acts}** action item(s) across "
                f"**{len(sectors)}** sector(s): {sec_str}{info_note}"
            )
            # Show company info card if slide 2 data was found
            if co_info:
                with st.expander("📋 Company info extracted from PPT", expanded=False):
                    ci_cols = st.columns(3)
                    fields = [
                        ("Sector",       co_info.get("sector",        "")),
                        ("HQ",           co_info.get("hq",            "")),
                        ("AUM",          co_info.get("aum",           "")),
                        ("KSA Presence", co_info.get("ksa_presence",  "")),
                        ("Employees",    co_info.get("employees",     "")),
                        ("Website",      co_info.get("website",       "")),
                        ("Rep",          co_info.get("rep_name",      "")),
                        ("Email",        co_info.get("email",         "")),
                        ("Phone",        co_info.get("phone",         "")),
                    ]
                    for k, (label, val) in enumerate(fields):
                        if val:
                            ci_cols[k % 3].markdown(f"**{label}:** {val}")

    if excel_file is not None:
        _xl_key = f"{excel_file.name}_{excel_file.size}"
        _xl_ext = excel_file.name.rsplit(".", 1)[-1].lower() if "." in excel_file.name else ""
        _MIME_MAP_XL = {
            "pdf": "application/pdf",
            "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "webp": "image/webp",
        }
        _is_xlsx = _xl_ext in ("xlsx", "xls")
        if _xl_key != st.session_state.get("rb_excel_file_key", ""):
            # New file — parse and cache everything once
            st.session_state["rb_excel_file_key"] = _xl_key
            raw_xl = excel_file.read()
            if _is_xlsx:
                st.session_state["rb_excel_bytes"] = raw_xl
                hdr_data = _read_all_header_data(raw_xl)
                st.session_state["rb_excel_header_data"] = hdr_data
                xl_companies = list(hdr_data.keys())
                st.session_state["rb_excel_companies"] = xl_companies
                st.info(
                    f"Tracker loaded — {len(xl_companies)} company sheet(s): "
                    f"{', '.join(xl_companies[:6])}"
                )
                if xl_companies:
                    first_co = xl_companies[0]
                    if not st.session_state.get("rb_company"):
                        st.session_state["rb_company"] = first_co
                    cur_co = st.session_state["rb_company"]
                    _fill_company_fields(hdr_data, cur_co)
                    st.session_state["rb_last_co_fill"] = cur_co
                    # Auto-load action items from the selected company's sheet
                    _xl_acts = _load_excel_actions_for_company(raw_xl, cur_co)
                    if not _xl_acts.empty:
                        st.session_state["rb_actions"] = _xl_acts
                        st.info(
                            f"✅ Loaded **{len(_xl_acts)}** action item(s) from "
                            f"*{cur_co}* sheet."
                        )
                _det = _detect_opps_from_actions(raw_xl)
                st.session_state["rb_detected_opps"] = _det
            else:
                # Non-Excel: extract action items via Claude
                _xl_api = st.session_state.get("rb_ar_api_key", "").strip() or \
                          __import__("os").environ.get("ANTHROPIC_API_KEY", "")
                if _xl_api:
                    with st.spinner("Extracting action items from document…"):
                        _xl_mime = _MIME_MAP_XL.get(_xl_ext, "application/pdf")
                        _xl_data = _parse_tracker_via_claude(raw_xl, _xl_mime, _xl_api)
                    if _xl_data.get("error"):
                        st.error(f"Tracker parse error: {_xl_data['error']}")
                    else:
                        if not st.session_state.get("rb_company") and _xl_data.get("company"):
                            st.session_state["rb_company"] = _xl_data["company"]
                        items = _xl_data.get("action_items", [])
                        if items:
                            st.session_state["rb_actions"] = pd.DataFrame(items)
                            st.info(f"Tracker loaded — {len(items)} action item(s) extracted")
                        else:
                            st.warning("No action items found in the uploaded document.")
                else:
                    st.warning("Add your Anthropic API key in the Arabic Minutes Generator section to parse non-Excel files.")
        else:
            # Same file already processed — just show the cached info
            xl_companies = st.session_state.get("rb_excel_companies", [])
            if xl_companies:
                st.info(
                    f"Tracker loaded — {len(xl_companies)} company sheet(s): "
                    f"{', '.join(xl_companies[:6])}"
                )

    # ── Step 2: Configure ─────────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown('<p class="rb-section">Step 2 — Configure output</p>', unsafe_allow_html=True)
        investors    = dfs.get("Investor Master", pd.DataFrame())
        crm_companies = sorted(investors["Company Name"].dropna().unique().tolist()) \
            if not investors.empty and "Company Name" in investors.columns else []

        # Excel companies take precedence; CRM companies fill the rest
        xl_companies   = st.session_state.get("rb_excel_companies", [])
        company_list   = xl_companies + [c for c in crm_companies if c not in xl_companies]

        # Company selector — full width
        cur_co = st.session_state["rb_company"]
        if company_list:
            idx     = company_list.index(cur_co) if cur_co in company_list else 0
            company = st.selectbox("Company", company_list, index=idx, key="rb_co_sel")
        else:
            company = st.text_input("Company", value=cur_co,
                                    placeholder="e.g. Barclays", key="rb_co_txt")
        st.session_state["rb_company"] = company

        # Auto-fill AM / RM / date when company changes
        hdr_data = st.session_state.get("rb_excel_header_data", {})
        try:
            if company and company != st.session_state.get("rb_last_co_fill", "") and company in hdr_data:
                _fill_company_fields(hdr_data, company)
                st.session_state["rb_last_co_fill"] = company
                # Reload action items from the tracker for the newly selected company.
                # Exception: if a MoM was already parsed for this same company, keep
                # those actions — don't let the (often-empty) Excel sheet wipe them.
                _raw_xl_co = st.session_state.get("rb_excel_bytes")
                _parsed_co = (st.session_state.get("rb_parsed") or {}).get("company", "")
                _mom_matches = bool(
                    _parsed_co and (
                        _parsed_co.lower() in company.lower() or
                        company.lower() in _parsed_co.lower()
                    )
                )
                if not _mom_matches:
                    if _raw_xl_co:
                        _co_acts = _load_excel_actions_for_company(_raw_xl_co, company)
                        st.session_state["rb_actions"] = _co_acts if not _co_acts.empty else pd.DataFrame()
                    else:
                        st.session_state["rb_actions"] = pd.DataFrame()
                st.rerun()
        except Exception as _fill_err:
            st.warning(f"⚠️ Auto-fill error (non-fatal): {_fill_err}")

        r2c1, r2c2 = st.columns(2)
        with r2c1:
            st.session_state["rb_arm"] = st.text_input(
                "Account Manager (AM)",
                value=st.session_state["rb_arm"],
                placeholder="e.g. Waleed AlShehri", key="rb_arm_inp")
        with r2c2:
            st.session_state["rb_exec_rm"] = st.text_input(
                "Relationship Manager (RM) — Minister's Office",
                value=st.session_state["rb_exec_rm"],
                placeholder="e.g. Sultana Alsegaih", key="rb_exec_inp")

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
            try:
                actions = _valid_actions(s.get("rb_actions", _EMPTY_ACTIONS.copy()))
            except Exception as _ae:
                st.error(f"Could not read action items: {_ae}")
                actions = _EMPTY_ACTIONS.copy()
            company = s.get("rb_company") or "Company"
            arm     = s.get("rb_arm", "") or ""
            exec_rm = s.get("rb_exec_rm", "") or ""
            recipient = s.get("rb_recipient") or f"{company} Team"
            mtg_date  = s.get("rb_date") or date.today().strftime("%d %B %Y")

            prog   = st.progress(0)
            status = st.empty()
            status.markdown(f"→ Starting pipeline for **{company}** — "
                            f"{len(actions)} action item(s) loaded…")

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
                _ppt_mode  = bool(s.get("rb_ppt_bytes"))
                if s.get("rb_excel_bytes"):
                    _log(f"Updating Excel tracker for {company}…", 30)
                    try:
                        mtg_d = datetime.strptime(mtg_date, "%d %B %Y").date()
                    except ValueError:
                        mtg_d = date.today()
                    if not actions.empty:
                        if _ppt_mode:
                            updated_xl, _n_upd, _n_add = _merge_pptx_actions_to_excel(
                                s["rb_excel_bytes"], company,
                                actions.to_dict(orient="records"),
                            )
                            _log(
                                f"Excel updated — {_n_upd} row(s) updated, "
                                f"{_n_add} new row(s) added.", 50
                            )
                        else:
                            updated_xl = _build_excel(
                                existing_bytes=s["rb_excel_bytes"],
                                company=company,
                                meeting_date=mtg_d,
                                next_meeting=s.get("rb_next_meeting_text") or None,
                                chair=cfg["chair"],
                                actions_df=actions,
                            )
                            _log("Excel tracker updated.", 50)
                    else:
                        # No new action items — return the existing tracker unchanged
                        updated_xl = s["rb_excel_bytes"]
                        _log("No new action items — returning existing Excel tracker.", 50)

                # 2 — Word letter (company)
                _log("Generating company letter (.docx)…", 55)
                letter_bytes = _build_letter_docx(cfg, actions)

                # 2b — Internal AM letter
                _log("Generating internal AM letter (.docx)…", 62)
                am_letter_bytes = _build_internal_am_letter_docx(cfg, actions)
                st.session_state["rb_am_letter_bytes"] = am_letter_bytes

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
                    f"✓ Pipeline complete — 3 outputs ready (Excel + Company Letter + AM Letter)"
                    + (f"  |  {opp_count} opportunit{'y' if opp_count == 1 else 'ies'} synced" if opp_count else "")
                )

                st.session_state["rb_result"] = {
                    "actions":      actions,
                    "xl_bytes":     updated_xl,
                    "letter":       letter_bytes,
                    "am_letter":    am_letter_bytes,
                    "company":      company,
                    "meeting_date": mtg_date,
                    "opp_count":    opp_count,
                }

            except Exception as e:
                prog.progress(0)
                import traceback as _tb
                status.error(f"Pipeline error at last step → {e}")
                st.code(_tb.format_exc())

    # ── Output section ────────────────────────────────────────────────────────
    result = st.session_state.get("rb_result")
    if result:
        _render_output(result)

    # ══════════════════════════════════════════════════════════════════════════
    # Arabic Meeting Minutes Generator
    # ══════════════════════════════════════════════════════════════════════════
    st.markdown('<hr style="margin:2rem 0 1rem;border:none;border-top:2px solid #e5e7eb">',
                unsafe_allow_html=True)
    st.markdown("**Arabic Meeting Minutes Generator — محضر الاجتماع**")
    st.caption("Upload an English meeting summary → Claude translates and fills the Arabic Ministry template")

    with st.container(border=True):
        st.markdown('<p class="rb-section">Step A — English meeting summary</p>',
                    unsafe_allow_html=True)
        a1, a2 = st.columns([1, 1])
        with a1:
            ar_doc = st.file_uploader(
                "English summary (.docx)",
                type=["docx", "doc"],
                key="rb_ar_doc_up",
                label_visibility="collapsed",
                help="Upload the English meeting notes/summary .docx file",
            )
            # Parse file and pre-set widget key BEFORE text_area is instantiated
            if ar_doc is not None:
                _ar_doc_key = f"{ar_doc.name}_{ar_doc.size}"
                if _ar_doc_key != st.session_state.get("rb_ar_doc_key", ""):
                    st.session_state["rb_ar_doc_key"] = _ar_doc_key
                    try:
                        import docx as _dx
                        _raw_ar = ar_doc.read()
                        _doc    = _dx.Document(io.BytesIO(_raw_ar))
                        # Extract paragraphs
                        _txt = "\n".join(p.text for p in _doc.paragraphs if p.text.strip())
                        # Also extract table text so Claude has full context (action items live in tables)
                        for _tbl in _doc.tables:
                            for _trow in _tbl.rows:
                                _cells = " | ".join(c.text.strip() for c in _trow.cells if c.text.strip())
                                if _cells:
                                    _txt += "\n" + _cells
                        st.session_state["rb_ar_summary_ta"] = _txt
                        st.session_state["rb_ar_summary"]    = _txt
                        # Also try to parse as Arabic minutes directly (action items are in tables)
                        _pre = _parse_word(_raw_ar)
                        if _pre.get("action_items"):
                            _pre_items  = _pre["action_items"]
                            _pre_ar     = [i.get("Action (AR)", "") for i in _pre_items]
                            _pre_en     = _translate_list(_pre_ar)
                            for _pi, _pen in enumerate(_pre_en):
                                if not _pre_items[_pi].get("Action (EN)"):
                                    _pre_items[_pi]["Action (EN)"] = _pen
                            _due_ar = [i.get("Due Text", "") for i in _pre_items]
                            _due_en = _translate_list([d for d in _due_ar if d])
                            _due_idx = 0
                            for _pi in _pre_items:
                                if _pi.get("Due Text"):
                                    _pi["Due Text EN"] = _due_en[_due_idx] if _due_idx < len(_due_en) else _pi["Due Text"]
                                    _due_idx += 1
                            st.session_state["rb_actions"] = pd.DataFrame(_pre_items)
                            if not st.session_state.get("rb_company") and _pre.get("company"):
                                st.session_state["rb_company"] = _pre["company"]
                            st.success(
                                f"✅ Extracted {len(_txt.split())} words — "
                                f"**{len(_pre_items)} action item(s) pre-loaded** into the pipeline."
                            )
                        else:
                            st.success(f"✅ Extracted {len(_txt.split())} words — text loaded below.")
                    except Exception as _e:
                        st.error(f"Could not read .docx: {_e}")
            st.caption("Extracted text (edit if needed):")
            st.session_state["rb_ar_summary"] = st.text_area(
                "Meeting summary text",
                height=180,
                key="rb_ar_summary_ta",
                label_visibility="collapsed",
                placeholder="Upload a .docx above, or paste English meeting notes here…",
            )
        with a2:
            st.markdown("**Anthropic API Key**")
            st.session_state["rb_ar_api_key"] = st.text_input(
                "API key",
                value=st.session_state["rb_ar_api_key"],
                type="password",
                key="rb_ar_key_inp",
                label_visibility="collapsed",
                placeholder="sk-ant-…  (or set ANTHROPIC_API_KEY env var)",
                help="Get your key at console.anthropic.com",
            )
            st.caption("Your key is never stored — used only for this session.")

    with st.container(border=True):
        st.markdown('<p class="rb-section">Step B — Meeting metadata</p>',
                    unsafe_allow_html=True)
        b1, b2, b3 = st.columns(3)
        with b1:
            _ar_co_list = (
                st.session_state.get("rb_excel_companies", [])
                or [st.session_state.get("rb_company", "")]
            )
            _ar_co_list = [c for c in _ar_co_list if c]  # drop blanks
            _ar_cur = (
                st.session_state.get("rb_ar_company")
                or st.session_state.get("rb_company", "")
            )
            if _ar_co_list:
                _ar_idx = _ar_co_list.index(_ar_cur) if _ar_cur in _ar_co_list else 0
                st.session_state["rb_ar_company"] = st.selectbox(
                    "Company", _ar_co_list, index=_ar_idx, key="rb_ar_co")
            else:
                st.session_state["rb_ar_company"] = st.text_input(
                    "Company", value=_ar_cur, key="rb_ar_co")
            st.session_state["rb_ar_date"] = st.text_input(
                "Meeting date", value=st.session_state["rb_ar_date"], key="rb_ar_dt",
                placeholder="DD/MM/YYYY")
        with b2:
            st.session_state["rb_ar_location"] = st.text_input(
                "Location (Arabic)", value=st.session_state["rb_ar_location"], key="rb_ar_loc")
            st.session_state["rb_ar_chair"] = st.text_input(
                "Chaired by (Arabic)", value=st.session_state["rb_ar_chair"], key="rb_ar_ch")
        with b3:
            st.session_state["rb_ar_priority"] = st.selectbox(
                "Priority", ["مهم جدا", "مهم", "متوسط", "عادي"],
                index=["مهم جدا", "مهم", "متوسط", "عادي"].index(
                    st.session_state["rb_ar_priority"]
                    if st.session_state["rb_ar_priority"] in ["مهم جدا","مهم","متوسط","عادي"] else "مهم جدا"
                ),
                key="rb_ar_pri")
            st.session_state["rb_ar_next_mtg"] = st.text_input(
                "Next meeting", value=st.session_state["rb_ar_next_mtg"], key="rb_ar_nx",
                placeholder="e.g. Q3 2026 or leave blank")

        b4, b5 = st.columns(2)
        with b4:
            st.session_state["rb_ar_arm"] = st.text_input(
                "Account Manager (AM)", value=st.session_state.get("rb_arm") or st.session_state["rb_ar_arm"],
                key="rb_ar_arm_inp")
        with b5:
            st.session_state["rb_ar_exec_rm"] = st.text_input(
                "Exec RM name", value=st.session_state.get("rb_exec_rm") or st.session_state["rb_ar_exec_rm"],
                key="rb_ar_exec_inp")

    with st.container(border=True):
        st.markdown('<p class="rb-section">Step C — Generate Arabic minutes</p>',
                    unsafe_allow_html=True)

        # Pre-flight: check anthropic is installed
        try:
            import anthropic as _ant_check  # noqa: F401
            _ant_ok = True
        except ImportError:
            _ant_ok = False
            st.warning(
                "⚠️ The `anthropic` package is not installed. "
                "Open a new PowerShell in the project folder and run:\n\n"
                "```\nvenv\\Scripts\\pip install anthropic\n```\n\n"
                "Then restart the CRM."
            )

        if _ant_ok and st.button("▶  Extract, translate & generate Arabic .docx",
                     type="primary", use_container_width=True, key="rb_ar_run"):
            _s    = st.session_state
            _text = _s["rb_ar_summary"].strip()
            _key  = _s["rb_ar_api_key"].strip() or __import__("os").environ.get("ANTHROPIC_API_KEY", "")
            if not _text:
                st.error("Please upload a document or paste meeting text in Step A.")
            elif not _key:
                st.error("Please enter your Anthropic API key in Step A.")
            else:
                _cfg = {
                    "company":    _s["rb_ar_company"],
                    "date":       _s["rb_ar_date"],
                    "location":   _s["rb_ar_location"],
                    "chair":      _s["rb_ar_chair"],
                    "priority":   _s["rb_ar_priority"],
                    "next_mtg":   _s["rb_ar_next_mtg"],
                    "arm":        _s["rb_ar_arm"],
                    "exec_rm":    _s["rb_ar_exec_rm"],
                }
                _status = st.empty()
                _prog   = st.progress(0)
                _prog.progress(10)
                with st.spinner("Calling Claude API — this takes 15–30 seconds…"):
                    _status.markdown("→ Sending to Claude for extraction and translation…")
                    _prog.progress(30)
                    _content = _extract_via_claude(_text, _cfg, _key)
                if "error" in _content:
                    _prog.empty()
                    _status.error(f"Claude API error: {_content['error']}")
                else:
                    _status.markdown("→ Generating Arabic Word document…")
                    _prog.progress(80)
                    _docx_bytes = _build_arabic_minutes_docx(_cfg, _content)
                    _prog.progress(100)
                    _status.success("✓ Arabic meeting minutes ready — download below.")
                    _s["rb_ar_content"]    = _content
                    _s["rb_ar_docx_bytes"] = _docx_bytes
                    # Also generate English minutes automatically
                    try:
                        _en_docx = _build_english_minutes_docx({
                            "company":      _s.get("rb_ar_company", ""),
                            "meeting_date": _s.get("rb_ar_date", ""),
                            "arm":          _s.get("rb_ar_arm", ""),
                            "exec_rm":      _s.get("rb_ar_exec_rm", ""),
                        }, _content)
                        _s["rb_en_docx_bytes"] = _en_docx
                    except Exception:
                        _s["rb_en_docx_bytes"] = None

                    # Push extracted action items into rb_actions so the main
                    # pipeline can write them to the Excel tracker
                    _prio_map = {
                        "مهم جدا": "Very High", "مهم": "High",
                        "متوسط": "Medium",      "عادي": "Low",
                    }
                    _ar_rows = []
                    for _ai in _content.get("action_items", []):
                        _ar_rows.append({
                            "Action (EN)": _ai.get("task_en", ""),
                            "Action (AR)": _ai.get("task", ""),
                            "Assigned To": _ai.get("owner", ""),
                            "Priority":    _prio_map.get(_ai.get("priority", ""), "Medium"),
                            "Due Date":    None,
                            "Due Text":    _ai.get("due", ""),
                            "Remarks":     "",
                        })
                    if _ar_rows:
                        _s["rb_actions"] = pd.DataFrame(_ar_rows)
                        _s["rb_company"] = _s["rb_ar_company"]
                        st.info(f"✅ {len(_ar_rows)} action item(s) loaded into the pipeline — "
                                "scroll up to Step 3 and click Run to write them to Excel.")

    _ar_bytes = st.session_state.get("rb_ar_docx_bytes")
    if _ar_bytes:
        _ar_company  = st.session_state.get("rb_ar_company", "meeting").replace(" ", "_")
        _ar_date_raw = st.session_state.get("rb_ar_date", "").replace("/", "-")
        st.download_button(
            "📄  Download Arabic meeting minutes (.docx)",
            data=_ar_bytes,
            file_name=f"محضر_{_ar_company}_{_ar_date_raw}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="rb_ar_dl",
        )
    # English minutes download (auto-generated alongside Arabic)
    _en_bytes = st.session_state.get("rb_en_docx_bytes")
    if _en_bytes:
        _en_company  = st.session_state.get("rb_ar_company", "meeting").replace(" ", "_")
        _en_date_raw = st.session_state.get("rb_ar_date", "").replace("/", "-")
        st.download_button(
            "📄  Download English meeting minutes (.docx)",
            data=_en_bytes,
            file_name=f"MoM_{_en_company}_{_en_date_raw}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="rb_en_dl",
        )

    # Internal AM letter download
    _ar_content_for_am = st.session_state.get("rb_ar_content", {})
    if _ar_content_for_am and _ar_content_for_am.get("action_items"):
        _am_company = st.session_state.get("rb_ar_company", "Company")
        _am_actions = pd.DataFrame([
            {
                "Action (EN)": (ai.get("task_en") or ai.get("Action (EN)", "")),
                "Action (AR)": ai.get("task", ""),
                "Assigned To": ai.get("owner", ""),
                "Priority":    {"مهم جدا": "Very High", "مهم": "High",
                                "متوسط": "Medium", "عادي": "Low"}.get(
                                    ai.get("priority", ""), "Medium"),
                "Due Text":    ai.get("due", ""),
                "Status":      "Not Started",
            }
            for ai in _ar_content_for_am["action_items"]
        ])
        _am_cfg = {
            "company":      _am_company,
            "meeting_date": st.session_state.get("rb_ar_date", ""),
            "arm":          st.session_state.get("rb_ar_arm", ""),
            "exec_rm":      st.session_state.get("rb_ar_exec_rm", ""),
        }
        if st.button("📧 Generate Internal AM Letter", use_container_width=True,
                     key="rb_am_letter_btn"):
            try:
                _am_bytes = _build_pre_review_am_letter_docx(_am_cfg, _am_actions)
                st.session_state["rb_am_letter_bytes"] = _am_bytes
            except Exception as _e:
                st.error(f"Could not generate AM letter: {_e}")
    _am_letter_bytes = st.session_state.get("rb_am_letter_bytes")
    if _am_letter_bytes:
        _am_co = st.session_state.get("rb_ar_company", "meeting").replace(" ", "_")
        st.download_button(
            "⬇️  Download Internal AM Letter (.docx)",
            data=_am_letter_bytes,
            file_name=f"AM_Internal_{_am_co}_{st.session_state.get('rb_ar_date','').replace('/','_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="rb_am_letter_dl",
        )

    # Preview extracted content
    _ar_content = st.session_state.get("rb_ar_content", {})
    if _ar_content:
        with st.expander("📋 Preview extracted content", expanded=False):
            if _ar_content.get("subject_ar"):
                st.markdown(f"**Subject:** {_ar_content['subject_ar']}")
            if _ar_content.get("discussion_points"):
                st.markdown("**Discussion points:**")
                for pt in _ar_content["discussion_points"]:
                    st.markdown(f"- {pt}")
            if _ar_content.get("action_items"):
                st.markdown(f"**Action items:** {len(_ar_content['action_items'])}")
            if _ar_content.get("attendees"):
                st.markdown(f"**Attendees:** {len(_ar_content['attendees'])}")


# ─── Output renderer ──────────────────────────────────────────────────────────

def _render_output(result: dict):
    st.markdown('<hr style="margin:1.5rem 0;border:none;border-top:0.5px solid #e5e7eb">',
                unsafe_allow_html=True)

    actions    = result["actions"]
    xl_bytes   = result["xl_bytes"]
    letter     = result["letter"]
    am_letter  = result.get("am_letter") or st.session_state.get("rb_am_letter_bytes")
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
        dc1, dc2, dc3 = st.columns(3)

        with dc1:
            if xl_bytes:
                st.download_button(
                    "📊  Excel Tracker",
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
                    "📄  Company Letter",
                    data=letter,
                    file_name=f"Letter_{fname_base}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True, key="rb_dl_letter",
                )

        with dc3:
            if am_letter:
                st.download_button(
                    "📧  Internal AM Letter",
                    data=am_letter,
                    file_name=f"AM_Letter_{fname_base}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True, key="rb_dl_am_letter",
                )
            else:
                st.caption("AM letter will appear here after generation")

        opp_count = result.get("opp_count", 0)
        ready_items = sum([bool(xl_bytes), bool(letter), bool(am_letter)])
        if ready_items > 0:
            msg = f"✓ {ready_items} output{'s' if ready_items > 1 else ''} ready. Action items synced to CRM."
            if opp_count:
                msg += f" **{opp_count} opportunit{'y' if opp_count == 1 else 'ies'}** added to Opportunity Pipeline."
            st.success(msg)


def _render_action_table(df: pd.DataFrame):
    has_sector = "Sector" in df.columns and df["Sector"].fillna("").str.strip().any()
    rows_html  = ""
    for i, (_, row) in enumerate(df.iterrows()):
        en     = (row.get("Action (EN)") or row.get("Action (AR)", "")).strip()
        owner  = str(row.get("Assigned To", "") or "")
        sector = str(row.get("Sector", "") or "") if has_sector else ""
        prio   = str(row.get("Priority", "Medium") or "Medium")
        due    = str(row.get("Due Text EN", "") or row.get("Due Text", "") or
                     row.get("Due Date", "") or "TBD")
        status = str(row.get("Status", "Not Started") or "Not Started")
        rmk    = str(row.get("Remarks", "") or "")

        p_css  = _BADGE_CSS.get(prio,   "background:#f3f4f6;color:#6b7280")
        s_css  = _BADGE_CSS.get(status, "background:#f3f4f6;color:#6b7280")
        sec_td = f"<td style='font-size:11px;color:#1B5C3F'>{sector}</td>" if has_sector else ""

        rows_html += (
            f"<tr>"
            f"<td style='font-weight:500;text-align:center;width:36px'>{i+1}</td>"
            f"<td>{en}</td>{sec_td}<td>{owner}</td>"
            f"<td><span class='rb-pk' style='{p_css}'>{prio}</span></td>"
            f"<td>{due}</td>"
            f"<td><span class='rb-pk' style='{s_css}'>{status}</span></td>"
            f"<td style='font-size:11px;color:#6b7280'>{rmk}</td>"
            f"</tr>"
        )

    sector_hdr = "<th>Sector</th>" if has_sector else ""
    st.markdown(f"""
    <div style="overflow-x:auto">
    <table class="rb-action-table">
      <thead><tr><th>#</th><th>Action Item</th>{sector_hdr}<th>Owner</th>
      <th>Priority</th><th>Timeline</th><th>Status</th><th>Remarks</th></tr></thead>
      <tbody>{rows_html}</tbody>
    </table></div>
    """, unsafe_allow_html=True)


# ─── Multi-format meeting minutes parser ─────────────────────────────────────

def _parse_minutes_via_claude(file_bytes: bytes, media_type: str, api_key: str) -> dict:
    """Use Claude to extract meeting-minutes structure from any file format (PDF, image, etc.)."""
    try:
        import anthropic, json, base64
    except ImportError:
        return {}

    prompt = """You are a bilingual Arabic-English assistant for the Ministry of Investment of Saudi Arabia (MISA).
Extract meeting-minutes content from the attached document and return ONLY a JSON object with this structure:

{
  "company": "name of the visiting company",
  "date": "DD-MM-YYYY or empty string",
  "subject_ar": "Arabic meeting subject/title",
  "subject_en": "English translation of the subject",
  "location": "meeting location (Arabic)",
  "chair": "name of the chair (Arabic)",
  "priority": "Very High | High | Medium | Low",
  "next_meeting_text": "any text about next meeting date/topic",
  "attendees": "comma-separated list of attendee names and titles",
  "discussion_ar": "full discussion notes in Arabic (one paragraph)",
  "action_items": [
    {
      "Action (AR)": "action description in Arabic",
      "Action (EN)": "English translation",
      "Assigned To": "owner name or department",
      "Type": "Support | Opportunity | Challenge | Follow-up | Action | Administrative",
      "Priority": "Very High | High | Medium | Low",
      "Due Date": "YYYY-MM-DD or empty",
      "Due Text": "human readable due date or timeframe",
      "Remarks": ""
    }
  ]
}

Rules:
- Extract ALL action items mentioned anywhere in the document
- Keep Arabic text in Arabic script; provide English translations where shown
- Dates: output as DD-MM-YYYY for "date" field
- Respond ONLY with the JSON object — no markdown fences, no explanation
- If a field cannot be found, use an empty string or empty array"""

    try:
        client = anthropic.Anthropic(api_key=api_key)
        b64 = base64.standard_b64encode(file_bytes).decode()
        msg = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=8192,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "document" if media_type == "application/pdf" else "image",
                     "source": {"type": "base64", "media_type": media_type, "data": b64}}
                    if media_type != "application/pdf"
                    else
                    {"type": "document",
                     "source": {"type": "base64", "media_type": "application/pdf", "data": b64}},
                    {"type": "text", "text": prompt},
                ],
            }],
        )
        raw = msg.content[0].text.strip()
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
            raw = raw.rstrip("`").strip()
        data = json.loads(raw)
        # Normalise date field if returned
        from datetime import datetime as _dt
        for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                data["date"] = _dt.strptime(data.get("date", ""), fmt).date()
                break
            except (ValueError, TypeError):
                pass
        if not isinstance(data.get("date"), type(date.today())):
            data["date"] = None
        if not isinstance(data.get("action_items"), list):
            data["action_items"] = []
        return data
    except Exception:
        return {}


def _parse_brief_via_claude(file_bytes: bytes, media_type: str, api_key: str) -> dict:
    """Use Claude to extract company brief data from any file format (PDF, image, etc.)."""
    try:
        import anthropic, json, base64
    except ImportError:
        return {"error": "anthropic package not installed"}

    prompt = """You are an assistant for the Ministry of Investment of Saudi Arabia (MISA).
Extract all data from the attached company brief document and return ONLY a JSON object:

{
  "company": "company name",
  "company_info": {
    "sector": "primary sector/industry",
    "hq": "headquarters location",
    "aum": "assets under management or revenue figure",
    "ksa_presence": "Yes/No or description",
    "employees": "employee count",
    "website": "website URL",
    "rep_name": "company representative name",
    "email": "representative email",
    "phone": "representative phone"
  },
  "sectors": ["Sector A", "Sector B"],
  "action_items": [
    {
      "Action (EN)": "description of the action item",
      "Sector": "relevant sector name",
      "Assigned To": "owner name or department",
      "Type": "Support | Opportunity | Challenge | Follow-up | Action | Administrative",
      "Priority": "Very High | High | Medium | Low",
      "Start Date": "YYYY-MM-DD or empty",
      "Due Date": "YYYY-MM-DD or empty",
      "Progress": 0,
      "Status": "Not Started | In Progress | Completed | Blocked",
      "Remarks": ""
    }
  ]
}

Rules:
- Extract ALL action items mentioned anywhere in the document
- sectors: unique list of sector names that have action items
- Respond ONLY with the JSON object — no markdown fences, no explanation
- If a field cannot be found use empty string, empty array, or 0"""

    try:
        client = anthropic.Anthropic(api_key=api_key)
        b64 = base64.standard_b64encode(file_bytes).decode()
        content_block = (
            {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": b64}}
            if media_type == "application/pdf"
            else {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}}
        )
        msg = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=8192,
            messages=[{"role": "user", "content": [content_block, {"type": "text", "text": prompt}]}],
        )
        raw = msg.content[0].text.strip()
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
            raw = raw.rstrip("`").strip()
        data = json.loads(raw)
        if not isinstance(data.get("action_items"), list):
            data["action_items"] = []
        if not isinstance(data.get("sectors"), list):
            data["sectors"] = []
        if not isinstance(data.get("company_info"), dict):
            data["company_info"] = {}
        return data
    except Exception as exc:
        return {"error": str(exc)}


def _parse_tracker_via_claude(file_bytes: bytes, media_type: str, api_key: str) -> dict:
    """Use Claude to extract action items from a tracker document (PDF, image, etc.)."""
    try:
        import anthropic, json, base64
    except ImportError:
        return {"error": "anthropic package not installed"}

    prompt = """You are an assistant for the Ministry of Investment of Saudi Arabia (MISA).
Extract all action items from the attached document and return ONLY a JSON object:

{
  "company": "company name if visible",
  "action_items": [
    {
      "Action (AR)": "Arabic text if present, else empty string",
      "Action (EN)": "English description of the action",
      "Assigned To": "owner name or department",
      "Type": "Support | Opportunity | Challenge | Follow-up | Action | Administrative",
      "Priority": "Very High | High | Medium | Low",
      "Due Date": "YYYY-MM-DD or empty",
      "Due Text": "human readable due date or timeframe",
      "Status": "Not Started | In Progress | Completed | Blocked",
      "Remarks": ""
    }
  ]
}

Rules:
- Extract ALL action items, tasks, follow-ups, and commitments
- Respond ONLY with the JSON object — no markdown fences, no explanation
- If a field is not found use empty string"""

    try:
        client = anthropic.Anthropic(api_key=api_key)
        b64 = base64.standard_b64encode(file_bytes).decode()
        content_block = (
            {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": b64}}
            if media_type == "application/pdf"
            else {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}}
        )
        msg = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=8192,
            messages=[{"role": "user", "content": [content_block, {"type": "text", "text": prompt}]}],
        )
        raw = msg.content[0].text.strip()
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
            raw = raw.rstrip("`").strip()
        data = json.loads(raw)
        if not isinstance(data.get("action_items"), list):
            data["action_items"] = []
        return data
    except Exception as exc:
        return {"error": str(exc)}


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
        header_lower = header_text.lower()
        n_cols = max(len(r) for r in rows) if rows else 0

        # English attendee markers — check BEFORE Arabic markers to avoid false
        # positives: Arabic role titles (مسؤول…) appear in attendee cells and
        # would otherwise trigger the Arabic action-item branch.
        if len(rows) <= 6 and n_cols <= 3 and "name" in header_lower and ("role" in header_lower or "title" in header_lower):
            return "attendees"
        # Arabic attendee markers
        if any(m in header_text for m in ["الاسم", "المسمى", "الجهة", "حضر", "المشاركون", "التوقيع"]):
            return "attendees"

        # English action item markers (header row has these column names)
        if any(m in header_lower for m in ["action item", "due date", "success measure", "deliverable"]):
            return "actions"
        # Arabic action item markers — exclude مسؤول: it's a role title, not a
        # column header, and causes attendee tables to be misclassified.
        if any(m in header_text for m in ["التوجيه", "المهمة", "الأولوية", "الموعد النهائي", "الإجراء"]):
            return "actions"

        if len(rows) <= 5:
            for row in rows:
                for cell in row:
                    if len(cell) > 120:
                        return "discussion"
        if any(m in header_text for m in ["أبرز ما تم مناقشته", "نقاط النقاش", "مناقشة"]):
            return "discussion"
        return "metadata"

    # Classify all tables; when two tables share a kind, keep the one with more
    # columns (the real action table beats a misclassified 2-col attendee table).
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
    # Each row uses [label, value, label, value] pairs (not separate header/value rows)
    if tables:
        meta = tables[0]
        prio_map = {"مهم جدا": "Very High", "مهم": "High", "متوسط": "Medium", "عادي": "Low"}

        # Build label→value dict by reading alternating label/value cells per row
        _meta_kv: dict = {}
        for row in meta:
            for idx in range(0, len(row) - 1, 2):
                lbl = (row[idx] or "").strip()
                val = (row[idx + 1] if idx + 1 < len(row) else "") or ""
                val = val.strip()
                if lbl:
                    _meta_kv[lbl] = val

        if "الموضوع" in _meta_kv:
            result["subject_ar"] = _meta_kv["الموضوع"]
        if "الموقع" in _meta_kv and _meta_kv["الموقع"]:
            result["location"] = _meta_kv["الموقع"]
        if "برئاسة" in _meta_kv and _meta_kv["برئاسة"]:
            result["chair"] = _meta_kv["برئاسة"]
        if "الأولوية" in _meta_kv:
            result["priority"] = prio_map.get(_meta_kv["الأولوية"], "High")
        if "الاجتماع القادم" in _meta_kv:
            result["next_meeting_text"] = _meta_kv["الاجتماع القادم"]

        # Date: prefer "التاريخ", fallback "اليوم"
        for _dk in ["التاريخ", "اليوم"]:
            if _meta_kv.get(_dk):
                d = _parse_date(_meta_kv[_dk])
                if d:
                    result["date"] = d
                    break

        # Extract company name — first try splitting subject on em-dash "—"
        _subject = result["subject_ar"]
        for _sep in ("—", "–", "-"):
            _idx = _subject.find(_sep)
            if _idx != -1:
                _co_candidate = _subject[_idx + len(_sep):].strip()
                if _co_candidate and len(_co_candidate) >= 2:
                    result["company"] = _co_candidate
                    result["subject_en"] = f"Latest Updates — {_co_candidate}"
                    break
        # Fallback: scan for known company names anywhere in subject
        if not result["company"]:
            known = ["Rothschild", "Barclays", "BlackRock", "Brookfield", "Goldman",
                     "HSBC", "JPMorgan", "Morgan Stanley", "UBS", "Citi", "Deutsche",
                     "Allianz", "Lazard", "Blackstone", "Carlyle", "KKR",
                     "Vanguard", "Fidelity", "Temasek", "Mubadala", "ADQ", "PIF"]
            for name in known:
                if name.lower() in _subject.lower():
                    result["company"] = name
                    result["subject_en"] = f"Latest Updates — {name}"
                    break

    # ── Fallback: extract company, date, subject from paragraphs (English MoMs) ─
    if not result["company"] or not result["date"]:
        for para in doc.paragraphs:
            txt = para.text.strip()
            if not txt:
                continue
            # Subject line: "Subject: Latest updates — bnp paribas"
            if not result["company"] and re.match(r"(?i)subject\s*[:\-–]", txt):
                for _sep in ("—", "–", "-"):
                    if _sep in txt:
                        _co = txt.split(_sep, 1)[-1].strip()
                        if _co and len(_co) >= 2:
                            result["company"] = _co
                            result["subject_en"] = f"Latest Updates — {_co}"
                            break
            # Date anywhere in paragraph: DD/MM/YYYY or YYYY-MM-DD
            if not result["date"]:
                _dm = re.search(r"\b(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{4})\b", txt)
                if _dm:
                    try:
                        result["date"] = date(int(_dm.group(3)), int(_dm.group(2)), int(_dm.group(1)))
                    except Exception:
                        pass

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
        # Build column-index map from header row so insertion order doesn't matter
        _hdr = [c.strip().lower() for c in act_tbl[0]] if act_tbl else []
        def _col(names):
            for n in names:
                for i, h in enumerate(_hdr):
                    if n in h:
                        return i
            return -1
        _ci_action = _col(["action item", "التوجيه", "المهمة", "الإجراء"])
        _ci_owner  = _col(["owner", "assigned", "مسؤول", "المسؤول"])
        _ci_prio   = _col(["priority", "الأولوية"])
        _ci_due    = _col(["due date", "due", "الموعد", "timeline", "target"])

        for row in act_tbl[1:]:
            cells = _dedup(row)
            if not any(c for c in cells if c and not re.match(r"^\d+$", c)):
                continue

            def _get(idx):
                return cells[idx].strip() if 0 <= idx < len(cells) else ""

            # Fall back to positional (non-empty, non-numeric) if header not found
            non_empty = [c for c in cells if c and not re.match(r"^\d+$", c)]
            action_ar = _get(_ci_action) if _ci_action >= 0 else (non_empty[0] if non_empty else "")
            if not action_ar or action_ar in ("م", "التوجيه / المهمة", "Action Item", "#"):
                continue

            owner_ar = _get(_ci_owner) if _ci_owner >= 0 else (non_empty[1] if len(non_empty) > 1 else "")
            prio_ar  = _get(_ci_prio)  if _ci_prio  >= 0 else (non_empty[2] if len(non_empty) > 2 else "")
            due_ar   = _get(_ci_due)   if _ci_due   >= 0 else (non_empty[3] if len(non_empty) > 3 else "")

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


# ─── PPT company brief parser ────────────────────────────────────────────────

def _parse_company_brief_pptx(file_bytes: bytes) -> dict:
    """Parse GA Company Brief PPT format.

    Slide 1 = cover (company name), Slide 2 = fact sheet,
    Slides 3+ = per-sector action tables (each has Sector POV label + action table).
    Returns {company, company_info, action_items, sectors} or {error: msg}.
    """
    try:
        from pptx import Presentation  # type: ignore
    except ImportError:
        return {"error": "python-pptx not installed — run: pip install python-pptx"}
    try:
        prs = Presentation(io.BytesIO(file_bytes))
    except Exception as exc:
        return {"error": str(exc)}

    _STATUS_MAP = {
        "completed":   "Completed",
        "complete":    "Completed",
        "in progress": "In Progress",
        "inprogress":  "In Progress",
        "not started": "Not Started",
        "blocked":     "Blocked",
        "cancelled":   "Cancelled",
    }

    result: dict = {"company": "", "company_info": {}, "action_items": [], "sectors": []}

    def _shape_texts(slide):
        for shp in slide.shapes:
            if shp.has_table:
                continue
            if shp.has_text_frame:
                txt = shp.text_frame.text.strip()
                if txt:
                    yield txt

    # ── Slide 1 — company name ─────────────────────────────────────────────────
    if prs.slides:
        _skip_kw = {"misa", "ministry", "kingdom", "vision 2030", "investor"}
        for txt in _shape_texts(list(prs.slides)[0]):
            if 2 < len(txt) < 80:
                low = txt.lower()
                if not any(k in low for k in _skip_kw):
                    result["company"] = txt
                    break

    # ── Slide 2 — company fact sheet ──────────────────────────────────────────
    if len(prs.slides) > 1:
        info: dict = {}
        for txt in _shape_texts(list(prs.slides)[1]):
            for line in txt.splitlines():
                line = line.strip()
                if ":" not in line:
                    continue
                lbl, _, val = line.partition(":")
                lbl = lbl.strip().lower()
                val = val.strip()
                if not val:
                    continue
                if "background" in lbl:
                    info["background"] = val
                elif lbl == "sector" or ("sector" in lbl and "pov" not in lbl):
                    info["sector"] = val
                elif "hq" in lbl or "headquarter" in lbl:
                    info["hq"] = val
                elif "website" in lbl or lbl == "web":
                    info["website"] = val
                elif "aum" in lbl:
                    info["aum"] = val
                elif "ksa" in lbl or "presence" in lbl:
                    info["ksa_presence"] = val
                elif "employee" in lbl:
                    info["employees"] = val
                elif "rep" in lbl:
                    info["rep_name"] = val
                elif "email" in lbl:
                    info["email"] = val
                elif "phone" in lbl or "mobile" in lbl or lbl == "tel":
                    info["phone"] = val
                elif "position" in lbl or "title" in lbl or "role" in lbl:
                    info["position"] = val
                elif "vision" in lbl:
                    info["vision_alignment"] = val
        result["company_info"] = info

    # ── Slides 3+ — sector action slides ──────────────────────────────────────
    for slide in list(prs.slides)[2:]:
        # Find action table
        action_tbl = None
        for shp in slide.shapes:
            if not shp.has_table:
                continue
            tbl = shp.table
            if len(tbl.rows) < 2 or len(tbl.columns) < 3:
                continue
            hdr_texts = [tbl.cell(0, c).text.strip().lower()
                         for c in range(len(tbl.columns))]
            if any("action" in h for h in hdr_texts):
                action_tbl = tbl
                break
        if action_tbl is None:
            continue

        # Extract metadata from text shapes
        sector_name   = ""
        sector_pov    = ""
        next_meeting  = ""
        latest_update = ""
        candidates: list[str] = []

        for txt in _shape_texts(slide):
            low = txt.lower()
            if "sector pov" in low or low.startswith("pov"):
                sector_pov = txt.split(":", 1)[-1].strip() if ":" in txt else sector_pov
            elif "next meeting" in low:
                next_meeting = txt.split(":", 1)[-1].strip() if ":" in txt else next_meeting
            elif "latest update" in low or "last update" in low:
                latest_update = txt.split(":", 1)[-1].strip() if ":" in txt else latest_update
            elif ":" not in txt and 3 < len(txt) < 60:
                candidates.append(txt)

        # Best candidate for sector name
        for cand in candidates:
            low = cand.lower()
            if low.endswith(" sector"):
                sector_name = cand[:-7].strip()
                break
            if not re.match(r"^\d", cand) and low not in ("action items", "status", "remarks"):
                sector_name = cand
                break

        # Column map
        col_map: dict[str, int] = {}
        for c in range(len(action_tbl.columns)):
            col_map[action_tbl.cell(0, c).text.strip().lower()] = c

        def _find_col(*terms):
            for term in terms:
                for k, v in col_map.items():
                    if term in k:
                        return v
            return None

        action_col = _find_col("action item", "action")
        type_col   = _find_col("type of engagement", "type")
        start_col  = _find_col("start date")
        due_col    = _find_col("due date", "due")
        prio_col   = _find_col("priority")
        prog_col   = _find_col("progress")
        status_col = _find_col("status")
        remark_col = _find_col("remark", "note")

        if action_col is None:
            continue

        def _cell_txt(r, c):
            if c is None:
                return ""
            try:
                return action_tbl.cell(r, c).text.strip()
            except Exception:
                return ""

        for r in range(1, len(action_tbl.rows)):
            action = _cell_txt(r, action_col)
            if not action:
                continue
            st_raw  = _cell_txt(r, status_col)
            st_norm = _STATUS_MAP.get(st_raw.lower(), st_raw or "Not Started")
            try:
                prog_val = int(_cell_txt(r, prog_col).replace("%", "").strip() or "0")
            except Exception:
                prog_val = 0
            result["action_items"].append({
                "Action (AR)":    "",
                "Action (EN)":    action,
                "Sector":         sector_name,
                "Assigned To":    sector_pov,
                "Type":           _cell_txt(r, type_col) or "Action",
                "Priority":       _cell_txt(r, prio_col) or "Medium",
                "Start Date":     _cell_txt(r, start_col),
                "Due Date":       _cell_txt(r, due_col),
                "Progress":       prog_val,
                "Status":         st_norm,
                "Remarks":        _cell_txt(r, remark_col),
                "_latest_update": latest_update,
            })

        if sector_name and sector_name not in result["sectors"]:
            result["sectors"].append(sector_name)

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
    arm        = cfg.get("arm") or ""
    exec_rm    = cfg.get("exec_rm") or ""
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
    _run(ref_p, f"     |     Ref: MISA / {co} / AM / {yr_mon}",
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
    if arm:
        _run(p4, "In this regard, we are pleased to confirm that ")
        _run(p4, arm, bold=True)
        _run(p4, " will be serving as the ")
        _run(p4, f"Account Manager (AM)", bold=True)
        _run(p4, f" for {co} — your primary point of contact for all day-to-day "
             "operational matters and coordination. ")
    if exec_rm:
        _run(p4, exec_rm, bold=True)
        _run(p4, " from the Minister's Office will serve as the ")
        _run(p4, "Relationship Manager (RM)", bold=True)
        _run(p4, " for any topics related to the Minister.")
    if not arm and not exec_rm:
        _run(p4, "Our team remains fully committed to supporting your continued growth "
             "and success in the Kingdom.")
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
    _run(sig_t, "Relationship Manager (RM)  |  Minister's Office",
         size_pt=10.5, color="555555")

    sig_m = doc.add_paragraph()
    _run(sig_m, "Ministry of Investment  |  Kingdom of Saudi Arabia",
         size_pt=10.5, color="555555")

    sig_e = doc.add_paragraph()
    if exec_rm:
        _run(sig_e, "E: ",          size_pt=10, color="888888")
        _run(sig_e, exec_email,     size_pt=10, color="1A5276")
    if arm:
        _run(sig_e, "   |   AM: ",  size_pt=10, color="888888")
        _run(sig_e, arm_email,      size_pt=10, color="1A5276")

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

def _find_company_sheet(wb, company: str) -> str | None:
    """Return the sheet name that best matches `company`, or None."""
    co = company.lower().strip()
    for sname in wb.sheetnames:
        sn = sname.lower()
        if co in sn or sn in co:
            return sname
        # Strip "Action Items" prefix and try again
        stripped = sn
        for pfx in ("action items ", "actions ", "action item "):
            if stripped.startswith(pfx):
                stripped = stripped[len(pfx):].strip()
                break
        if stripped and (stripped in co or co in stripped):
            return sname
        # Word-by-word: any 4+ char token from stripped appears in company
        for word in stripped.split():
            if len(word) >= 4 and word in co:
                return sname
    return None


def _build_excel(existing_bytes, company, meeting_date, next_meeting, chair, actions_df) -> bytes:
    if existing_bytes:
        wb = openpyxl.load_workbook(io.BytesIO(existing_bytes))
    else:
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    sheet_name = _find_company_sheet(wb, company) or f"Action Items {company}"[:31]

    if sheet_name in wb.sheetnames:
        ws       = wb[sheet_name]
        # Scan backwards from capped max_row to find last data row quickly
        last_row = 20
        scan_max = min(ws.max_row or 21, 5000)
        for r in range(scan_max, 20, -1):
            if any(ws.cell(r, c).value is not None for c in range(7, 17)):
                last_row = r
                break
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


# ─── PPT→Excel merge ─────────────────────────────────────────────────────────

def _merge_pptx_actions_to_excel(
    existing_bytes: bytes, company: str, pptx_actions: list
) -> tuple:
    """
    Upsert PPT action items into the company's Excel sheet.
    Existing rows matched by Action Item text are updated (Status, Remarks,
    Progress, Due Date). Unmatched rows are appended at the end.
    Returns (updated_bytes, n_updated, n_added).
    """
    wb = openpyxl.load_workbook(io.BytesIO(existing_bytes))

    # Locate or create sheet
    target = _find_company_sheet(wb, company)
    if not target:
        ws = wb.create_sheet(f"Action Items {company}"[:31])
        _write_sheet_header(ws, company, date.today(), None, "")
        target = ws.title

    ws = wb[target]

    # Find header row — scan rows 15-25 for "Action Item" cell
    hdr_row = 20
    found   = False
    for r in range(15, 26):
        for c in range(1, 20):
            if str(ws.cell(row=r, column=c).value or "").strip() == "Action Item":
                hdr_row = r
                found   = True
                break
        if found:
            break

    # Build column index from header row
    col_map: dict[str, int] = {}
    for c in range(1, 20):
        v = str(ws.cell(row=hdr_row, column=c).value or "").strip()
        if v:
            col_map[v] = c

    id_col     = col_map.get("ID", 7)
    act_col    = col_map.get("Action Item", 8)
    asgn_col   = next((v for k, v in col_map.items() if "Assigned" in k), 9)
    type_col   = next((v for k, v in col_map.items() if "Type" in k and "Engagement" in k), 10)
    start_col  = col_map.get("Start Date", 11)
    due_col    = col_map.get("Due Date", 12)
    prio_col   = col_map.get("Priority", 13)
    prog_col   = col_map.get("Progress", 14)
    status_col = col_map.get("Status", 15)
    rmk_col    = col_map.get("Remarks", 16)

    # Build lookup: normalised action text → row number
    existing: dict[str, int] = {}
    for r in range(hdr_row + 1, hdr_row + 500):
        val = ws.cell(row=r, column=act_col).value
        if val is None:
            break
        norm = str(val).strip().lower()
        if norm:
            existing[norm] = r

    normal_font = Font(size=10)
    center_al   = Alignment(horizontal="center", vertical="center")
    right_al    = Alignment(horizontal="right",  vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"),  right=Side(style="thin"),
        top=Side(style="thin"),   bottom=Side(style="thin"),
    )

    n_updated = 0
    n_added   = 0
    next_row  = hdr_row + 1 + len(existing)

    for act in pptx_actions:
        text = str(act.get("Action (EN)") or act.get("Action (AR)") or "").strip()
        if not text:
            continue
        norm = text.lower()

        if norm in existing:
            r = existing[norm]
            new_status = str(act.get("Status") or "").strip()
            if new_status:
                ws.cell(r, status_col, new_status)
            new_rem = str(act.get("Remarks") or "").strip()
            if new_rem and new_rem != str(ws.cell(r, rmk_col).value or "").strip():
                ws.cell(r, rmk_col, new_rem)
            if act.get("Progress") is not None:
                ws.cell(r, prog_col, int(act["Progress"]))
            if act.get("Due Date"):
                ws.cell(r, due_col, str(act["Due Date"]))
            n_updated += 1
        else:
            r        = next_row
            next_row += 1
            seq      = len(existing) + n_added + 1
            ws.cell(r, id_col,     seq)
            ws.cell(r, act_col,    text)
            ws.cell(r, asgn_col,   str(act.get("Assigned To") or ""))
            ws.cell(r, type_col,   str(act.get("Type") or "Action"))
            ws.cell(r, start_col,  str(act.get("Start Date") or ""))
            ws.cell(r, due_col,    str(act.get("Due Date") or ""))
            ws.cell(r, prio_col,   str(act.get("Priority") or "Medium"))
            ws.cell(r, prog_col,   int(act.get("Progress") or 0))
            ws.cell(r, status_col, str(act.get("Status") or "Not Started"))
            ws.cell(r, rmk_col,    str(act.get("Remarks") or ""))
            for col in range(7, 17):
                cell           = ws.cell(r, col)
                cell.font      = normal_font
                cell.border    = thin_border
                cell.alignment = right_al if col == act_col else center_al
            ws.row_dimensions[r].height = 40
            existing[norm] = r
            n_added += 1

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), n_updated, n_added


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
            "Last Updated":       np.datetime64(datetime.now()),
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
            "Sector":             str(row.get("Sector", "") or ""),
            "Type of Engagement": str(row.get("Type", "Action") or "Action"),
            "Start Date":         date.today(),
            "Due Date":           row.get("Due Date") if row.get("Due Date") and pd.notna(row.get("Due Date")) else None,
            "Priority":           str(row.get("Priority", "Medium") or "Medium"),
            "Progress":           "0%",
            "Status":             "Not Started",
            "Escalation Flag":    "None",
            "Remarks":            str(row.get("Remarks", "") or ""),
            "Last Updated":       np.datetime64(datetime.now()),
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
                "company_name": "",
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
                    elif "COMPANY" in lu and "NAME" in lu:
                        info["company_name"] = val
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
            # Also index by the "company name" cell value so the dropdown
            # (which may show the full name e.g. "Lulu Group") resolves correctly
            if info["company_name"] and info["company_name"] != co:
                result[info["company_name"]] = info
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
    if info.get("last_updated"):
        st.session_state["rb_date"]      = info["last_updated"]
        st.session_state["rb_date_inp"]  = info["last_updated"]
        st.session_state["rb_ar_date"]   = info["last_updated"]
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


def _load_excel_actions_for_company(raw_bytes: bytes, company: str) -> pd.DataFrame:
    """
    Extract action item rows from the Excel tracker sheet matching `company`.
    The sheet title is "Action Items <company>" and the header row is auto-detected
    by scanning for the cell 'Action Item' anywhere in rows 15-25.
    Returns a DataFrame in rb_actions format, or empty DataFrame on failure.
    """
    _PRIORITY_MAP = {"Very High": "Very High", "High": "High", "Medium": "Medium", "Low": "Low",
                     "مهم جدا": "Very High", "مهم": "High", "متوسط": "Medium", "عادي": "Low"}
    try:
        wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
        # Find the sheet whose name matches the company
        target_ws = None
        co_lower = company.strip().lower()
        for ws in wb.worksheets:
            sheet_co = ws.title.strip()
            for prefix in ("Action Items", "Actions", "Action Item"):
                if sheet_co.lower().startswith(prefix.lower()):
                    sheet_co = sheet_co[len(prefix):].strip()
                    break
            if co_lower in sheet_co.lower() or sheet_co.lower() in co_lower:
                target_ws = ws
                break

        if target_ws is None:
            return pd.DataFrame()

        ws = target_ws
        # Find header row
        hdr_row = None
        col_map: dict[str, int] = {}
        for r in range(15, 26):
            for c in range(1, 20):
                v = str(ws.cell(row=r, column=c).value or "").strip()
                if v.lower() in ("action item", "action", "task", "بند العمل"):
                    hdr_row = r
                    break
            if hdr_row:
                break
        if not hdr_row:
            return pd.DataFrame()

        # Map column headers
        for c in range(1, 20):
            v = str(ws.cell(row=hdr_row, column=c).value or "").strip()
            if v:
                col_map[v.lower()] = c

        def _col(*aliases):
            for a in aliases:
                for k, v in col_map.items():
                    if a.lower() in k or k in a.lower():
                        return v
            return None

        id_c    = _col("id", "#", "no", "م")
        item_c  = _col("action item", "action", "task", "بند")
        owner_c = _col("assigned to", "owner", "responsible", "rep", "المسؤول")
        type_c  = _col("type of engagement", "type", "engagement", "النوع")
        start_c = _col("start date", "start", "بداية")
        due_c   = _col("due date", "due", "deadline", "الاستحقاق")
        prio_c  = _col("priority", "الأولوية")
        prog_c  = _col("progress", "التقدم")
        rmk_c   = _col("remarks", "am input", "notes", "ملاحظات")

        if not item_c:
            return pd.DataFrame()

        rows = []
        for r in range(hdr_row + 1, hdr_row + 200):
            id_val   = str(ws.cell(row=r, column=id_c).value or "").strip() if id_c else ""
            item_val = str(ws.cell(row=r, column=item_c).value or "").strip()
            if not item_val:
                break  # end of data
            prio_raw = str(ws.cell(row=r, column=prio_c).value or "").strip() if prio_c else ""
            prio_en  = _PRIORITY_MAP.get(prio_raw, "Medium")
            due_raw  = ws.cell(row=r, column=due_c).value if due_c else None
            if hasattr(due_raw, "strftime"):
                due_str = due_raw.strftime("%d %B %Y")
            else:
                due_str = str(due_raw or "").strip()
            rows.append({
                "Action (AR)": item_val,
                "Action (EN)": item_val,
                "Assigned To": str(ws.cell(row=r, column=owner_c).value or "").strip() if owner_c else "",
                "Type":        str(ws.cell(row=r, column=type_c).value or "Action").strip() if type_c else "Action",
                "Priority":    prio_en,
                "Due Date":    None,
                "Due Text":    due_str,
                "Due Text EN": due_str,
                "Remarks":     str(ws.cell(row=r, column=rmk_c).value or "").strip() if rmk_c else "",
                "Status":      "Not Started",
            })

        return pd.DataFrame(rows) if rows else pd.DataFrame()

    except Exception:
        return pd.DataFrame()


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


# ─── Arabic minutes — Claude extraction ──────────────────────────────────────

def _extract_via_claude(summary_text: str, cfg: dict, api_key: str) -> dict:
    """Call Claude to extract and translate meeting content into Arabic JSON."""
    try:
        import anthropic
    except ImportError:
        return {"error": "anthropic package not installed — run: pip install anthropic"}

    company = cfg.get("company", "")
    date_str = cfg.get("date", "")
    arm      = cfg.get("arm", "")
    exec_rm  = cfg.get("exec_rm", "")

    prompt = f"""You are a bilingual English-Arabic assistant for the Ministry of Investment of Saudi Arabia (MISA).

Analyse the English meeting summary below, then return a single JSON object with all values in Arabic.

Required JSON structure:
{{
  "subject_ar": "short Arabic subject line, e.g. آخر المستجدات — {company}",
  "discussion_points": ["Arabic bullet 1", "Arabic bullet 2", ...],
  "action_items": [
    {{
      "task":     "Arabic description of the action",
      "task_en":  "English description of the same action",
      "owner":    "person or department name in Arabic",
      "priority": "مهم جدا | مهم | متوسط | عادي",
      "due":      "expected completion date or Arabic timeframe"
    }}
  ],
  "attendees": [
    {{
      "name":  "Full name",
      "title": "Job title in Arabic"
    }}
  ]
}}

Rules:
- discussion_points: 4–8 key points, each a full Arabic sentence
- action_items: one entry per distinct task mentioned
- attendees: include all named people from both sides (MISA and {company})
- If AM (Account Manager) is "{arm}" or RM (Relationship Manager) is "{exec_rm}", include them in attendees as MISA staff
- task_en: English translation of the action (used for the Excel tracker)
- All other text values MUST be in Arabic (names may stay in original script)
- Respond ONLY with the JSON object — no markdown fences, no explanation

Meeting metadata:
  Company : {company}
  Date    : {date_str}

Meeting summary:
{summary_text[:5000]}"""

    try:
        client = anthropic.Anthropic(api_key=api_key)
        msg = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=8192,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = msg.content[0].text.strip()
        # Strip accidental markdown fences
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
            raw = raw.rstrip("`").strip()
        import json

        def _repair_json(s: str) -> str:
            """Escape literal newlines inside JSON strings; close truncated structures."""
            out = []
            in_str = False
            i = 0
            while i < len(s):
                c = s[i]
                if c == "\\" and in_str:          # escaped char — copy both bytes
                    out.append(c)
                    i += 1
                    if i < len(s):
                        out.append(s[i])
                    i += 1
                    continue
                if c == '"':
                    in_str = not in_str
                    out.append(c)
                    i += 1
                    continue
                if in_str and c in "\n\r":         # bare newline inside a string value
                    out.append("\\n")
                    if c == "\r" and i + 1 < len(s) and s[i + 1] == "\n":
                        i += 1                     # skip \r of \r\n pair
                    i += 1
                    continue
                out.append(c)
                i += 1

            fixed = "".join(out)
            if in_str:                             # string was never closed (truncated)
                fixed += '"'
            fixed = fixed.rstrip().rstrip(",")     # drop trailing comma
            fixed += "]" * max(0, fixed.count("[") - fixed.count("]"))
            fixed += "}" * max(0, fixed.count("{") - fixed.count("}"))
            return fixed

        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            return json.loads(_repair_json(raw))
    except Exception as exc:
        return {"error": str(exc)}


# ─── Arabic minutes — Word document builder ──────────────────────────────────

def _build_arabic_minutes_docx(cfg: dict, content: dict) -> bytes:
    """
    Generate the Ministry of Investment Arabic meeting minutes Word document.

    4-table structure:
      Table 0 — meeting metadata (subject, location, day/date, chair, priority, next meeting)
      Table 1 — key discussion points (bullet list in one cell)
      Table 2 — action items (م | التوجيه/المهمة | المسؤول | الأولوية | تاريخ الإنجاز)
      Table 3 — attendees (# | الاسم | الوظيفة) with dotted borders
    """
    from docx import Document as _Doc
    from docx.shared import Pt, RGBColor, Inches, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    AR_FONT   = "Sakkal Majalla"
    GREEN_HEX = "1B5C3F"
    GOLD_HEX  = "C9974A"

    doc = _Doc()
    sec = doc.sections[0]
    sec.page_width    = Cm(29.7)
    sec.page_height   = Cm(21.0)   # A4 landscape
    sec.left_margin   = sec.right_margin  = Cm(1.8)
    sec.top_margin    = sec.bottom_margin = Cm(1.5)

    def _rgb_from_hex(h: str) -> RGBColor:
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

    def _set_cell_bg(cell, hex_color: str):
        tcPr = cell._tc.get_or_add_tcPr()
        shd  = OxmlElement("w:shd")
        shd.set(qn("w:val"),   "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"),  hex_color.upper())
        tcPr.append(shd)

    def _set_rtl_para(para):
        pPr = para._p.get_or_add_pPr()
        bidi = OxmlElement("w:bidi")
        pPr.append(bidi)

    def _fill_cell(cell, text: str, bold=False, size=12,
                   bg_hex=None, color_hex="000000",
                   align=WD_ALIGN_PARAGRAPH.RIGHT):
        cell.text = ""
        para = cell.paragraphs[0]
        para.alignment = align
        _set_rtl_para(para)
        run = para.add_run(text)
        run.bold           = bold
        run.font.name      = AR_FONT
        run.font.size      = Pt(size)
        run.font.color.rgb = _rgb_from_hex(color_hex)
        if bg_hex:
            _set_cell_bg(cell, bg_hex)

    def _ar_heading(text: str, size=13, bold=True):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        _set_rtl_para(p)
        r = p.add_run(text)
        r.bold       = bold
        r.font.name  = AR_FONT
        r.font.size  = Pt(size)
        r.font.color.rgb = _rgb_from_hex(GREEN_HEX)
        return p

    # ── Page header ────────────────────────────────────────────────────────────
    hdr_p = doc.add_paragraph()
    hdr_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_rtl_para(hdr_p)
    hdr_r = hdr_p.add_run("وزارة الاستثمار | Ministry of Investment")
    hdr_r.bold       = True
    hdr_r.font.name  = AR_FONT
    hdr_r.font.size  = Pt(16)
    hdr_r.font.color.rgb = _rgb_from_hex(GREEN_HEX)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_rtl_para(title_p)
    title_r = title_p.add_run("محضر اجتماع")
    title_r.bold       = True
    title_r.font.name  = AR_FONT
    title_r.font.size  = Pt(20)
    title_r.font.color.rgb = _rgb_from_hex(GREEN_HEX)
    doc.add_paragraph()

    company  = cfg.get("company",  "")
    date_str = cfg.get("date",     "")
    location = cfg.get("location", "المقر الرئيسي – وزارة الاستثمار")
    chair    = cfg.get("chair",    "معالي الوزير")
    priority = cfg.get("priority", "مهم جدا")
    next_mtg = cfg.get("next_mtg", "")
    subject  = content.get("subject_ar", f"آخر المستجدات — {company}")

    # ── Table 0: Meeting metadata ──────────────────────────────────────────────
    _ar_heading("بيانات الاجتماع", size=12)
    tbl0 = doc.add_table(rows=4, cols=4)
    tbl0.style = "Table Grid"

    meta_rows = [
        [("الموضوع", True), (subject,   False), ("الموقع",         True), (location, False)],
        [("اليوم",   True), (date_str,  False), ("التاريخ",        True), (date_str, False)],
        [("الوقت",   True), ("",         False), ("برئاسة",         True), (chair,    False)],
        [("الأولوية",True), (priority,  False), ("الاجتماع القادم",True), (next_mtg, False)],
    ]
    for ri, row_data in enumerate(meta_rows):
        for ci, (text, is_label) in enumerate(row_data):
            c  = tbl0.cell(ri, ci)
            bg = GREEN_HEX if is_label else None
            fg = "FFFFFF"  if is_label else "1A1A1A"
            _fill_cell(c, text, bold=is_label, size=11, bg_hex=bg, color_hex=fg)
    doc.add_paragraph()

    # ── Table 1: Key discussion points ────────────────────────────────────────
    _ar_heading("أبرز ما تم مناقشته")
    tbl1 = doc.add_table(rows=1, cols=1)
    tbl1.style = "Table Grid"
    disc_lines = content.get("discussion_points", [])
    disc_text  = "\n".join(f"• {pt}" for pt in disc_lines) if disc_lines else "—"
    _fill_cell(tbl1.cell(0, 0), disc_text, size=12)
    doc.add_paragraph()

    # ── Table 2: Action items ─────────────────────────────────────────────────
    _ar_heading("الإجراءات والمهام المتفق عليها")
    action_items = content.get("action_items", [])
    tbl2 = doc.add_table(rows=1 + max(len(action_items), 1), cols=5)
    tbl2.style = "Table Grid"
    hdrs2 = ["م", "التوجيه / المهمة", "المسؤول", "الأولوية", "تاريخ الإنجاز المتوقع"]
    for ci, h in enumerate(hdrs2):
        _fill_cell(tbl2.cell(0, ci), h, bold=True, size=11,
                   bg_hex=GREEN_HEX, color_hex="FFFFFF")
    if action_items:
        for ri, act in enumerate(action_items):
            vals = [str(ri + 1), act.get("task", ""), act.get("owner", ""),
                    act.get("priority", ""), act.get("due", "")]
            for ci, val in enumerate(vals):
                _fill_cell(tbl2.cell(ri + 1, ci), val, size=11)
    else:
        for ci in range(5):
            _fill_cell(tbl2.cell(1, ci), "", size=11)
    doc.add_paragraph()

    # ── Table 3: Attendees ────────────────────────────────────────────────────
    _ar_heading("الحضور")
    attendees = content.get("attendees", [])
    tbl3 = doc.add_table(rows=1 + max(len(attendees), 1), cols=3)
    tbl3.style = "Table Grid"
    hdrs3 = ["#", "الاسم", "الوظيفة"]
    for ci, h in enumerate(hdrs3):
        _fill_cell(tbl3.cell(0, ci), h, bold=True, size=11,
                   bg_hex=GREEN_HEX, color_hex="FFFFFF")
    if attendees:
        for ri, att in enumerate(attendees):
            vals = [str(ri + 1), att.get("name", ""), att.get("title", "")]
            for ci, val in enumerate(vals):
                _fill_cell(tbl3.cell(ri + 1, ci), val, size=11)
    else:
        for ci in range(3):
            _fill_cell(tbl3.cell(1, ci), "", size=11)

    # ── Footer ────────────────────────────────────────────────────────────────
    doc.add_paragraph()
    ft = doc.add_paragraph()
    ft.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_rtl_para(ft)
    ft_r = ft.add_run("وزارة الاستثمار — المملكة العربية السعودية | www.misa.gov.sa")
    ft_r.font.name  = AR_FONT
    ft_r.font.size  = Pt(9)
    ft_r.font.color.rgb = _rgb_from_hex("888888")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _build_english_minutes_docx(cfg: dict, content: dict) -> bytes:
    """
    Generate English Meeting Minutes (MoM) matching the Minister Outreach Office template.
    Layout: header → subject → attendees table → meeting objective → discussion points
            → action items table → next steps → immediate priority → footer.
    """
    from docx import Document as _Doc
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    GREEN_HEX = "1B5C3F"
    GOLD_HEX  = "C9974A"

    doc = _Doc()
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
        r.bold           = bold
        r.font.size      = Pt(size_pt)
        r.font.color.rgb = _rgb(color)
        return r

    def _para(text="", bold=False, size_pt=11.0, color="1C1C1C",
              align=WD_ALIGN_PARAGRAPH.LEFT):
        p = doc.add_paragraph()
        p.alignment = align
        if text:
            _run(p, text, bold=bold, size_pt=size_pt, color=color)
        return p

    def _section_heading(text: str):
        p = doc.add_paragraph()
        r = p.add_run(text)
        r.bold           = True
        r.font.size      = Pt(12)
        r.font.color.rgb = _rgb(GREEN_HEX)
        return p

    company  = cfg.get("company", "")
    mtg_date = cfg.get("meeting_date", "") or cfg.get("date", "")
    arm      = cfg.get("arm", "") or ""
    exec_rm  = cfg.get("exec_rm", "") or ""
    yr_mon   = date.today().strftime("%Y-%m")

    subject_ar = content.get("subject_ar", "")
    subject_en = content.get("subject_en", "")
    if not subject_en and subject_ar:
        subject_en = _translate_to_en(subject_ar) or subject_ar

    discussion_ar = content.get("discussion_points", [])
    discussion_en = _translate_list(discussion_ar) if discussion_ar else []
    action_items  = content.get("action_items", [])
    attendees     = content.get("attendees",    [])

    # Prepared-by line
    prep = doc.add_paragraph()
    r1 = prep.add_run("Prepared by: ")
    r1.bold = True; r1.font.size = Pt(9); r1.font.color.rgb = _rgb("555555")
    r2 = prep.add_run("Minister Outreach Office, MISA")
    r2.font.size = Pt(9); r2.font.color.rgb = _rgb("555555")
    r3 = prep.add_run("    For internal use only – Ministry of Investment of Saudi Arabia")
    r3.italic = True; r3.font.size = Pt(9); r3.font.color.rgb = _rgb("888888")

    # Title block
    hdr_tbl = doc.add_table(rows=2, cols=1)
    hdr_tbl.style = "Table Grid"
    c0 = hdr_tbl.rows[0].cells[0]
    _cell_bg(c0, GREEN_HEX)
    p0 = c0.paragraphs[0]
    r_t = p0.add_run("Minutes of Meeting (MoM)")
    r_t.bold = True; r_t.font.size = Pt(14); r_t.font.color.rgb = _rgb("FFFFFF")
    c1 = hdr_tbl.rows[1].cells[0]
    _cell_bg(c1, "1A5C3F")
    p1 = c1.paragraphs[0]
    r_s = p1.add_run("Ministry of Investment of Saudi Arabia (MISA)")
    r_s.font.size = Pt(10); r_s.font.color.rgb = _rgb(GOLD_HEX)

    _para()

    ref_p = doc.add_paragraph()
    ref_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _run(ref_p, "CONFIDENTIAL", bold=True, size_pt=10, color=GOLD_HEX)
    _run(ref_p, f"    {mtg_date}", size_pt=10, color="555555")
    ref_str = f"MISA / {company} / {yr_mon}"
    _run(ref_p, f"    Ref: {ref_str}", size_pt=10, color="888888")

    _para()

    subj_p = doc.add_paragraph()
    _run(subj_p, "Subject: ", bold=True, size_pt=11.5)
    _run(subj_p, subject_en or f"Meeting with {company}", bold=True, size_pt=11.5)

    _para()

    # Attendees
    _section_heading("Attendees")
    att_rows = list(attendees) if attendees else []
    if arm and not any(a.get("name", "") == arm for a in att_rows):
        att_rows.append({"name": arm,     "title": "Account Manager (AM), MISA"})
    if exec_rm and not any(a.get("name", "") == exec_rm for a in att_rows):
        att_rows.append({"name": exec_rm, "title": "Relationship Manager (RM), Minister's Office"})

    att_tbl = doc.add_table(rows=1 + max(len(att_rows), 1), cols=2)
    att_tbl.style = "Table Grid"
    for ci, hdr in enumerate(["Name", "Role"]):
        c = att_tbl.rows[0].cells[ci]
        _cell_bg(c, GREEN_HEX)
        _run(c.paragraphs[0], hdr, bold=True, size_pt=10, color="FFFFFF")
    for ri, att in enumerate(att_rows):
        for ci, val in enumerate([att.get("name", ""), att.get("title", "")]):
            _run(att_tbl.rows[ri + 1].cells[ci].paragraphs[0], val, size_pt=10)

    _para()

    # Meeting Objective
    _section_heading("Meeting Objective")
    obj_text = (content.get("meeting_objective") or
                f"Discuss the proposed engagement and assess strategic value, "
                f"requirements, and next steps for {company}.")
    _para(obj_text, size_pt=11)
    _para()

    # Discussion Points
    _section_heading("Key Discussion Points")
    for pt in (discussion_en or [f"Discussion on strategic alignment with {company}."]):
        p = doc.add_paragraph(style="List Bullet")
        _run(p, pt, size_pt=10.5)
    _para()

    # Action Items table
    _section_heading("Action Items")
    if action_items:
        act_tbl = doc.add_table(rows=1 + len(action_items), cols=6)
        act_tbl.style = "Table Grid"
        hdrs6  = ["#", "Action Item", "Owner", "Deliverable", "Due Date", "Success Measure"]
        widths6 = [0.3, 2.5, 1.2, 1.2, 0.85, 1.45]
        for ci, (h, w) in enumerate(zip(hdrs6, widths6)):
            c = act_tbl.rows[0].cells[ci]
            _cell_bg(c, GREEN_HEX)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _run(p, h, bold=True, size_pt=9, color="FFFFFF")
            try:
                tcPr = c._tc.get_or_add_tcPr()
                tcW  = OxmlElement("w:tcW")
                tcW.set(qn("w:w"), str(int(w * 1440)))
                tcW.set(qn("w:type"), "dxa")
                tcPr.append(tcW)
            except Exception:
                pass
        for ri, ai in enumerate(action_items):
            task_en = (ai.get("task_en") or ai.get("Action (EN)") or
                       _translate_to_en(ai.get("task", "")) or "")
            owner   = ai.get("owner", "") or ai.get("Assigned To", "")
            due     = ai.get("due", "") or str(ai.get("Due Date", "") or "")
            deliverable = (task_en[:40].rsplit(" ", 1)[0] + "…") if len(task_en) > 40 else task_en
            vals = [str(ri + 1), task_en, owner, deliverable, due, "Task completed and reviewed"]
            bg   = "FFFFFF" if ri % 2 == 0 else "F5F5F5"
            for ci, val in enumerate(vals):
                c = act_tbl.rows[ri + 1].cells[ci]
                _cell_bg(c, bg)
                p = c.paragraphs[0]
                if ci == 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _run(p, val, size_pt=9)
    _para()

    # Next Steps
    _section_heading("Summary of Next Steps")
    for ai in (action_items or []):
        task_en = (ai.get("task_en") or ai.get("Action (EN)") or
                   _translate_to_en(ai.get("task", "")) or "")
        owner   = ai.get("owner", "") or ai.get("Assigned To", "")
        if task_en:
            line = f"{owner} to {task_en[0].lower()}{task_en[1:]}" if owner else task_en
            p = doc.add_paragraph(style="List Bullet")
            _run(p, line[:140], size_pt=10.5)
    _para()

    # Immediate Priority
    _section_heading("Immediate Priority")
    first   = (action_items[0] if action_items else {})
    imm     = (first.get("task_en") or first.get("Action (EN)") or
               _translate_to_en(first.get("task", "")) or
               f"Receive and review deliverables from {company} before wider stakeholder engagement.")
    _para(imm, size_pt=11)
    _para()

    # Footer
    ft_en = doc.add_paragraph()
    ft_en.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f1 = ft_en.add_run("Prepared by: Minister Outreach Office, MISA")
    r_f1.font.size = Pt(9); r_f1.font.color.rgb = _rgb("888888")
    r_f2 = ft_en.add_run("    For internal use only — Ministry of Investment of Saudi Arabia")
    r_f2.italic = True; r_f2.font.size = Pt(9); r_f2.font.color.rgb = _rgb("AAAAAA")

    buf_en = io.BytesIO()
    doc.save(buf_en)
    return buf_en.getvalue()


def _build_internal_am_letter_docx(cfg: dict, actions_df: pd.DataFrame) -> bytes:
    """
    Internal memo from RM to AM with action items to execute after a meeting.
    """
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    co      = cfg.get("company", "")
    mtg     = cfg.get("meeting_date", "") or date.today().strftime("%d %B %Y")
    arm     = cfg.get("arm", "") or "Account Manager"
    exec_rm = cfg.get("exec_rm", "") or ""
    yr_mon  = date.today().strftime("%Y-%m")

    doc = docx.Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(1.0)
    sec.top_margin  = sec.bottom_margin = Inches(0.75)

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

    def _run(para, text, bold=False, size_pt=11.0, color="1C1C1C"):
        r = para.add_run(text)
        r.bold = bold; r.font.size = Pt(size_pt)
        r.font.color.rgb = _rgb(color)
        return r

    def _para(text="", bold=False, size_pt=11.0, color="1C1C1C"):
        p = doc.add_paragraph()
        if text:
            _run(p, text, bold=bold, size_pt=size_pt, color=color)
        return p

    # Header
    hdr = doc.add_table(rows=2, cols=1)
    hdr.style = "Table Grid"
    c0 = hdr.rows[0].cells[0]
    _cell_bg(c0, "217141")
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p0, "Ministry of Investment  |  وزارة الاستثمار", bold=True, size_pt=12, color="FFFFFF")
    c1 = hdr.rows[1].cells[0]
    _cell_bg(c1, "1A5C3F")
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p1, "Minister's Office  |  Internal Memo", size_pt=10, color="C9974A")

    _para()

    ref_p = doc.add_paragraph()
    _run(ref_p, "Date: ", bold=True, size_pt=10, color="555555")
    _run(ref_p, mtg, size_pt=10, color="555555")
    _run(ref_p, f"     |     Ref: MISA / {co} / AM-INTERNAL / {yr_mon}",
         size_pt=10, color="999999")

    _para()

    for lbl, val in [("To: ", arm), ("From: ", exec_rm or "Minister's Office"),
                     ("Re: ", f"Action Items — {co} (Meeting: {mtg})")]:
        p = doc.add_paragraph()
        _run(p, lbl, bold=True, size_pt=11)
        _run(p, val, bold=(lbl == "Re: "), size_pt=11)

    _para()

    p_body = doc.add_paragraph()
    _run(p_body, "Dear ")
    _run(p_body, arm, bold=True)
    _run(p_body, ",")
    _para()
    p_body2 = doc.add_paragraph()
    _run(p_body2, "Based on the internal alignment following our meeting with ")
    _run(p_body2, co, bold=True)
    _run(p_body2, f" on {mtg}, please find the action items below that you need to start "
         "engaging with the relevant stakeholders to start delivering.")
    _para()

    if not actions_df.empty:
        tbl = doc.add_table(rows=1, cols=6)
        tbl.style = "Table Grid"
        hdrs = ["#", "Action Item", "Assigned To", "Priority", "Timeline", "Status"]
        widths = [0.3, 2.8, 1.3, 0.7, 0.9, 0.8]
        for j, (cell, h) in enumerate(zip(tbl.rows[0].cells, hdrs)):
            _cell_bg(cell, "217141")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _run(p, h, bold=True, size_pt=9, color="FFFFFF")
            try:
                tcPr = cell._tc.get_or_add_tcPr()
                tcW  = OxmlElement("w:tcW")
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
            bg     = "FFFFFF" if i % 2 == 0 else "F5F5F5"
            data_row = tbl.add_row()
            for j, val in enumerate([str(i + 1), en, owner, prio, due, status]):
                cell = data_row.cells[j]
                _cell_bg(cell, bg)
                p = cell.paragraphs[0]
                if j == 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _run(p, val, size_pt=9)

    _para()
    _para("Please acknowledge receipt and confirm your plan to initiate the above "
          "within the agreed timelines. Escalate any blockers to the Minister's Office "
          "immediately.", size_pt=11)
    _para()
    _para("Best regards,")
    _para()

    sig_n = doc.add_paragraph()
    _run(sig_n, exec_rm or "Minister's Office", bold=True, size_pt=11.5, color="217141")
    sig_t = doc.add_paragraph()
    _run(sig_t, "Relationship Manager (RM)  |  Minister's Office", size_pt=10.5, color="555555")
    _para("Ministry of Investment  |  Kingdom of Saudi Arabia", size_pt=10.5, color="555555")

    ftr = doc.add_table(rows=1, cols=1)
    ftr.style = "Table Grid"
    fc = ftr.rows[0].cells[0]
    _cell_bg(fc, "F5F5F5")
    fp = fc.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(fp, "CONFIDENTIAL — INTERNAL USE ONLY  |  Ministry of Investment",
         size_pt=9, color="888888")

    buf_am = io.BytesIO()
    doc.save(buf_am)
    return buf_am.getvalue()


def _build_pre_review_am_letter_docx(cfg: dict, actions_df: pd.DataFrame) -> bytes:
    """
    Initial internal memo from RM to AM — sent immediately after a minister meeting.
    Shares the draft action items and requests an alignment session before stakeholder engagement.
    """
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    co      = cfg.get("company", "")
    mtg     = cfg.get("meeting_date", "") or date.today().strftime("%d %B %Y")
    arm     = cfg.get("arm", "") or "Account Manager"
    exec_rm = cfg.get("exec_rm", "") or ""
    yr_mon  = date.today().strftime("%Y-%m")

    doc = docx.Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(1.0)
    sec.top_margin  = sec.bottom_margin = Inches(0.75)

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

    def _run(para, text, bold=False, size_pt=11.0, color="1C1C1C"):
        r = para.add_run(text)
        r.bold = bold; r.font.size = Pt(size_pt)
        r.font.color.rgb = _rgb(color)
        return r

    def _para(text="", bold=False, size_pt=11.0, color="1C1C1C"):
        p = doc.add_paragraph()
        if text:
            _run(p, text, bold=bold, size_pt=size_pt, color=color)
        return p

    # Header
    hdr = doc.add_table(rows=2, cols=1)
    hdr.style = "Table Grid"
    c0 = hdr.rows[0].cells[0]
    _cell_bg(c0, "217141")
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p0, "Ministry of Investment  |  وزارة الاستثمار", bold=True, size_pt=12, color="FFFFFF")
    c1 = hdr.rows[1].cells[0]
    _cell_bg(c1, "1A5C3F")
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p1, "Minister's Office  |  Internal Memo", size_pt=10, color="C9974A")

    _para()

    ref_p = doc.add_paragraph()
    _run(ref_p, "Date: ", bold=True, size_pt=10, color="555555")
    _run(ref_p, mtg, size_pt=10, color="555555")
    _run(ref_p, f"     |     Ref: MISA / {co} / AM-INTERNAL / {yr_mon}",
         size_pt=10, color="999999")

    _para()

    for lbl, val in [("To: ", arm), ("From: ", exec_rm or "Minister's Office"),
                     ("Re: ", f"Action Items — {co} (Meeting: {mtg})")]:
        p = doc.add_paragraph()
        _run(p, lbl, bold=True, size_pt=11)
        _run(p, val, bold=(lbl == "Re: "), size_pt=11)

    _para()

    # Salutation
    p_sal = doc.add_paragraph()
    _run(p_sal, "Dear ")
    _run(p_sal, arm, bold=True)
    _run(p_sal, ",")
    _para()

    # Body — pre-review version
    p_b1 = doc.add_paragraph()
    _run(p_b1, "Following the Minister's meeting with ")
    _run(p_b1, co, bold=True)
    _run(p_b1, f" on {mtg}, please find below the initial action items captured during "
         "the meeting.")
    _para()

    p_b2 = doc.add_paragraph()
    _run(p_b2, "Prior to sharing these with the relevant stakeholders for delivery, we "
         "would like to schedule a brief ")
    _run(p_b2, "Action Item Review Session", bold=True)
    _run(p_b2, " with you to walk through each item, ensure full alignment, and confirm "
         "your acknowledgement.")
    _para()

    # Action items table
    if not actions_df.empty:
        tbl = doc.add_table(rows=1, cols=6)
        tbl.style = "Table Grid"
        hdrs   = ["#", "Action Item", "Assigned To", "Priority", "Timeline", "Status"]
        widths = [0.3, 2.8, 1.3, 0.7, 0.9, 0.8]
        for j, (cell, h) in enumerate(zip(tbl.rows[0].cells, hdrs)):
            _cell_bg(cell, "217141")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _run(p, h, bold=True, size_pt=9, color="FFFFFF")
            try:
                tcPr = cell._tc.get_or_add_tcPr()
                tcW  = OxmlElement("w:tcW")
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
            bg     = "FFFFFF" if i % 2 == 0 else "F5F5F5"
            data_row = tbl.add_row()
            for j, val in enumerate([str(i + 1), en, owner, prio, due, status]):
                cell = data_row.cells[j]
                _cell_bg(cell, bg)
                p = cell.paragraphs[0]
                if j == 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _run(p, val, size_pt=9)

    _para()
    _para("Kindly share your availability for the review session at your earliest "
          "convenience so we can align before engaging the relevant parties.",
          size_pt=11)
    _para()
    _para("Best regards,")
    _para()

    sig_n = doc.add_paragraph()
    _run(sig_n, exec_rm or "Minister's Office", bold=True, size_pt=11.5, color="217141")
    sig_t = doc.add_paragraph()
    _run(sig_t, "Relationship Manager (RM)  |  Minister's Office", size_pt=10.5, color="555555")
    _para("Ministry of Investment  |  Kingdom of Saudi Arabia", size_pt=10.5, color="555555")

    ftr = doc.add_table(rows=1, cols=1)
    ftr.style = "Table Grid"
    fc = ftr.rows[0].cells[0]
    _cell_bg(fc, "F5F5F5")
    fp = fc.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(fp, "CONFIDENTIAL — INTERNAL USE ONLY  |  Ministry of Investment",
         size_pt=9, color="888888")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
