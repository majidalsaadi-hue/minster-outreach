
# Report Builder — Ministry of Investment
# 3-step workflow: Upload → Configure → Generate
# Outputs: Extracted action items table + Internal summary + Email draft
# All processing server-side (python-docx + deep-translator + openpyxl)

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
_PRIORITY_AR = {"Very High": "مهم جدا", "High": "مهم", "Medium": "متوسط", "Low": "عادي"}
_PRIORITY_EN = {"مهم جدا": "Very High", "مهم": "High", "متوسط": "Medium", "عادي": "Low",
                "مستمر": "Ongoing"}

_EMPTY_ACTIONS = pd.DataFrame(columns=[
    "Action (AR)", "Action (EN)", "Assigned To", "Type", "Priority", "Due Date", "Remarks"
])

_BADGE_CSS = {
    "Very High": "background:#fee2e2;color:#dc2626",
    "High":      "background:#fee2e2;color:#dc2626",
    "Medium":    "background:#fef3c7;color:#d97706",
    "Low":       "background:#f3f4f6;color:#6b7280",
    "Completed": "background:#d1fae5;color:#065f46",
    "Not Started":"background:#f3f4f6;color:#6b7280",
    "Inprogress":"background:#fef3c7;color:#d97706",
    "In Progress":"background:#fef3c7;color:#d97706",
    "Blocked":   "background:#fee2e2;color:#dc2626",
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
        "rb_parsed":      None,
        "rb_actions":     _EMPTY_ACTIONS.copy(),
        "rb_company":     "",
        "rb_recipient":   "",
        "rb_arm":         "",
        "rb_exec_rm":     "",
        "rb_date":        date.today().strftime("%d %B %Y"),
        "rb_subject_ar":  "",
        "rb_subject_en":  "",
        "rb_location":    "المقر الرئيسي – وزارة الاستثمار",
        "rb_chair":       "معالي الوزير",
        "rb_next_mtg":    None,
        "rb_attendees":   "",
        "rb_disc_ar":     "",
        "rb_disc_en":     "",
        "rb_output_lang": "English",
        "rb_excel_bytes": None,
        "rb_result":      None,   # dict after pipeline runs
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ─── CSS injection ────────────────────────────────────────────────────────────

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

    # ── Header ────────────────────────────────────────────────────────────────
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
                unsafe_allow_html=True,
            )
            st.caption("Standard Ministry of Investment meeting template")
            word_file = st.file_uploader("word", type=["docx", "doc"],
                                         key="rb_word_up", label_visibility="collapsed")

        with u2:
            st.markdown(
                '<span class="rb-num">2</span>'
                '<strong style="font-size:12px">Action Item Tracker (.xlsx)</strong>',
                unsafe_allow_html=True,
            )
            st.caption("V5 tracker — format preserved, only relevant sheet updated")
            excel_file = st.file_uploader("excel", type=["xlsx", "xls"],
                                          key="rb_excel_up", label_visibility="collapsed")

    # Auto-parse Word on upload
    if word_file is not None:
        raw = word_file.read()
        with st.spinner("Parsing document and translating action items…"):
            parsed = _parse_word(raw)
        if parsed:
            st.session_state["rb_parsed"]    = parsed
            st.session_state["rb_company"]   = parsed.get("company", "")
            st.session_state["rb_date"]      = (
                parsed["date"].strftime("%d %B %Y") if parsed.get("date") else date.today().strftime("%d %B %Y")
            )
            st.session_state["rb_subject_ar"] = parsed.get("subject_ar", "")
            st.session_state["rb_subject_en"] = parsed.get("subject_en", "")
            st.session_state["rb_location"]   = parsed.get("location", "المقر الرئيسي – وزارة الاستثمار")
            st.session_state["rb_chair"]      = parsed.get("chair", "معالي الوزير")
            st.session_state["rb_attendees"]  = parsed.get("attendees", "")
            st.session_state["rb_disc_ar"]    = parsed.get("discussion_ar", "")
            if parsed.get("action_items"):
                items   = parsed["action_items"]
                ar_texts = [i.get("Action (AR)", "") for i in items]
                en_texts = _translate_list(ar_texts)
                for i, en in enumerate(en_texts):
                    if not items[i].get("Action (EN)"):
                        items[i]["Action (EN)"] = en
                st.session_state["rb_actions"] = pd.DataFrame(items)
            n = len(parsed.get("action_items", []))
            st.success(f"✅ Parsed — {n} action item(s) extracted and translated.")

    if excel_file is not None:
        st.session_state["rb_excel_bytes"] = excel_file.read()
        sheets = _get_excel_sheets(st.session_state["rb_excel_bytes"])
        st.info(f"Excel loaded — sheets: {', '.join(sheets)}")

    # ── Step 2: Configure ─────────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown('<p class="rb-section">Step 2 — Configure output</p>', unsafe_allow_html=True)

        investors    = dfs.get("Investor Master", pd.DataFrame())
        company_list = sorted(investors["Company Name"].dropna().unique().tolist()) \
            if not investors.empty and "Company Name" in investors.columns else []

        r1c1, r1c2 = st.columns(2)
        with r1c1:
            if company_list:
                idx = company_list.index(st.session_state["rb_company"]) \
                    if st.session_state["rb_company"] in company_list else 0
                company = st.selectbox("Company", company_list, index=idx, key="rb_co_sel")
            else:
                company = st.text_input("Company",
                                        value=st.session_state["rb_company"],
                                        placeholder="e.g. Barclays", key="rb_co_txt")
            st.session_state["rb_company"] = company

        with r1c2:
            st.session_state["rb_recipient"] = st.text_input(
                "Email recipient name",
                value=st.session_state["rb_recipient"],
                placeholder="e.g. Khalid Al-Dabbagh", key="rb_recip",
            )

        r2c1, r2c2 = st.columns(2)
        with r2c1:
            st.session_state["rb_arm"] = st.text_input(
                "ARM (Account Relationship Manager)",
                value=st.session_state["rb_arm"],
                placeholder="e.g. Dana Aljarbu", key="rb_arm_inp",
            )
        with r2c2:
            st.session_state["rb_exec_rm"] = st.text_input(
                "Executive RM (Minister's Office)",
                value=st.session_state["rb_exec_rm"],
                placeholder="e.g. Sara Al-Sayed", key="rb_exec_inp",
            )

        r3c1, r3c2 = st.columns(2)
        with r3c1:
            st.session_state["rb_date"] = st.text_input(
                "Meeting date",
                value=st.session_state["rb_date"],
                placeholder=f"e.g. {date.today().strftime('%d %B %Y')}",
                key="rb_date_inp",
            )
        with r3c2:
            st.session_state["rb_output_lang"] = st.selectbox(
                "Output language",
                ["English", "Arabic", "Bilingual (EN + AR)"],
                key="rb_lang_sel",
            )

    # ── Step 3: Generate ──────────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown('<p class="rb-section">Step 3 — Generate</p>', unsafe_allow_html=True)

        run_btn = st.button(
            "▶  Run pipeline — extract actions, update Excel & draft email",
            type="primary", use_container_width=True, key="rb_run",
        )

        if run_btn:
            s = st.session_state
            if not word_file and not s.get("rb_parsed") and not s.get("rb_actions", _EMPTY_ACTIONS.copy()).shape[0]:
                st.warning("Upload at least the Word meeting minutes file.")
            else:
                prog   = st.progress(0)
                status = st.empty()

                def _log(msg: str, pct: int = None):
                    status.markdown(f"→ {msg}")
                    if pct is not None:
                        prog.progress(pct)

                try:
                    _log("Reading files…", 10)
                    actions  = _valid_actions(s["rb_actions"])
                    company  = s["rb_company"] or "Company"
                    arm      = s["rb_arm"] or "ARM"
                    exec_rm  = s["rb_exec_rm"] or "Executive RM"
                    recipient = s["rb_recipient"] or "Dear Sir/Madam"
                    mtg_date = s["rb_date"] or date.today().strftime("%d %B %Y")

                    _log(f"Processing {len(actions)} action item(s) for {company}…", 35)

                    cfg = {
                        "company":   company,
                        "recipient": recipient,
                        "arm":       arm,
                        "exec_rm":   exec_rm,
                        "meeting_date": mtg_date,
                        "chair":     s["rb_chair"] or "H.E. The Minister",
                        "next_mtg":  s["rb_next_mtg"],
                        "disc_en":   s["rb_disc_en"] or s["rb_disc_ar"] or "",
                    }

                    _log("Generating internal summary…", 55)
                    summary = _build_internal_summary(cfg, actions)

                    _log("Drafting outreach email…", 72)
                    email_body = _build_email_from_cfg(cfg, actions)

                    updated_xl = None
                    if s.get("rb_excel_bytes") and not actions.empty:
                        _log("Updating Excel tracker…", 85)
                        try:
                            mtg_d = datetime.strptime(mtg_date, "%d %B %Y").date() if mtg_date else date.today()
                        except ValueError:
                            mtg_d = date.today()
                        updated_xl = _build_excel(
                            existing_bytes=s["rb_excel_bytes"],
                            company=company,
                            meeting_date=mtg_d,
                            next_meeting=s.get("rb_next_mtg"),
                            chair=cfg["chair"],
                            actions_df=actions,
                        )

                    _log("Syncing to CRM…", 93)
                    _sync_actions_to_crm(dfs, actions, company)

                    prog.progress(100)
                    status.success("✓ Pipeline complete")

                    st.session_state["rb_result"] = {
                        "actions":    actions,
                        "summary":    summary,
                        "email":      email_body,
                        "xl_bytes":   updated_xl,
                        "company":    company,
                        "meeting_date": mtg_date,
                    }

                except Exception as e:
                    status.error(f"Error: {e}")

    # ── Output section ────────────────────────────────────────────────────────
    result = st.session_state.get("rb_result")
    if result:
        _render_output(result)


# ─── Output renderer ──────────────────────────────────────────────────────────

def _render_output(result: dict):
    st.markdown('<hr style="margin:1.5rem 0;border:none;border-top:0.5px solid #e5e7eb">',
                unsafe_allow_html=True)

    actions    = result["actions"]
    summary    = result["summary"]
    email_body = result["email"]
    xl_bytes   = result["xl_bytes"]
    company    = result["company"]
    mtg_date   = result["meeting_date"]

    # ── Action items table ────────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown("#### ✅ Extracted action items")
        if not actions.empty:
            _render_action_table(actions)
        else:
            st.info("No action items extracted.")
        if xl_bytes:
            st.download_button(
                "⬇ Download updated Excel tracker",
                data=xl_bytes,
                file_name=f"ActionTracker_{company.replace(' ','_')}_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="rb_dl_xl",
            )

    # ── Internal summary ──────────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown("#### 📄 Internal summary")
        edited_summary = st.text_area(
            "", value=summary, height=300, key="rb_sum_ed", label_visibility="collapsed"
        )
        c1, _ = st.columns([1, 4])
        c1.download_button(
            "⬇ Download",
            data=edited_summary.encode("utf-8"),
            file_name=f"Summary_{company.replace(' ','_')}_{mtg_date}.txt",
            mime="text/plain",
            key="rb_dl_sum",
        )

    # ── Email draft ───────────────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown("#### ✉️ Email draft — ready to send")
        edited_email = st.text_area(
            "", value=email_body, height=360, key="rb_email_ed", label_visibility="collapsed"
        )
        c1, _ = st.columns([1, 4])
        c1.download_button(
            "⬇ Download",
            data=edited_email.encode("utf-8"),
            file_name=f"Email_{company.replace(' ','_')}_{mtg_date}.txt",
            mime="text/plain",
            key="rb_dl_email",
        )


def _render_action_table(df: pd.DataFrame):
    rows_html = ""
    for i, (_, row) in enumerate(df.iterrows()):
        en     = row.get("Action (EN)") or row.get("Action (AR)", "")
        owner  = str(row.get("Assigned To", "") or "")
        prio   = str(row.get("Priority", "Medium") or "Medium")
        due    = str(row.get("Due Date", "TBD") or "TBD")
        status = str(row.get("Status", "Not Started") or "Not Started")
        rmk    = str(row.get("Remarks", "") or "")

        p_css = _BADGE_CSS.get(prio, "background:#f3f4f6;color:#6b7280")
        s_css = _BADGE_CSS.get(status, "background:#f3f4f6;color:#6b7280")

        rows_html += (
            f"<tr>"
            f"<td style='font-weight:500;text-align:center;width:36px'>{i+1}</td>"
            f"<td>{en}</td>"
            f"<td>{owner}</td>"
            f"<td><span class='rb-pk' style='{p_css}'>{prio}</span></td>"
            f"<td>{due}</td>"
            f"<td><span class='rb-pk' style='{s_css}'>{status}</span></td>"
            f"<td style='font-size:11px;color:#6b7280'>{rmk}</td>"
            f"</tr>"
        )

    st.markdown(f"""
    <div style="overflow-x:auto">
    <table class="rb-action-table">
      <thead><tr>
        <th>#</th><th>Action Item</th><th>Owner</th>
        <th>Priority</th><th>Due Date</th><th>Status</th><th>Remarks</th>
      </tr></thead>
      <tbody>{rows_html}</tbody>
    </table>
    </div>
    """, unsafe_allow_html=True)


# ─── Summary & email builders ──────────────────────────────────────────────────

def _build_internal_summary(cfg: dict, actions: pd.DataFrame) -> str:
    co      = cfg["company"]
    mtg     = cfg["meeting_date"]
    arm     = cfg["arm"]
    exec_rm = cfg["exec_rm"]
    chair   = cfg["chair"]
    disc    = cfg.get("disc_en", "")
    yr      = date.today().year

    lines = [
        "MINISTRY OF INVESTMENT — MEETING SUMMARY",
        "=" * 56,
        f"Company:        {co}",
        f"Meeting Date:   {mtg}",
        f"Chaired By:     {chair}",
        f"ARM:            {arm}",
        f"Executive RM:   {exec_rm}",
        f"Reference:      MISA/{co.upper()[:6].replace(' ','')}/ARM/{yr}",
        "",
        "KEY DISCUSSION POINTS",
        "-" * 56,
    ]

    disc_lines = [ln.strip().lstrip("•● -") for ln in disc.split("\n") if ln.strip()] if disc else []
    if disc_lines:
        for ln in disc_lines:
            lines.append(f"  • {ln}")
    else:
        lines.append("  (See attached meeting minutes for full discussion)")

    lines += ["", "AGREED ACTION ITEMS", "-" * 56]
    if not actions.empty:
        for i, (_, row) in enumerate(actions.iterrows(), 1):
            en    = row.get("Action (EN)") or row.get("Action (AR)", "")
            owner = row.get("Assigned To", "TBD")
            prio  = row.get("Priority", "Medium")
            due   = str(row.get("Due Date", "TBD")) if row.get("Due Date") else "TBD"
            lines.append(f"  {i}. {en}")
            lines.append(f"     Owner: {owner}  |  Priority: {prio}  |  Due: {due}")
    else:
        lines.append("  (No action items recorded)")

    lines += [
        "",
        "PREPARED BY",
        "-" * 56,
        f"  {arm}, Account Relationship Manager",
        f"  {exec_rm}, Executive Relationship Manager",
        f"  Ministry of Investment — وزارة الاستثمار",
        f"  Date: {date.today().strftime('%d %B %Y')}",
    ]

    return "\n".join(lines)


def _build_email_from_cfg(cfg: dict, actions: pd.DataFrame) -> str:
    co       = cfg["company"]
    mtg      = cfg["meeting_date"]
    arm      = cfg["arm"]
    exec_rm  = cfg["exec_rm"]
    chair    = cfg["chair"]
    recipient = cfg["recipient"]
    sub_en   = f"Follow-up on Latest Updates — {co}"
    yr       = date.today().year
    ref      = f"MISA/{co.upper()[:6].replace(' ','')}/ARM/{yr}"

    subject = f"SUBJECT: {sub_en} | Ref: {ref}"

    intro = (
        f"Dear {recipient},\n\n"
        f"I hope this message finds you well.\n\n"
        f"On behalf of H.E. {chair} and the Ministry of Investment of Saudi Arabia, "
        f"I am writing to follow up on our productive meeting held on {mtg}. "
        f"We appreciate your continued engagement and strategic partnership with the Kingdom.\n\n"
        f"As part of our ongoing commitment to investor support, {arm} has been assigned "
        f"as your dedicated Account Relationship Manager (ARM), "
        f"with {exec_rm} serving as Executive Relationship Manager from the Minister's Office."
    )

    act_block = "\n\nAGREED ACTION ITEMS\n" + "-" * 56
    if not actions.empty:
        act_block += (
            f"\n{'#':<4} {'Action Item':<55} {'Owner':<22} {'Due':<14} {'Priority'}"
            f"\n{'-'*4} {'-'*55} {'-'*22} {'-'*14} {'-'*10}"
        )
        for i, (_, row) in enumerate(actions.iterrows(), 1):
            en    = (row.get("Action (EN)") or row.get("Action (AR)", ""))[:54]
            owner = str(row.get("Assigned To", "TBD"))
            due   = str(row.get("Due Date", "TBD")) if row.get("Due Date") else "TBD"
            prio  = str(row.get("Priority", "Medium"))
            act_block += f"\n{i:<4} {en:<55} {owner:<22} {due:<14} {prio}"
    else:
        act_block += "\n  (No action items recorded)"

    closing = (
        "\n\nNEXT STEPS\n" + "-" * 56 +
        "\nWe look forward to continued progress on the above action items. "
        "Please do not hesitate to contact us for any clarifications or support.\n\n"
        f"For day-to-day coordination, please reach out to {arm} (ARM).\n\n"
        "Best regards,\n\n"
        f"{exec_rm}\n"
        "Executive Relationship Manager, Minister's Office\n"
        "Ministry of Investment of Saudi Arabia\n"
        "وزارة الاستثمار — المملكة العربية السعودية"
    )

    return "\n".join([subject, "", intro, act_block, closing])


# ─── CRM sync (direct, no UI) ─────────────────────────────────────────────────

def _sync_actions_to_crm(dfs: dict, actions: pd.DataFrame, company: str):
    if actions.empty or not company:
        return
    investors = dfs.get("Investor Master", pd.DataFrame())
    inv_id    = _investor_id(investors, company)
    existing  = dfs.get("Action Items", pd.DataFrame())
    new_rows  = []

    for i, (_, row) in enumerate(actions.iterrows()):
        desc   = row.get("Action (EN)") or row.get("Action (AR)", "")
        if not desc:
            continue
        # Skip if already in CRM
        if not existing.empty and "Action Description" in existing.columns:
            if (existing["Action Description"].astype(str).str.strip() == str(desc).strip()).any():
                continue
        new_id = _next_act_id(pd.concat([existing, pd.DataFrame(new_rows)], ignore_index=True) if new_rows else existing, 0)
        due = row.get("Due Date")
        new_rows.append({
            "Action ID":          new_id,
            "Investor ID":        inv_id,
            "Company Name":       company,
            "Action Description": desc,
            "Assigned To":        row.get("Assigned To", ""),
            "Sector":             "",
            "Type of Engagement": row.get("Type", "Action"),
            "Start Date":         date.today(),
            "Due Date":           due if (due and pd.notna(due)) else None,
            "Priority":           _map_prio(row.get("Priority", "Medium")),
            "Progress":           "0%",
            "Status":             "Not Started",
            "Escalation Flag":    "None",
            "Remarks":            row.get("Remarks", ""),
            "Last Updated":       date.today(),
        })
        existing = pd.concat([existing, pd.DataFrame([new_rows[-1]])], ignore_index=True)

    if new_rows:
        dfs["Action Items"] = existing
        save_session(dfs)


# ─── Word parser ──────────────────────────────────────────────────────────────

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

    tables = []
    for tbl in doc.tables:
        rows = []
        for row in tbl.rows:
            cells = [c.text.strip() for c in row.cells]
            rows.append(cells)
        tables.append(rows)

    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    for tbl in tables[:3]:
        flat = " ".join(" ".join(r) for r in tbl)
        m = re.search(r"(\d{4}[-/]\d{1,2}[-/]\d{1,2})", flat)
        if m and not result["date"]:
            try:
                result["date"] = datetime.strptime(m.group(1).replace("/", "-"), "%Y-%m-%d").date()
            except ValueError:
                pass
        for row in tbl:
            joined = " ".join(row)
            if any(kw in joined for kw in ["مستجدات", "اجتماع", "تحديث", "متابعة"]):
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

    known = ["Barclays", "BlackRock", "Brookfield", "Goldman", "HSBC",
             "JPMorgan", "Morgan Stanley", "UBS", "Citi", "Deutsche"]
    full_text = " ".join(paras)
    for name in known:
        if name.lower() in full_text.lower():
            result["company"] = name
            if not result["subject_en"]:
                result["subject_en"] = f"Latest Updates — {name}"
            break

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

    for tbl in tables:
        if not tbl:
            continue
        header_flat = " ".join(tbl[0])
        is_actions  = any(kw in header_flat for kw in ["المهمة", "التوجيه", "Action Item", "ID"])
        if not is_actions and len(tbl) >= 2:
            is_actions = any(kw in " ".join(tbl[1]) for kw in ["المهمة", "التوجيه"])
        if not is_actions:
            continue
        for row in tbl[1:]:
            cells      = [c.strip() for c in row if c.strip()]
            text_cells = [c for c in cells if not re.match(r"^\d+$", c)]
            if not text_cells:
                continue
            action_ar = text_cells[0] if text_cells else ""
            owner, prio, due = "", "High", None
            for cell in text_cells[1:]:
                if re.search(r"(\d{4}[-/]\d{1,2}[-/]\d{1,2})", cell):
                    try:
                        m = re.search(r"(\d{4}[-/]\d{1,2}[-/]\d{1,2})", cell)
                        if m:
                            due = datetime.strptime(m.group(1).replace("/", "-"), "%Y-%m-%d").date()
                    except Exception:
                        pass
                elif cell in _PRIORITY_EN:
                    prio = _PRIORITY_EN[cell]
                elif not owner and len(cell) > 1 and cell not in _PRIORITY_AR.values():
                    owner = cell
            if action_ar and action_ar not in ("م", "التوجيه / المهمة", "Action Item"):
                result["action_items"].append({
                    "Action (AR)": action_ar, "Action (EN)": "",
                    "Assigned To": owner, "Type": "Action",
                    "Priority": prio, "Due Date": due, "Remarks": "",
                })

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


# ─── Excel builder ────────────────────────────────────────────────────────────

def _build_excel(existing_bytes, company, meeting_date, next_meeting, chair, actions_df) -> bytes:
    if existing_bytes:
        wb = openpyxl.load_workbook(io.BytesIO(existing_bytes))
    else:
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    sheet_name = company[:31]

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
        r = start_row + i
        due = row.get("Due Date")
        due_val  = due if (due and str(due) not in ("NaT", "None", "")) else None
        prio_val = _map_prio(row.get("Priority", "Medium"))

        ws.cell(r, 7,  i + 1)
        ws.cell(r, 8,  row.get("Action (EN)") or row.get("Action (AR)", ""))
        ws.cell(r, 9,  row.get("Assigned To", ""))
        ws.cell(r, 10, row.get("Type", "Action"))
        ws.cell(r, 11, meeting_date)
        ws.cell(r, 12, due_val)
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

    ws.cell(13, 11, "Outreach").font = white_font
    ws.cell(13, 11).fill = green_fill
    ws.cell(13, 15, "Last Updated").font = white_font
    ws.cell(13, 15).fill = green_fill
    ws.cell(13, 16, meeting_date).font = white_font
    ws.cell(13, 16).fill = gold_fill
    ws.cell(14, 11, "Manager").font = white_font
    ws.cell(14, 11).fill = green_fill
    ws.cell(14, 12, chair).font = Font(bold=True, size=10)
    ws.cell(16, 15, "Next Meeting").font = white_font
    ws.cell(16, 15).fill = green_fill
    ws.cell(16, 16, next_meeting or "TBD").font = gold_font
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


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _valid_actions(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    mask = (
        (df.get("Action (EN)", pd.Series(dtype=str)).notna() & (df.get("Action (EN)", pd.Series(dtype=str)) != "")) |
        (df.get("Action (AR)", pd.Series(dtype=str)).notna() & (df.get("Action (AR)", pd.Series(dtype=str)) != ""))
    )
    return df[mask]


def _map_prio(p: str) -> str:
    return {"Very High": "High", "High": "High", "Medium": "Medium", "Low": "Low"}.get(p, "Medium")


def _parse_lines(text: str) -> list:
    if not text:
        return []
    return [ln.strip().lstrip("•●-– ") for ln in text.split("\n") if ln.strip()]


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
    return f"ACT-{(max(nums) + 1 + offset if nums else 1 + offset):03d}"
