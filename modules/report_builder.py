
# Report Builder — compose meeting minutes → internal Arabic report + external English email
# → sync action items directly into the CRM Action Items tracker.

import io
import pandas as pd
import streamlit as st
from datetime import date

from config.settings import MISA_GREEN, MISA_GOLD, MISA_GREEN_DARK, ENGAGEMENT_TYPES
from config.translations import t
from modules.persistence import save_session


_PRIORITIES  = ["Very High", "High", "Medium", "Low"]
_ENGAGE_OPTS = ["Support", "Opportunity", "Challenge", "Follow-up", "Action", "Administrative"]

_PRIORITY_AR = {
    "Very High": "مهم جدا",
    "High":      "مهم",
    "Medium":    "متوسط",
    "Low":       "عادي",
}

_DAY_AR = {
    "Monday": "الاثنين", "Tuesday": "الثلاثاء", "Wednesday": "الأربعاء",
    "Thursday": "الخميس", "Friday": "الجمعة", "Saturday": "السبت", "Sunday": "الأحد",
}

_EMPTY_ACTIONS = pd.DataFrame(columns=[
    "Action (AR)", "Action (EN)", "Assigned To", "Type", "Priority", "Due Date", "Remarks"
])


def _init():
    keys = {
        "rb_company":       "",
        "rb_date":          date.today(),
        "rb_priority":      "Very High",
        "rb_location":      "المقر الرئيسي – وزارة الاستثمار",
        "rb_chair":         "معالي الوزير",
        "rb_next_mtg":      None,
        "rb_subject_ar":    "",
        "rb_subject_en":    "",
        "rb_attendees":     "",
        "rb_disc_ar":       "",
        "rb_disc_en":       "",
        "rb_actions":       _EMPTY_ACTIONS.copy(),
    }
    for k, v in keys.items():
        if k not in st.session_state:
            st.session_state[k] = v


def render(dfs: dict, lang: str):
    _init()

    st.markdown("### 📝 Report Builder")
    st.caption(
        "Fill in the meeting details below → switch to the output tabs to preview, copy, or "
        "download the formatted report and email → use Sync to add action items to the CRM tracker."
    )

    investors  = dfs.get("Investor Master", pd.DataFrame())
    companies  = [""] + (
        sorted(investors["Company Name"].dropna().unique().tolist())
        if not investors.empty and "Company Name" in investors.columns else []
    )

    # ── Input section ──────────────────────────────────────────────────────────
    with st.expander("📋 Step 1 — Meeting Details", expanded=True):
        _meta_form(companies)

    with st.expander("💬 Step 2 — Discussion Points / نقاط النقاش", expanded=True):
        _discussion_form()

    with st.expander("✅ Step 3 — Action Items / بنود العمل", expanded=True):
        _action_items_form()

    st.markdown("---")

    # ── Output tabs ────────────────────────────────────────────────────────────
    tab1, tab2, tab3 = st.tabs([
        "📄 Internal Report (Arabic)",
        "✉️ External Email (English)",
        "🔄 Sync to CRM Tracker",
    ])

    with tab1:
        _internal_report_tab()

    with tab2:
        _email_tab()

    with tab3:
        _sync_tab(dfs, lang)


# ── Step 1: Meeting metadata ───────────────────────────────────────────────────

def _meta_form(companies: list):
    c1, c2, c3 = st.columns(3)
    st.session_state["rb_company"]  = c1.selectbox(
        "Company", companies,
        index=companies.index(st.session_state["rb_company"]) if st.session_state["rb_company"] in companies else 0,
        key="_rb_company_sel"
    )
    st.session_state["rb_date"]     = c2.date_input("Meeting Date", value=st.session_state["rb_date"], key="_rb_date")
    st.session_state["rb_priority"] = c3.selectbox(
        "Priority / الأولوية", _PRIORITIES,
        index=_PRIORITIES.index(st.session_state["rb_priority"]),
        key="_rb_priority"
    )

    c4, c5, c6 = st.columns(3)
    st.session_state["rb_location"] = c4.text_input("Location / الموقع", value=st.session_state["rb_location"], key="_rb_loc")
    st.session_state["rb_chair"]    = c5.text_input("Chaired By / برئاسة", value=st.session_state["rb_chair"], key="_rb_chair")
    st.session_state["rb_next_mtg"] = c6.date_input("Next Meeting / القادم", value=st.session_state["rb_next_mtg"], key="_rb_next")

    c7, c8 = st.columns(2)
    st.session_state["rb_subject_ar"] = c7.text_input(
        "Subject (Arabic) / الموضوع", value=st.session_state["rb_subject_ar"], key="_rb_sub_ar"
    )
    st.session_state["rb_subject_en"] = c8.text_input(
        "Subject (English)", value=st.session_state["rb_subject_en"], key="_rb_sub_en"
    )

    st.markdown("**Attendees / الحضور** — one per line: Name | Title")
    st.session_state["rb_attendees"] = st.text_area(
        "Attendees", value=st.session_state["rb_attendees"], height=100,
        key="_rb_att",
        placeholder="Sara Al-Sayed | Consultant\nZiad Al-Juhaiman | Minister's Office Supervisor\nKhalid Al-Dabbagh | Co-CEO Barclays",
        label_visibility="collapsed",
    )


# ── Step 2: Discussion points ─────────────────────────────────────────────────

def _discussion_form():
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Arabic — paste meeting notes directly**")
        st.session_state["rb_disc_ar"] = st.text_area(
            "Arabic discussion", value=st.session_state["rb_disc_ar"], height=220,
            key="_rb_disc_ar", label_visibility="collapsed",
            placeholder="• أشار معالي الوزير إلى ...\n• أكدت الشركة ...\n• تم الاتفاق على ..."
        )
    with c2:
        st.markdown("**English — leave blank to use Arabic text as-is**")
        st.session_state["rb_disc_en"] = st.text_area(
            "English discussion", value=st.session_state["rb_disc_en"], height=220,
            key="_rb_disc_en", label_visibility="collapsed",
            placeholder="• The Minister noted that ...\n• Barclays confirmed ...\n• Both parties agreed to ..."
        )


# ── Step 3: Action items table ────────────────────────────────────────────────

def _action_items_form():
    st.caption("Add each action item from the meeting. 'Action (AR)' is for the internal report; 'Action (EN)' for the email.")

    edited = st.data_editor(
        st.session_state["rb_actions"],
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
        key="_rb_actions_editor",
        hide_index=True,
    )
    st.session_state["rb_actions"] = edited

    col_clear, _ = st.columns([1, 4])
    if col_clear.button("🗑️ Clear all action items", key="_rb_clear"):
        st.session_state["rb_actions"] = _EMPTY_ACTIONS.copy()
        st.rerun()


# ── Tab 1: Internal Arabic report ─────────────────────────────────────────────

def _internal_report_tab():
    st.markdown("#### وقائع الاجتماع — Internal Meeting Minutes")

    html = _build_internal_html()
    st.markdown(html, unsafe_allow_html=True)

    # Plain-text download
    text = _build_internal_text()
    st.download_button(
        "⬇️ Download as .txt",
        data=text.encode("utf-8"),
        file_name=f"Meeting_Minutes_{st.session_state['rb_company']}_{st.session_state['rb_date']}.txt",
        mime="text/plain",
        use_container_width=False,
    )


def _build_internal_html() -> str:
    s     = st.session_state
    mtg   = s["rb_date"]
    day   = _DAY_AR.get(mtg.strftime("%A"), mtg.strftime("%A")) if mtg else "—"
    pri   = _PRIORITY_AR.get(s["rb_priority"], s["rb_priority"])
    next_m = s["rb_next_mtg"].strftime("%B %Y") if s["rb_next_mtg"] else "—"

    # Attendee rows
    att_rows = ""
    for i, line in enumerate(_parse_lines(s["rb_attendees"]), 1):
        parts = line.split("|")
        name  = parts[0].strip() if parts else line
        role  = parts[1].strip() if len(parts) > 1 else ""
        att_rows += f"<tr><td style='text-align:center;width:40px'>{i}</td><td>{name}</td><td>{role}</td></tr>"

    # Discussion bullets
    disc_ar = s["rb_disc_ar"] or ""
    disc_bullets = "".join(f"<li style='margin:4px 0'>{l}</li>" for l in _parse_lines(disc_ar)) if disc_ar else "<li>—</li>"

    # Action rows
    actions = s["rb_actions"]
    act_rows = ""
    for i, (_, row) in enumerate(actions.iterrows(), 1):
        if not row.get("Action (AR)"):
            continue
        due   = row.get("Due Date", "")
        due_s = str(due) if due and str(due) != "NaT" else "—"
        act_rows += (
            f"<tr>"
            f"<td style='text-align:center'>{i}</td>"
            f"<td style='text-align:right'>{row.get('Action (AR)','—')}</td>"
            f"<td>{row.get('Assigned To','—')}</td>"
            f"<td style='text-align:center'>{_PRIORITY_AR.get(row.get('Priority',''),'—')}</td>"
            f"<td style='text-align:center'>{due_s}</td>"
            f"</tr>"
        )
    if not act_rows:
        act_rows = "<tr><td colspan='5' style='text-align:center;color:#999'>لا توجد بنود عمل</td></tr>"

    return f"""
<div dir="rtl" style="font-family:'Arial',sans-serif;background:#fff;border:1px solid #ddd;
     border-radius:10px;padding:28px;max-width:820px;margin:0 auto;">

  <!-- Header -->
  <div style="display:flex;justify-content:space-between;align-items:center;
              border-bottom:3px solid {MISA_GREEN};padding-bottom:12px;margin-bottom:16px;">
    <div>
      <div style="font-size:20px;font-weight:700;color:{MISA_GREEN}">وزارة الاستثمار</div>
      <div style="font-size:12px;color:#666">Ministry of Investment</div>
      <div style="font-size:11px;color:#666">مكتب الوزير — Minster Office</div>
    </div>
    <div style="text-align:center">
      <div style="font-size:18px;font-weight:700;color:{MISA_GREEN}">وقائع الاجتماع</div>
      <div style="font-size:13px;color:#444">Meeting Minutes</div>
    </div>
  </div>

  <!-- Priority legend -->
  <div style="display:flex;gap:12px;justify-content:flex-end;margin-bottom:12px;font-size:12px;">
    <span>● عادي</span>
    <span style="color:{MISA_GOLD}">● متوسط</span>
    <span style="color:orange">● مهم</span>
    <span style="color:red">● مهم جدا</span>
  </div>

  <!-- Meeting metadata table -->
  <table style="width:100%;border-collapse:collapse;margin-bottom:16px;font-size:13px;">
    <thead>
      <tr style="background:{MISA_GREEN};color:white;">
        <th style="padding:8px;text-align:center">الموضوع</th>
        <th style="padding:8px;text-align:center">الموقع</th>
        <th style="padding:8px;text-align:center">اليوم</th>
        <th style="padding:8px;text-align:center">التاريخ</th>
      </tr>
    </thead>
    <tbody>
      <tr style="background:#f9f9f9;">
        <td style="padding:8px;text-align:center;font-weight:700">{s["rb_subject_ar"] or "—"}</td>
        <td style="padding:8px;text-align:center">{s["rb_location"]}</td>
        <td style="padding:8px;text-align:center">{day}</td>
        <td style="padding:8px;text-align:center">{str(mtg) if mtg else "—"}</td>
      </tr>
    </tbody>
  </table>

  <table style="width:100%;border-collapse:collapse;margin-bottom:20px;font-size:13px;">
    <thead>
      <tr style="background:{MISA_GREEN};color:white;">
        <th style="padding:8px;text-align:center">الاجتماع القادم</th>
        <th style="padding:8px;text-align:center">الأولوية</th>
        <th style="padding:8px;text-align:center">الاجتماع برئاسة</th>
        <th style="padding:8px;text-align:center">وقت الاجتماع</th>
      </tr>
    </thead>
    <tbody>
      <tr style="background:#f9f9f9;">
        <td style="padding:8px;text-align:center">{next_m}</td>
        <td style="padding:8px;text-align:center;font-weight:700;color:red">{pri}</td>
        <td style="padding:8px;text-align:center">{s["rb_chair"]}</td>
        <td style="padding:8px;text-align:center">—</td>
      </tr>
    </tbody>
  </table>

  <!-- Discussion points -->
  <div style="margin-bottom:20px;">
    <div style="font-size:15px;font-weight:700;color:{MISA_GREEN};border-bottom:1px solid #ddd;
                padding-bottom:6px;margin-bottom:10px;">أبرز نقاط الاجتماع:</div>
    <div style="font-size:13px;">
      <div style="margin-bottom:4px;font-weight:600">أبرز ما تم مناقشته:</div>
      <ul style="padding-right:20px;line-height:2">{disc_bullets}</ul>
    </div>
  </div>

  <!-- Action items -->
  <div style="margin-bottom:20px;">
    <div style="font-size:15px;font-weight:700;color:{MISA_GREEN};border-bottom:1px solid #ddd;
                padding-bottom:6px;margin-bottom:10px;">توجيهات معاليه وأبرز المهام:</div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      <thead>
        <tr style="background:{MISA_GREEN};color:white;">
          <th style="padding:7px;text-align:center;width:30px">م</th>
          <th style="padding:7px;text-align:center">التوجيه / المهمة</th>
          <th style="padding:7px;text-align:center;width:120px">المسؤول</th>
          <th style="padding:7px;text-align:center;width:80px">الأولوية</th>
          <th style="padding:7px;text-align:center;width:100px">تاريخ الإنجاز المتوقع</th>
        </tr>
      </thead>
      <tbody>{act_rows}</tbody>
    </table>
  </div>

  <!-- Attendees -->
  <div>
    <div style="font-size:15px;font-weight:700;color:{MISA_GREEN};border-bottom:1px solid #ddd;
                padding-bottom:6px;margin-bottom:10px;">الحضور:</div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      <thead>
        <tr style="background:{MISA_GREEN};color:white;">
          <th style="padding:7px;text-align:center;width:30px">#</th>
          <th style="padding:7px;text-align:right">الاسم</th>
          <th style="padding:7px;text-align:right">الوظيفة</th>
        </tr>
      </thead>
      <tbody>{att_rows}</tbody>
    </table>
  </div>

</div>
"""


def _build_internal_text() -> str:
    s      = st.session_state
    mtg    = s["rb_date"]
    day    = _DAY_AR.get(mtg.strftime("%A"), mtg.strftime("%A")) if mtg else "—"
    pri    = _PRIORITY_AR.get(s["rb_priority"], s["rb_priority"])
    next_m = s["rb_next_mtg"].strftime("%B %Y") if s["rb_next_mtg"] else "—"

    lines = [
        "وقائع الاجتماع — Meeting Minutes",
        "وزارة الاستثمار | Ministry of Investment | مكتب الوزير",
        "=" * 60,
        f"الموضوع:         {s['rb_subject_ar']}",
        f"التاريخ:          {mtg} ({day})",
        f"الموقع:           {s['rb_location']}",
        f"الاجتماع برئاسة: {s['rb_chair']}",
        f"الأولوية:         {pri}",
        f"الاجتماع القادم: {next_m}",
        "",
        "أبرز نقاط الاجتماع:",
        "-" * 40,
    ]
    for ln in _parse_lines(s["rb_disc_ar"]):
        lines.append(f"  • {ln}")

    lines += ["", "توجيهات معاليه وأبرز المهام:", "-" * 40]
    for i, (_, row) in enumerate(s["rb_actions"].iterrows(), 1):
        if not row.get("Action (AR)"):
            continue
        due = str(row.get("Due Date", "—")) if row.get("Due Date") else "—"
        lines.append(f"  {i}. {row['Action (AR)']}  |  {row.get('Assigned To','—')}  |  {due}")

    lines += ["", "الحضور:", "-" * 40]
    for i, ln in enumerate(_parse_lines(s["rb_attendees"]), 1):
        lines.append(f"  {i}. {ln}")

    return "\n".join(lines)


# ── Tab 2: External email ──────────────────────────────────────────────────────

def _email_tab():
    st.markdown("#### External Email — Professional English Draft")

    email_text = _build_email_text()

    st.text_area(
        "Email draft (copy from here)",
        value=email_text,
        height=500,
        key="_rb_email_preview",
    )

    st.download_button(
        "⬇️ Download as .txt",
        data=email_text.encode("utf-8"),
        file_name=f"Email_{st.session_state['rb_company']}_{st.session_state['rb_date']}.txt",
        mime="text/plain",
    )


def _build_email_text() -> str:
    s      = st.session_state
    co     = s["rb_company"] or "[Company]"
    mtg    = s["rb_date"]
    sub_en = s["rb_subject_en"] or s["rb_subject_ar"] or "Meeting Follow-up"
    chair  = s["rb_chair"]
    next_m = s["rb_next_mtg"].strftime("%d %B %Y") if s["rb_next_mtg"] else "TBD"
    disc   = s["rb_disc_en"] or s["rb_disc_ar"] or ""

    subject_line = f"SUBJECT: Follow-up on {sub_en} — {co} | Ministry of Investment | {mtg}"

    greeting = f"Dear {co} Team,"

    intro = (
        f"Thank you for the productive meeting held on {mtg}, "
        f"chaired by {chair}. "
        "We value the continued partnership and wanted to share a brief summary of the key outcomes."
    )

    disc_block = "KEY DISCUSSION POINTS\n" + ("-" * 40)
    for ln in _parse_lines(disc):
        disc_block += f"\n  • {ln}"

    actions = s["rb_actions"]
    act_block = "\nAGREED ACTION ITEMS\n" + ("-" * 40)
    has_actions = False
    for i, (_, row) in enumerate(actions.iterrows(), 1):
        en_action = row.get("Action (EN)") or row.get("Action (AR)", "")
        if not en_action:
            continue
        has_actions = True
        due = str(row.get("Due Date", "")) if row.get("Due Date") else "TBD"
        owner = row.get("Assigned To", "TBD")
        prio  = row.get("Priority", "")
        act_block += f"\n  {i}. {en_action}"
        act_block += f"\n     Owner: {owner}  |  Due: {due}  |  Priority: {prio}"
    if not has_actions:
        act_block += "\n  (No action items recorded)"

    next_block = f"\nNEXT STEPS\nOur next meeting is tentatively scheduled for {next_m}. We look forward to continued progress."

    closing = (
        "\nPlease do not hesitate to reach out should you have any questions or require further clarification.\n\n"
        "Best regards,\n\n"
        f"{chair}\n"
        "Ministry of Investment of Saudi Arabia\n"
        "وزارة الاستثمار — المملكة العربية السعودية"
    )

    return "\n\n".join([subject_line, greeting, intro, disc_block, act_block, next_block, closing])


# ── Tab 3: Sync to CRM ────────────────────────────────────────────────────────

def _sync_tab(dfs: dict, lang: str):
    st.markdown("#### Sync Action Items to CRM Tracker")

    s        = st.session_state
    actions  = s["rb_actions"]
    company  = s["rb_company"]
    mtg_date = s["rb_date"]

    valid = actions[actions["Action (EN)"].notna() & (actions["Action (EN)"] != "") |
                    actions["Action (AR)"].notna() & (actions["Action (AR)"] != "")]

    if valid.empty:
        st.info("No action items to sync yet. Add them in Step 3 above.")
        return

    st.markdown(f"**{len(valid)} action item(s) will be added** to the CRM for **{company or '(no company selected)'}**.")

    # Preview
    preview_cols = ["Action (EN)", "Assigned To", "Type", "Priority", "Due Date", "Remarks"]
    st.dataframe(
        valid[[c for c in preview_cols if c in valid.columns]],
        use_container_width=True,
        hide_index=True,
    )

    if not company:
        st.warning("Please select a company in Step 1 before syncing.")
        return

    if st.button("🔄 Sync to Action Items sheet", type="primary", use_container_width=False, key="_rb_sync_btn"):
        investors = dfs.get("Investor Master", pd.DataFrame())
        inv_id    = _get_investor_id(investors, company)
        existing  = dfs.get("Action Items", pd.DataFrame())
        new_rows  = []

        for _, row in valid.iterrows():
            en_action = row.get("Action (EN)") or row.get("Action (AR)", "")
            new_id    = _next_act_id(existing, len(new_rows))
            new_row   = {
                "Action ID":          new_id,
                "Investor ID":        inv_id,
                "Company Name":       company,
                "Meeting ID":         "",
                "Opportunity ID":     "",
                "Action Description": en_action,
                "Assigned To":        row.get("Assigned To", ""),
                "Department":         "",
                "Sector":             "",
                "Type of Engagement": row.get("Type", "Action"),
                "Start Date":         mtg_date,
                "Due Date":           row.get("Due Date") if pd.notna(row.get("Due Date")) else None,
                "Priority":           _map_priority(row.get("Priority", "Medium")),
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
            new_rows.append(new_row)
            existing = pd.concat([existing, pd.DataFrame([new_row])], ignore_index=True)

        dfs["Action Items"] = existing
        save_session(dfs)
        st.success(f"✅ {len(new_rows)} action item(s) added to the CRM tracker for {company}.")
        st.balloons()


# ── Helpers ───────────────────────────────────────────────────────────────────

def _parse_lines(text: str) -> list:
    if not text:
        return []
    return [ln.strip().lstrip("•-– ") for ln in text.split("\n") if ln.strip()]


def _map_priority(p: str) -> str:
    mapping = {"Very High": "High", "High": "High", "Medium": "Medium", "Low": "Low"}
    return mapping.get(p, "Medium")


def _get_investor_id(investors: pd.DataFrame, company: str) -> str:
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
