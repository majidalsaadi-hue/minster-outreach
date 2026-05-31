
# Meeting Log — view and add meeting records.

import pandas as pd
import streamlit as st
from datetime import date

from config.settings import (
    MEETING_TYPES, MEETING_OBJECTIVES, MEETING_STATUSES, MISA_GREEN,
)
from config.translations import t
from modules.persistence import save_session


def render(dfs: dict, lang: str):
    meetings  = dfs.get("Meeting Log",    pd.DataFrame())
    investors = dfs.get("Investor Master", pd.DataFrame())

    st.markdown(f"### {t('nav_meetings', lang)}")

    # ── Add meeting ───────────────────────────────────────────────────────────
    with st.expander(f"➕ {t('add_meeting', lang)}", expanded=False):
        _add_meeting_form(dfs, investors, lang)

    # ── Filters ───────────────────────────────────────────────────────────────
    with st.expander(t("filter", lang), expanded=False):
        fc1, fc2, fc3 = st.columns(3)
        companies   = sorted(meetings["Company Name"].dropna().unique().tolist()) if not meetings.empty and "Company Name" in meetings.columns else []
        filter_co   = fc1.multiselect(t("company_name", lang), companies)
        filter_stat = fc2.multiselect(t("meeting_status", lang), MEETING_STATUSES)
        filter_type = fc3.multiselect(t("meeting_type", lang), MEETING_TYPES)

    filtered = meetings.copy()
    if not filtered.empty:
        if filter_co   and "Company Name"   in filtered.columns:
            filtered = filtered[filtered["Company Name"].isin(filter_co)]
        if filter_stat and "Meeting Status" in filtered.columns:
            filtered = filtered[filtered["Meeting Status"].isin(filter_stat)]
        if filter_type and "Meeting Type"   in filtered.columns:
            filtered = filtered[filtered["Meeting Type"].isin(filter_type)]

    if filtered.empty:
        st.info("No meetings found. Use 'Log Meeting' above to add one.")
        return

    display_cols = [
        c for c in [
            "Meeting ID", "Company Name", "Meeting Date", "Meeting Type",
            "Meeting Objective", "Meeting Status", "Follow-Up Owner",
            "Follow-Up Due Date", "RM Reviewed",
        ] if c in filtered.columns
    ]

    st.dataframe(
        filtered[display_cols].sort_values("Meeting Date", ascending=False)
        if "Meeting Date" in filtered.columns else filtered[display_cols],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Meeting Date":      st.column_config.DateColumn(),
            "Follow-Up Due Date":st.column_config.DateColumn(),
            "Meeting Status":    st.column_config.SelectboxColumn(options=MEETING_STATUSES),
            "RM Reviewed":       st.column_config.SelectboxColumn(options=["Yes", "No"]),
        },
    )
    st.caption(f"{len(filtered)} {t('meetings_count', lang)}")

    # ── Detail view ───────────────────────────────────────────────────────────
    if not filtered.empty and "Meeting ID" in filtered.columns:
        selected_id = st.selectbox(
            "View meeting detail",
            ["— select —"] + filtered["Meeting ID"].dropna().tolist(),
        )
        if selected_id != "— select —":
            row = filtered[filtered["Meeting ID"] == selected_id].iloc[0]
            _render_meeting_detail(row, lang)


def _add_meeting_form(dfs: dict, investors: pd.DataFrame, lang: str):
    company_options = [""] + (
        sorted(investors["Company Name"].dropna().unique().tolist())
        if not investors.empty and "Company Name" in investors.columns else []
    )

    with st.form("add_meeting_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        company   = c1.selectbox(t("company_name", lang), company_options)
        mtg_date  = c2.date_input(t("meeting_date", lang), value=date.today())
        c3, c4 = st.columns(2)
        mtg_type  = c3.selectbox(t("meeting_type", lang), MEETING_TYPES)
        objective = c4.selectbox(t("meeting_objective", lang), MEETING_OBJECTIVES)
        c5, c6 = st.columns(2)
        misa_att  = c5.text_input(t("misa_attendees", lang))
        inv_att   = c6.text_input(t("investor_attendees", lang))
        location  = st.text_input(t("location", lang))
        discussion = st.text_area(t("key_discussion", lang))
        decisions  = st.text_area(t("decisions_made", lang))
        blockers   = st.text_area(t("blockers_identified", lang))
        next_steps = st.text_area(t("next_steps", lang))
        c7, c8 = st.columns(2)
        fu_owner  = c7.text_input(t("follow_up_owner", lang))
        fu_date   = c8.date_input(t("follow_up_due", lang), value=None)
        c9, c10 = st.columns(2)
        status    = c9.selectbox(t("meeting_status", lang), MEETING_STATUSES)
        reviewed  = c10.selectbox(t("rm_reviewed", lang), ["No", "Yes"])
        logged_by = st.text_input(t("logged_by", lang))

        if st.form_submit_button(t("add_meeting", lang)):
            if not company:
                st.warning("Please select an investor.")
                return

            meetings = dfs.get("Meeting Log", pd.DataFrame())
            inv_id   = _get_investor_id(investors, company)
            new_id   = _next_id(meetings, "Meeting ID", "MTG")

            new_row = {
                "Meeting ID":           new_id,
                "Investor ID":          inv_id,
                "Company Name":         company,
                "Meeting Date":         mtg_date,
                "Meeting Type":         mtg_type,
                "Location":             location,
                "MISA Attendees":       misa_att,
                "Investor Attendees":   inv_att,
                "Meeting Objective":    objective,
                "Key Discussion Points":discussion,
                "Decisions Made":       decisions,
                "Blockers Identified":  blockers,
                "Next Steps":           next_steps,
                "Follow-Up Owner":      fu_owner,
                "Follow-Up Due Date":   fu_date,
                "Meeting Status":       status,
                "RM Reviewed":          reviewed,
                "Logged By":            logged_by,
            }
            dfs["Meeting Log"] = pd.concat(
                [meetings, pd.DataFrame([new_row])], ignore_index=True
            )
            # Update last meeting date on investor record
            _update_last_meeting(dfs, company, mtg_date)
            save_session(dfs)
            st.success(f"✅ Meeting {new_id} logged for {company}")
            st.rerun()


def _render_meeting_detail(row, lang: str):
    st.markdown("---")
    st.markdown(f"#### {row.get('Company Name','?')} — {row.get('Meeting Date','')}")
    cols = [
        (t("meeting_type",      lang), "Meeting Type"),
        (t("meeting_objective", lang), "Meeting Objective"),
        (t("location",          lang), "Location"),
        (t("misa_attendees",    lang), "MISA Attendees"),
        (t("investor_attendees",lang), "Investor Attendees"),
        (t("meeting_status",    lang), "Meeting Status"),
        (t("rm_reviewed",       lang), "RM Reviewed"),
    ]
    c1, c2 = st.columns(2)
    for i, (label, key) in enumerate(cols):
        (c1 if i % 2 == 0 else c2).markdown(f"**{label}:** {row.get(key,'—')}")

    for label, key in [
        (t("key_discussion", lang),     "Key Discussion Points"),
        (t("decisions_made", lang),     "Decisions Made"),
        (t("blockers_identified", lang),"Blockers Identified"),
        (t("next_steps", lang),         "Next Steps"),
    ]:
        val = row.get(key, "")
        if val:
            st.markdown(f"**{label}:** {val}")


def _get_investor_id(investors: pd.DataFrame, company: str) -> str:
    if investors.empty or "Company Name" not in investors.columns:
        return ""
    matches = investors[investors["Company Name"] == company]
    if matches.empty:
        return ""
    return str(matches.iloc[0].get("Investor ID", ""))


def _update_last_meeting(dfs: dict, company: str, mtg_date):
    investors = dfs.get("Investor Master", pd.DataFrame())
    if investors.empty or "Company Name" not in investors.columns:
        return
    mask = investors["Company Name"] == company
    investors.loc[mask, "Last Meeting Date"] = mtg_date
    dfs["Investor Master"] = investors


def _next_id(df: pd.DataFrame, id_col: str, prefix: str) -> str:
    if df.empty or id_col not in df.columns:
        return f"{prefix}-001"
    nums = []
    for v in df[id_col].dropna():
        try:
            nums.append(int(str(v).split("-")[-1]))
        except ValueError:
            pass
    return f"{prefix}-{(max(nums)+1 if nums else 1):03d}"
