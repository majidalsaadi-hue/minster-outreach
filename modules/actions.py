
# Action Items — consolidated view across all investors.

import pandas as pd
import streamlit as st
from datetime import date

from config.settings import (
    SECTORS, ENGAGEMENT_TYPES, ACTION_STATUSES, PROGRESS_OPTIONS,
    ESCALATION_FLAGS, TASK_PRIORITIES, DEPARTMENTS, MISA_GREEN,
    STATUS_COLORS, PRIORITY_COLORS,
)
from config.translations import t
from modules.persistence import save_session


def render(dfs: dict, lang: str):
    actions   = dfs.get("Action Items",    pd.DataFrame())
    investors = dfs.get("Investor Master", pd.DataFrame())

    st.markdown(f"### {t('nav_actions', lang)}")

    # ── Add action ─────────────────────────────────────────────────────────────
    with st.expander(f"➕ {t('add_action', lang)}", expanded=False):
        _add_action_form(dfs, investors, lang)

    # ── Filters ───────────────────────────────────────────────────────────────
    with st.expander(t("filter", lang), expanded=False):
        fc1, fc2, fc3, fc4, fc5 = st.columns(5)
        companies    = sorted(actions["Company Name"].dropna().unique().tolist()) if not actions.empty and "Company Name" in actions.columns else []
        filter_co    = fc1.multiselect(t("company_name", lang), companies)
        filter_stat  = fc2.multiselect(t("status", lang), ACTION_STATUSES)
        filter_pri   = fc3.multiselect(t("priority", lang), ["High", "Medium", "Low"])
        filter_type  = fc4.multiselect(t("engagement_type", lang), ENGAGEMENT_TYPES)
        filter_escl  = fc5.selectbox("Escalation Only", ["All", "Flag for RM", "Flag for Leadership"])

    filtered = actions.copy()
    if not filtered.empty:
        if filter_co   and "Company Name"       in filtered.columns:
            filtered = filtered[filtered["Company Name"].isin(filter_co)]
        if filter_stat and "Status"             in filtered.columns:
            filtered = filtered[filtered["Status"].isin(filter_stat)]
        if filter_pri  and "Priority"           in filtered.columns:
            filtered = filtered[filtered["Priority"].isin(filter_pri)]
        if filter_type and "Type of Engagement" in filtered.columns:
            filtered = filtered[filtered["Type of Engagement"].isin(filter_type)]
        if filter_escl != "All" and "Escalation Flag" in filtered.columns:
            filtered = filtered[filtered["Escalation Flag"] == filter_escl]

    # ── Summary strip ─────────────────────────────────────────────────────────
    if not filtered.empty and "Status" in filtered.columns:
        _render_status_summary(filtered, lang)

    # ── Table ─────────────────────────────────────────────────────────────────
    if filtered.empty:
        st.info("No action items found.")
        return

    today = date.today()
    if "Due Date" in filtered.columns and "Status" in filtered.columns:
        due = pd.to_datetime(filtered["Due Date"], errors="coerce")
        statuses = filtered["Status"].values
        filtered = filtered.copy()
        overdue_flags = []
        for i, d in enumerate(due):
            if pd.notna(d) and d.date() < today and statuses[i] not in ("Completed", "Cancelled"):
                overdue_flags.append("🔴")
            else:
                overdue_flags.append("")
        filtered["⚠️"] = overdue_flags

    display_cols = [
        c for c in [
            "⚠️", "Action ID", "Company Name", "Action Description",
            "Type of Engagement", "Sector", "Assigned To",
            "Priority", "Status", "Progress", "Due Date",
            "Escalation Flag", "Next Action Date",
        ] if c in filtered.columns
    ]

    st.dataframe(
        filtered[display_cols].sort_values("Due Date", ascending=True)
        if "Due Date" in filtered.columns else filtered[display_cols],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Due Date":           st.column_config.DateColumn(),
            "Next Action Date":   st.column_config.DateColumn(),
            "Status":             st.column_config.SelectboxColumn(options=ACTION_STATUSES),
            "Priority":           st.column_config.SelectboxColumn(options=["High", "Medium", "Low"]),
            "Progress":           st.column_config.SelectboxColumn(options=PROGRESS_OPTIONS),
            "Escalation Flag":    st.column_config.SelectboxColumn(options=ESCALATION_FLAGS),
            "Type of Engagement": st.column_config.SelectboxColumn(options=ENGAGEMENT_TYPES),
        },
    )
    st.caption(f"{len(filtered)} {t('actions_count', lang)}")


def _render_status_summary(df: pd.DataFrame, lang: str):
    counts = df["Status"].value_counts()
    cols   = st.columns(len(ACTION_STATUSES))
    for i, status in enumerate(ACTION_STATUSES):
        count = counts.get(status, 0)
        color = STATUS_COLORS.get(status, "#9B9B9B")
        cols[i].markdown(
            f"<div style='background:{color};color:#fff;padding:8px 4px;"
            f"border-radius:6px;text-align:center;font-size:12px;'>"
            f"<b>{count}</b><br>{status}</div>",
            unsafe_allow_html=True,
        )
    st.markdown("")


def _add_action_form(dfs: dict, investors: pd.DataFrame, lang: str):
    company_options = [""] + (
        sorted(investors["Company Name"].dropna().unique().tolist())
        if not investors.empty and "Company Name" in investors.columns else []
    )

    with st.form("add_action_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        company = c1.selectbox(t("company_name", lang), company_options)
        action  = c2.text_input(t("action_description", lang))
        c3, c4 = st.columns(2)
        eng_type = c3.selectbox(t("engagement_type", lang), ENGAGEMENT_TYPES)
        sector   = c4.selectbox(t("sector", lang), [""] + SECTORS)
        c5, c6 = st.columns(2)
        assigned  = c5.text_input(t("assigned_to", lang))
        dept      = c6.selectbox(t("department", lang), [""] + DEPARTMENTS)
        c7, c8 = st.columns(2)
        priority  = c7.selectbox(t("priority", lang), ["High", "Medium", "Low"])
        status    = c8.selectbox(t("status", lang), ACTION_STATUSES)
        c9, c10 = st.columns(2)
        start_d   = c9.date_input(t("start_date", lang), value=date.today())
        due_d     = c10.date_input(t("due_date", lang), value=None)
        c11, c12 = st.columns(2)
        escalation = c11.selectbox(t("escalation_flag", lang), ESCALATION_FLAGS)
        progress   = c12.selectbox(t("progress", lang), PROGRESS_OPTIONS)
        remarks    = st.text_area(t("remarks", lang))
        next_action = st.text_input(t("next_action", lang))
        na_date    = st.date_input(t("next_action_date", lang), value=None)

        if st.form_submit_button(t("add_action", lang)):
            if not company or not action:
                st.warning("Company and action description are required.")
                return

            actions = dfs.get("Action Items", pd.DataFrame())
            inv_id  = _get_investor_id(investors, company)
            new_id  = _next_id(actions, "Action ID", "ACT")

            new_row = {
                "Action ID":          new_id,
                "Investor ID":        inv_id,
                "Company Name":       company,
                "Meeting ID":         "",
                "Opportunity ID":     "",
                "Action Description": action,
                "Assigned To":        assigned,
                "Department":         dept,
                "Sector":             sector,
                "Type of Engagement": eng_type,
                "Start Date":         start_d,
                "Due Date":           due_d,
                "Priority":           priority,
                "Progress":           progress,
                "Status":             status,
                "Escalation Flag":    escalation,
                "Remarks":            remarks,
                "Outcome":            "",
                "Next Action":        next_action,
                "Next Action Date":   na_date,
                "Last Updated":       date.today(),
                "Updated By":         "",
            }
            dfs["Action Items"] = pd.concat(
                [actions, pd.DataFrame([new_row])], ignore_index=True
            )
            save_session(dfs)
            st.success(f"✅ Action {new_id} added for {company}")
            st.rerun()


def _get_investor_id(investors: pd.DataFrame, company: str) -> str:
    if investors.empty or "Company Name" not in investors.columns:
        return ""
    m = investors[investors["Company Name"] == company]
    return str(m.iloc[0].get("Investor ID", "")) if not m.empty else ""


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
