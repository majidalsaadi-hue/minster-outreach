
# Opportunity Pipeline — view and add investment opportunities.

import pandas as pd
import plotly.express as px
import streamlit as st
from datetime import date

from config.settings import (
    SECTORS, OPPORTUNITY_TYPES, OPPORTUNITY_SOURCES, OPPORTUNITY_STAGES,
    CONFIDENCE_LEVELS, ESCALATION_FLAGS, STATUS_COLORS, MISA_GREEN,
)
from config.translations import t

OPP_STATUSES = ["Active", "Under Review", "Blocked", "Converted to Deal", "Dropped"]


def render(dfs: dict, lang: str):
    opps      = dfs.get("Opportunity Pipeline", pd.DataFrame())
    investors = dfs.get("Investor Master",      pd.DataFrame())

    st.markdown(f"### {t('nav_opportunities', lang)}")

    # ── Add opportunity ────────────────────────────────────────────────────────
    with st.expander(f"➕ {t('add_opportunity', lang)}", expanded=False):
        _add_opportunity_form(dfs, investors, lang)

    # ── Filters ───────────────────────────────────────────────────────────────
    with st.expander(t("filter", lang), expanded=False):
        fc1, fc2, fc3, fc4 = st.columns(4)
        companies   = sorted(opps["Company Name"].dropna().unique().tolist()) if not opps.empty and "Company Name" in opps.columns else []
        filter_co   = fc1.multiselect(t("company_name", lang), companies)
        filter_sect = fc2.multiselect(t("sector", lang), SECTORS)
        filter_stg  = fc3.multiselect(t("opportunity_stage", lang), OPPORTUNITY_STAGES)
        filter_stat = fc4.multiselect(t("opportunity_status", lang), OPP_STATUSES)

    filtered = opps.copy()
    if not filtered.empty:
        if filter_co   and "Company Name"        in filtered.columns:
            filtered = filtered[filtered["Company Name"].isin(filter_co)]
        if filter_sect and "Sector"              in filtered.columns:
            filtered = filtered[filtered["Sector"].isin(filter_sect)]
        if filter_stg  and "Opportunity Stage"   in filtered.columns:
            filtered = filtered[filtered["Opportunity Stage"].isin(filter_stg)]
        if filter_stat and "Opportunity Status"  in filtered.columns:
            filtered = filtered[filtered["Opportunity Status"].isin(filter_stat)]

    # ── Mini charts ────────────────────────────────────────────────────────────
    if not filtered.empty:
        _render_opp_charts(filtered, lang)

    # ── Table ─────────────────────────────────────────────────────────────────
    if filtered.empty:
        st.info("No opportunities found.")
        return

    display_cols = [
        c for c in [
            "Opportunity ID", "Company Name", "Opportunity Name",
            "Opportunity Type", "Sector", "Opportunity Stage",
            "Confidence Level", "Est. Value (SAR)",
            "Opportunity Status", "Target Closure Date",
            "Escalation Required", "Assigned AM",
        ] if c in filtered.columns
    ]

    st.dataframe(
        filtered[display_cols],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Est. Value (SAR)":   st.column_config.NumberColumn(format="SAR %,.0f"),
            "Target Closure Date":st.column_config.DateColumn(),
            "Opportunity Status": st.column_config.SelectboxColumn(options=OPP_STATUSES),
            "Opportunity Stage":  st.column_config.SelectboxColumn(options=OPPORTUNITY_STAGES),
            "Confidence Level":   st.column_config.SelectboxColumn(options=CONFIDENCE_LEVELS),
            "Escalation Required":st.column_config.SelectboxColumn(options=["Yes", "No"]),
        },
    )
    st.caption(f"{len(filtered)} {t('opportunities_count', lang)}")


def _render_opp_charts(df: pd.DataFrame, lang: str):
    col1, col2, col3 = st.columns(3)

    with col1:
        if "Opportunity Stage" in df.columns:
            sc = df["Opportunity Stage"].value_counts().reset_index()
            sc.columns = ["Stage", "Count"]
            fig = px.bar(sc, x="Stage", y="Count", title=t("chart_by_stage", lang),
                         color_discrete_sequence=[MISA_GREEN])
            fig.update_layout(margin=dict(t=30, b=10, l=10, r=10),
                               paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        if "Confidence Level" in df.columns:
            cc = df["Confidence Level"].value_counts().reset_index()
            cc.columns = ["Confidence", "Count"]
            colors = {"High": MISA_GREEN, "Medium": "#C9974A", "Low": "#C0392B", "Speculative": "#9B9B9B"}
            fig2 = px.pie(cc, values="Count", names="Confidence",
                          title=t("confidence", lang),
                          color_discrete_map=colors)
            fig2.update_layout(margin=dict(t=30, b=10, l=10, r=10),
                                paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig2, use_container_width=True)

    with col3:
        if "Est. Value (SAR)" in df.columns and "Sector" in df.columns:
            sv = df.groupby("Sector")["Est. Value (SAR)"].sum().reset_index()
            sv = sv[sv["Est. Value (SAR)"] > 0]
            if not sv.empty:
                fig3 = px.bar(sv, x="Sector", y="Est. Value (SAR)",
                              title=t("value_sar", lang),
                              color_discrete_sequence=["#C9974A"])
                fig3.update_layout(margin=dict(t=30, b=10, l=10, r=10),
                                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig3, use_container_width=True)


def _add_opportunity_form(dfs: dict, investors: pd.DataFrame, lang: str):
    company_options = [""] + (
        sorted(investors["Company Name"].dropna().unique().tolist())
        if not investors.empty and "Company Name" in investors.columns else []
    )

    with st.form("add_opp_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        company  = c1.selectbox(t("company_name", lang), company_options)
        opp_name = c2.text_input(t("opportunity_name", lang))
        c3, c4 = st.columns(2)
        opp_type = c3.selectbox(t("opportunity_type", lang), OPPORTUNITY_TYPES)
        source   = c4.selectbox(t("opportunity_source", lang), OPPORTUNITY_SOURCES)
        c5, c6 = st.columns(2)
        sector   = c5.selectbox(t("sector", lang), SECTORS)
        stage    = c6.selectbox(t("opportunity_stage", lang), OPPORTUNITY_STAGES)
        c7, c8 = st.columns(2)
        conf     = c7.selectbox(t("confidence", lang), CONFIDENCE_LEVELS)
        status   = c8.selectbox(t("opportunity_status", lang), OPP_STATUSES)
        c9, c10 = st.columns(2)
        est_val  = c9.number_input(t("est_value", lang), min_value=0.0, step=1_000_000.0)
        target_d = c10.date_input(t("target_closure", lang), value=None)
        c11, c12 = st.columns(2)
        assigned = c11.text_input(t("am", lang))
        escalate = c12.selectbox(t("escalation_flag", lang), ["No", "Yes"])
        blockers = st.text_area(t("blockers_identified", lang))
        notes    = st.text_area(t("notes", lang))

        if st.form_submit_button(t("add_opportunity", lang)):
            if not company or not opp_name:
                st.warning("Company and opportunity name are required.")
                return
            opps   = dfs.get("Opportunity Pipeline", pd.DataFrame())
            inv_id = _get_investor_id(investors, company)
            new_id = _next_id(opps, "Opportunity ID", "OPP")

            new_row = {
                "Opportunity ID":    new_id,
                "Investor ID":       inv_id,
                "Company Name":      company,
                "Opportunity Name":  opp_name,
                "Opportunity Type":  opp_type,
                "Opportunity Source":source,
                "Sector":            sector,
                "Opportunity Stage": stage,
                "Confidence Level":  conf,
                "Opportunity Status":status,
                "Est. Value (SAR)":  est_val if est_val > 0 else None,
                "Assigned AM":       assigned,
                "Start Date":        date.today(),
                "Target Closure Date":target_d,
                "Blockers":          blockers,
                "Escalation Required": escalate,
                "Last Updated":      date.today(),
                "Notes":             notes,
            }
            dfs["Opportunity Pipeline"] = pd.concat(
                [opps, pd.DataFrame([new_row])], ignore_index=True
            )
            st.success(f"✅ Opportunity {new_id} added for {company}")
            st.rerun()


def _get_investor_id(investors: pd.DataFrame, company: str) -> str:
    if investors.empty or "Company Name" not in investors.columns:
        return ""
    matches = investors[investors["Company Name"] == company]
    return str(matches.iloc[0].get("Investor ID", "")) if not matches.empty else ""


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
