
# Investor Master List — view, add, edit investor records.

import pandas as pd
import streamlit as st
from datetime import date

from config.settings import (
    SECTORS, COUNTRIES, INVESTOR_TIERS, INVESTOR_STATUSES,
    JOURNEY_STAGES, ESCALATION_FLAGS, MISA_GREEN, MISA_GOLD,
    STATUS_COLORS, TIER_COLORS,
)
from config.translations import t


def render(dfs: dict, lang: str):
    investors = dfs.get("Investor Master", pd.DataFrame())

    st.markdown(f"### {t('nav_investors', lang)}")

    # ── Filters ───────────────────────────────────────────────────────────────
    with st.expander(t("filter", lang), expanded=False):
        fc1, fc2, fc3, fc4 = st.columns(4)
        filter_tier    = fc1.multiselect(t("investor_tier", lang),   INVESTOR_TIERS, default=[])
        filter_status  = fc2.multiselect(t("relationship_status", lang), INVESTOR_STATUSES, default=[])
        filter_sector  = fc3.multiselect(t("sector", lang),          SECTORS, default=[])
        filter_country = fc4.multiselect(t("country", lang),         COUNTRIES, default=[])

    filtered = investors.copy()
    if not filtered.empty:
        if filter_tier    and "Investor Tier"         in filtered.columns:
            filtered = filtered[filtered["Investor Tier"].isin(filter_tier)]
        if filter_status  and "Relationship Status"   in filtered.columns:
            filtered = filtered[filtered["Relationship Status"].isin(filter_status)]
        if filter_sector  and "Sector"                in filtered.columns:
            filtered = filtered[filtered["Sector"].isin(filter_sector)]
        if filter_country and "Country"               in filtered.columns:
            filtered = filtered[filtered["Country"].isin(filter_country)]

    # ── Add investor button ───────────────────────────────────────────────────
    with st.expander(f"➕ {t('add_investor', lang)}", expanded=False):
        _add_investor_form(dfs, lang)

    # ── Table ─────────────────────────────────────────────────────────────────
    if filtered.empty:
        st.info(t("no_data", lang))
        return

    display_cols = [
        c for c in [
            "Investor ID", "Company Name", "Country", "Sector",
            "Investor Tier", "Relationship Status", "Journey Stage",
            "Relationship Manager", "Account Manager",
            "Est. Investment Value (SAR)", "Next Meeting Date",
            "Escalation Flag",
        ] if c in filtered.columns
    ]

    styled = filtered[display_cols].copy()

    # Colour-code the status column via a caption trick
    st.dataframe(
        styled,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Est. Investment Value (SAR)": st.column_config.NumberColumn(
                format="SAR %,.0f"
            ),
            "Next Meeting Date": st.column_config.DateColumn(),
            "Escalation Flag": st.column_config.SelectboxColumn(
                options=ESCALATION_FLAGS
            ),
            "Investor Tier": st.column_config.SelectboxColumn(
                options=INVESTOR_TIERS
            ),
            "Relationship Status": st.column_config.SelectboxColumn(
                options=INVESTOR_STATUSES
            ),
            "Journey Stage": st.column_config.SelectboxColumn(
                options=JOURNEY_STAGES
            ),
        },
    )

    st.caption(f"{len(filtered)} {t('investors_count', lang)}")

    # ── Investor detail drill-down ─────────────────────────────────────────────
    if not filtered.empty and "Company Name" in filtered.columns:
        selected_company = st.selectbox(
            t("investor_profile", lang),
            ["— select —"] + sorted(filtered["Company Name"].dropna().unique().tolist()),
        )
        if selected_company != "— select —":
            _render_investor_profile(selected_company, dfs, lang)


def _add_investor_form(dfs: dict, lang: str):
    with st.form("add_investor_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        company   = c1.text_input(t("company_name", lang))
        country   = c2.selectbox(t("country", lang), COUNTRIES)
        c3, c4 = st.columns(2)
        sector    = c3.selectbox(t("sector", lang), SECTORS)
        tier      = c4.selectbox(t("investor_tier", lang), INVESTOR_TIERS)
        c5, c6 = st.columns(2)
        rm        = c5.text_input(t("rm", lang))
        am        = c6.text_input(t("am", lang) + " (leave blank if TBD)")
        c7, c8 = st.columns(2)
        stage     = c7.selectbox(t("journey_stage", lang), JOURNEY_STAGES)
        status    = c8.selectbox(t("relationship_status", lang), INVESTOR_STATUSES)
        c9, c10 = st.columns(2)
        est_val   = c9.number_input(t("est_investment", lang), min_value=0.0, step=1_000_000.0)
        next_mtg  = c10.date_input(t("next_meeting_date", lang), value=None)
        notes     = st.text_area(t("notes", lang))

        if st.form_submit_button(t("add_new", lang)):
            investors = dfs.get("Investor Master", pd.DataFrame())
            new_id    = _next_id(investors, "Investor ID", "INV")
            new_row   = {
                "Investor ID":              new_id,
                "Company Name":             company,
                "Country":                  country,
                "Sector":                   sector,
                "Investor Tier":            tier,
                "Relationship Manager":     rm,
                "Account Manager":          am or "TBD",
                "Journey Stage":            stage,
                "Relationship Status":      status,
                "Est. Investment Value (SAR)": est_val if est_val > 0 else None,
                "Actual Commitment (SAR)":  None,
                "Last Meeting Date":        None,
                "Next Meeting Date":        next_mtg,
                "Last Updated":             date.today(),
                "Escalation Flag":          "None",
                "Notes":                    notes,
            }
            dfs["Investor Master"] = pd.concat(
                [investors, pd.DataFrame([new_row])], ignore_index=True
            )
            st.success(f"✅ {company} added ({new_id})")
            st.rerun()


def _render_investor_profile(company: str, dfs: dict, lang: str):
    investors = dfs.get("Investor Master", pd.DataFrame())
    row = investors[investors["Company Name"] == company].iloc[0]

    st.markdown(f"---\n#### {company}")
    c1, c2, c3 = st.columns(3)

    tier = row.get("Investor Tier", "—")
    tier_color = TIER_COLORS.get(tier, MISA_GREEN)
    c1.markdown(f"**{t('investor_tier', lang)}:** "
                f"<span style='color:{tier_color};font-weight:700'>{tier}</span>",
                unsafe_allow_html=True)
    c2.markdown(f"**{t('country', lang)}:** {row.get('Country','—')}")
    c3.markdown(f"**{t('sector', lang)}:** {row.get('Sector','—')}")

    c4, c5, c6 = st.columns(3)
    c4.markdown(f"**{t('rm', lang)}:** {row.get('Relationship Manager','—')}")
    c5.markdown(f"**{t('am', lang)}:** {row.get('Account Manager','TBD')}")
    c6.markdown(f"**{t('journey_stage', lang)}:** {row.get('Journey Stage','—')}")

    c7, c8, c9 = st.columns(3)
    est = row.get("Est. Investment Value (SAR)")
    cmmt = row.get("Actual Commitment (SAR)")
    c7.metric(t("est_investment", lang),  f"SAR {est:,.0f}"  if pd.notna(est)  and est  else "—")
    c8.metric(t("actual_commitment", lang), f"SAR {cmmt:,.0f}" if pd.notna(cmmt) and cmmt else "—")
    c9.markdown(f"**{t('escalation_flag', lang)}:** {row.get('Escalation Flag','—')}")

    # Linked actions
    actions = dfs.get("Action Items", pd.DataFrame())
    if not actions.empty and "Company Name" in actions.columns:
        linked = actions[actions["Company Name"] == company]
        if not linked.empty:
            st.markdown(f"**{t('nav_actions', lang)}** ({len(linked)})")
            st.dataframe(
                linked[["Action Description", "Status", "Priority", "Due Date", "Assigned To"]],
                use_container_width=True, hide_index=True,
            )

    # Linked meetings
    meetings = dfs.get("Meeting Log", pd.DataFrame())
    if not meetings.empty and "Company Name" in meetings.columns:
        linked_m = meetings[meetings["Company Name"] == company]
        if not linked_m.empty:
            st.markdown(f"**{t('nav_meetings', lang)}** ({len(linked_m)})")
            st.dataframe(
                linked_m[["Meeting Date", "Meeting Type", "Meeting Status", "Meeting Objective"]],
                use_container_width=True, hide_index=True,
            )


def _next_id(df: pd.DataFrame, id_col: str, prefix: str) -> str:
    if df.empty or id_col not in df.columns:
        return f"{prefix}-001"
    existing = df[id_col].dropna().tolist()
    nums = []
    for v in existing:
        try:
            nums.append(int(str(v).split("-")[-1]))
        except ValueError:
            pass
    next_num = max(nums) + 1 if nums else 1
    return f"{prefix}-{next_num:03d}"
