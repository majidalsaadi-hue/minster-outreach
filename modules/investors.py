
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
from modules.persistence import save_session


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
            save_session(dfs)
            st.success(f"✅ {company} added ({new_id})")
            st.rerun()


def _render_investor_profile(company: str, dfs: dict, lang: str):
    investors    = dfs.get("Investor Master",     pd.DataFrame())
    actions      = dfs.get("Action Items",        pd.DataFrame())
    meetings     = dfs.get("Meeting Log",         pd.DataFrame())
    opportunities= dfs.get("Opportunity Pipeline",pd.DataFrame())
    tasks        = dfs.get("RM Tasks",            pd.DataFrame())
    deals        = dfs.get("Deal Progress",       pd.DataFrame())

    row = investors[investors["Company Name"] == company].iloc[0]

    # ── Company header banner ─────────────────────────────────────────────────
    tier       = row.get("Investor Tier", "—")
    tier_color = TIER_COLORS.get(tier, MISA_GREEN)
    status     = row.get("Relationship Status", "—")
    stage      = row.get("Journey Stage", "—")

    st.markdown(f"""
    <div style="background:{MISA_GREEN};border-radius:10px;padding:16px 20px;margin:12px 0;">
      <div style="display:flex;justify-content:space-between;align-items:center;">
        <div>
          <div style="color:white;font-size:22px;font-weight:700;">{company}</div>
          <div style="color:rgba(255,255,255,0.75);font-size:13px;margin-top:2px;">
            {row.get('Country','—')} &nbsp;|&nbsp; {row.get('Sector','—')} &nbsp;|&nbsp; RM: {row.get('Relationship Manager','—')} &nbsp;|&nbsp; AM: {row.get('Account Manager','TBD')}
          </div>
        </div>
        <div style="text-align:right;">
          <div style="background:{tier_color};color:white;padding:4px 12px;border-radius:20px;font-size:12px;font-weight:700;">{tier}</div>
          <div style="color:rgba(255,255,255,0.75);font-size:12px;margin-top:4px;">{stage}</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Key metrics strip ─────────────────────────────────────────────────────
    est  = row.get("Est. Investment Value (SAR)")
    cmmt = row.get("Actual Commitment (SAR)")
    jobs = row.get("Est. Jobs Created")
    priority = row.get("Strategic Priority Score", "—")
    min_action = row.get("Minister Action Required", "None Required")

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Est. Investment", f"SAR {est:,.0f}" if pd.notna(est) and est else "—")
    m2.metric("Commitment",      f"SAR {cmmt:,.0f}" if pd.notna(cmmt) and cmmt else "—")
    m3.metric("Est. Jobs",       f"{int(jobs):,}" if pd.notna(jobs) and jobs else "—")
    m4.metric("Priority Score",  str(priority) if str(priority) not in ("—", "nan", "") else "—")
    m5.metric("Minister Action", str(min_action) if str(min_action) not in ("None Required", "nan", "") else "None")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    # ── Filter linked data ────────────────────────────────────────────────────
    linked_meetings = meetings[meetings["Company Name"] == company] if not meetings.empty and "Company Name" in meetings.columns else pd.DataFrame()
    linked_opps     = opportunities[opportunities["Company Name"] == company] if not opportunities.empty and "Company Name" in opportunities.columns else pd.DataFrame()
    linked_actions  = actions[actions["Company Name"] == company] if not actions.empty and "Company Name" in actions.columns else pd.DataFrame()
    linked_tasks    = tasks[tasks["Linked Investor"] == company] if not tasks.empty and "Linked Investor" in tasks.columns else pd.DataFrame()
    linked_deals    = deals[deals["Company Name"] == company] if not deals.empty and "Company Name" in deals.columns else pd.DataFrame()

    # Separate challenges from action items
    if not linked_actions.empty and "Type of Engagement" in linked_actions.columns:
        linked_challenges = linked_actions[linked_actions["Type of Engagement"] == "Challenge"]
        linked_actions_only = linked_actions[linked_actions["Type of Engagement"] != "Challenge"]
    else:
        linked_challenges   = pd.DataFrame()
        linked_actions_only = linked_actions

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tab_meetings, tab_opps, tab_challenges, tab_actions, tab_tasks, tab_deals = st.tabs([
        f"🤝 Meetings ({len(linked_meetings)})",
        f"🎯 Opportunities ({len(linked_opps)})",
        f"⚠️ Challenges ({len(linked_challenges)})",
        f"✅ Action Items ({len(linked_actions_only)})",
        f"📋 RM Tasks ({len(linked_tasks)})",
        f"🏦 Deals ({len(linked_deals)})",
    ])

    # ── Meetings tab ──────────────────────────────────────────────────────────
    with tab_meetings:
        if linked_meetings.empty:
            st.info("No meetings logged for this company yet.")
        else:
            cols = [c for c in [
                "Meeting Date", "Meeting Type", "Meeting Status",
                "Meeting Objective", "Key Discussion Points",
                "Decisions Made", "Blockers Identified",
                "Next Steps", "Follow-Up Owner", "Follow-Up Due Date",
            ] if c in linked_meetings.columns]
            st.dataframe(
                linked_meetings[cols].sort_values("Meeting Date", ascending=False)
                if "Meeting Date" in linked_meetings.columns else linked_meetings[cols],
                use_container_width=True, hide_index=True,
            )

    # ── Opportunities tab ─────────────────────────────────────────────────────
    with tab_opps:
        if linked_opps.empty:
            st.info("No opportunities linked to this company yet.")
        else:
            cols = [c for c in [
                "Opportunity Name", "Opportunity Stage", "Opportunity Status",
                "Est. Value (SAR)", "Confidence Level", "Target Closure Date",
                "Blockers", "Escalation Required", "Notes",
            ] if c in linked_opps.columns]
            st.dataframe(linked_opps[cols], use_container_width=True, hide_index=True)

    # ── Challenges tab ────────────────────────────────────────────────────────
    with tab_challenges:
        if linked_challenges.empty:
            st.info("No challenges logged for this company.")
        else:
            cols = [c for c in [
                "Action Description", "Status", "Priority",
                "Due Date", "Assigned To", "Escalation Flag", "Remarks",
            ] if c in linked_challenges.columns]
            st.dataframe(linked_challenges[cols], use_container_width=True, hide_index=True)

    # ── Action Items tab ──────────────────────────────────────────────────────
    with tab_actions:
        if linked_actions_only.empty:
            st.info("No action items for this company.")
        else:
            cols = [c for c in [
                "Action ID", "Action Description", "Type of Engagement",
                "Status", "Priority", "Progress",
                "Due Date", "Assigned To", "Department",
                "Escalation Flag", "Next Action", "Next Action Date",
            ] if c in linked_actions_only.columns]
            st.dataframe(
                linked_actions_only[cols].sort_values("Due Date", ascending=True)
                if "Due Date" in linked_actions_only.columns else linked_actions_only[cols],
                use_container_width=True, hide_index=True,
            )

    # ── RM Tasks tab ──────────────────────────────────────────────────────────
    with tab_tasks:
        if linked_tasks.empty:
            st.info("No RM tasks linked to this company.")
        else:
            cols = [c for c in [
                "Task ID", "Task Title", "Priority", "Status",
                "Due Date", "Notes",
            ] if c in linked_tasks.columns]
            st.dataframe(linked_tasks[cols], use_container_width=True, hide_index=True)

    # ── Deals tab ─────────────────────────────────────────────────────────────
    with tab_deals:
        if linked_deals.empty:
            st.info("No deals in progress for this company.")
        else:
            cols = [c for c in [
                "Deal ID", "Deal Name", "Deal Stage", "Deal Status",
                "Challenge Severity", "Challenge Classification",
                "Est. Value (SAR)", "Escalation Required", "Escalation Level",
                "Escalation Status", "Assigned Owner", "Target Resolution Date",
                "Last Updated",
            ] if c in linked_deals.columns]
            st.dataframe(linked_deals[cols], use_container_width=True, hide_index=True,
                         column_config={
                             "Est. Value (SAR)":       st.column_config.NumberColumn(format="SAR %,.0f"),
                             "Target Resolution Date": st.column_config.DateColumn(),
                             "Last Updated":           st.column_config.DateColumn(),
                         })
            # Per-deal challenge detail for this company
            for _, drow in linked_deals.iterrows():
                deal_id   = str(drow.get("Deal ID", "—"))
                deal_name = str(drow.get("Deal Name", "—"))
                severity  = str(drow.get("Challenge Severity", ""))
                status    = str(drow.get("Deal Status", ""))
                flag = "🔴 " if (severity == "Critical" or status == "Blocked") else ("🟠 " if severity == "High" else "")
                with st.expander(f"{flag}{deal_id} — {deal_name}", expanded=False):
                    d1, d2 = st.columns(2)
                    d1.markdown(f"**Stage:** {drow.get('Deal Stage', '—')}")
                    d1.markdown(f"**Status:** {status or '—'}")
                    d1.markdown(f"**Challenge Classification:** {drow.get('Challenge Classification', '—')}")
                    d2.markdown(f"**Escalation Level:** {drow.get('Escalation Level', '—')}")
                    d2.markdown(f"**Escalation Status:** {drow.get('Escalation Status', '—')}")
                    d2.markdown(f"**Assigned Owner:** {drow.get('Assigned Owner', '—')}")
                    challenge_desc = str(drow.get("Challenge Description", "") or "")
                    proposed_sol   = str(drow.get("Proposed Solution", "") or "")
                    if challenge_desc:
                        st.markdown(f"**Challenge:** {challenge_desc}")
                    if proposed_sol:
                        st.markdown(f"**Proposed Solution:** {proposed_sol}")

    # ── Notes ─────────────────────────────────────────────────────────────────
    notes = str(row.get("Notes", "") or "")
    if notes:
        st.markdown(f"**Notes:** {notes}")


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
