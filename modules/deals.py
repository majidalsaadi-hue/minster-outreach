
# Deal Progress — track active deals, challenges, and escalations.

import pandas as pd
import plotly.express as px
import streamlit as st
from datetime import date

from config.settings import (
    MISA_GREEN, MISA_GOLD, STATUS_COLORS,
    DEAL_STAGES, DEAL_STATUSES, CHALLENGE_CLASSIFICATIONS,
    CHALLENGE_SEVERITIES, ESCALATION_LEVELS, ESCALATION_STATUSES,
)
from config.translations import t
from modules.persistence import save_session


def render(dfs: dict, lang: str):
    deals     = dfs.get("Deal Progress", pd.DataFrame())
    investors = dfs.get("Investor Master", pd.DataFrame())

    st.markdown("### 🤝 Deal Progress")

    # ── KPI strip ─────────────────────────────────────────────────────────────
    _render_kpi_strip(deals)

    # ── Add deal ──────────────────────────────────────────────────────────────
    with st.expander("➕ Add Deal", expanded=False):
        _add_deal_form(dfs, investors, lang)

    # ── Filters ───────────────────────────────────────────────────────────────
    with st.expander(t("filter", lang), expanded=False):
        fc1, fc2, fc3, fc4 = st.columns(4)
        companies    = sorted(deals["Company Name"].dropna().unique().tolist()) if not deals.empty and "Company Name" in deals.columns else []
        filter_co    = fc1.multiselect(t("company_name", lang), companies)
        filter_stage = fc2.multiselect("Deal Stage", DEAL_STAGES)
        filter_stat  = fc3.multiselect("Deal Status", DEAL_STATUSES)
        filter_sev   = fc4.multiselect("Challenge Severity", CHALLENGE_SEVERITIES)

    filtered = deals.copy()
    if not filtered.empty:
        if filter_co    and "Company Name"         in filtered.columns:
            filtered = filtered[filtered["Company Name"].isin(filter_co)]
        if filter_stage and "Deal Stage"           in filtered.columns:
            filtered = filtered[filtered["Deal Stage"].isin(filter_stage)]
        if filter_stat  and "Deal Status"          in filtered.columns:
            filtered = filtered[filtered["Deal Status"].isin(filter_stat)]
        if filter_sev   and "Challenge Severity"   in filtered.columns:
            filtered = filtered[filtered["Challenge Severity"].isin(filter_sev)]

    # ── Charts ────────────────────────────────────────────────────────────────
    if not filtered.empty:
        _render_deal_charts(filtered, lang)

    # ── Table ─────────────────────────────────────────────────────────────────
    if filtered.empty:
        st.info("No deals found. Add a deal above.")
        return

    display_cols = [
        c for c in [
            "Deal ID", "Company Name", "Deal Name", "Deal Stage",
            "Deal Status", "Challenge Severity", "Est. Value (SAR)",
            "Escalation Required", "Escalation Level", "Escalation Status",
            "Assigned Owner", "Target Resolution Date", "Last Updated",
        ] if c in filtered.columns
    ]

    # Colour-code rows: red background indicator via caption
    def _row_label(row):
        sev    = str(row.get("Challenge Severity", ""))
        status = str(row.get("Deal Status", ""))
        if sev == "Critical" or status == "Blocked":
            return "🔴"
        if sev == "High":
            return "🟠"
        return ""

    display_df = filtered[display_cols].copy()
    if "Challenge Severity" in filtered.columns and "Deal Status" in filtered.columns:
        display_df.insert(0, "⚠️", filtered.apply(_row_label, axis=1))

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Est. Value (SAR)":        st.column_config.NumberColumn(format="SAR %,.0f"),
            "Target Resolution Date":  st.column_config.DateColumn(),
            "Last Updated":            st.column_config.DateColumn(),
            "Deal Stage":              st.column_config.SelectboxColumn(options=DEAL_STAGES),
            "Deal Status":             st.column_config.SelectboxColumn(options=DEAL_STATUSES),
            "Challenge Severity":      st.column_config.SelectboxColumn(options=CHALLENGE_SEVERITIES),
            "Escalation Required":     st.column_config.SelectboxColumn(options=["Yes", "No"]),
            "Escalation Level":        st.column_config.SelectboxColumn(options=ESCALATION_LEVELS),
            "Escalation Status":       st.column_config.SelectboxColumn(options=ESCALATION_STATUSES),
        },
    )
    st.caption(f"{len(filtered)} deal(s)")

    # ── Per-deal challenge detail expander ────────────────────────────────────
    st.markdown("---")
    st.markdown("#### Deal Challenge Details")
    _render_deal_expanders(filtered)


def _render_kpi_strip(deals: pd.DataFrame):
    total          = len(deals) if not deals.empty else 0
    blocked        = 0
    critical       = 0
    minister_level = 0

    if not deals.empty:
        if "Deal Status" in deals.columns:
            blocked = int((deals["Deal Status"] == "Blocked").sum())
        if "Challenge Severity" in deals.columns:
            critical = int((deals["Challenge Severity"] == "Critical").sum())
        if "Escalation Level" in deals.columns:
            minister_level = int((deals["Escalation Level"] == "Minister Level").sum())

    k1, k2, k3, k4 = st.columns(4)

    _kpi_card(k1, "Total Deals",               str(total),          MISA_GREEN)
    _kpi_card(k2, "Blocked Deals",             str(blocked),        "#C0392B" if blocked > 0 else MISA_GREEN)
    _kpi_card(k3, "Critical Challenges",        str(critical),       "#C0392B" if critical > 0 else MISA_GREEN)
    _kpi_card(k4, "Minister-Level Escalations", str(minister_level), "#C0392B" if minister_level > 0 else MISA_GREEN)


def _kpi_card(col, label: str, value: str, color: str):
    col.markdown(f"""
    <div style="background:{color};border-radius:10px;padding:16px;text-align:center;margin-bottom:8px;">
      <div style="color:rgba(255,255,255,0.8);font-size:12px;font-weight:600;">{label}</div>
      <div style="color:white;font-size:28px;font-weight:700;line-height:1.2;">{value}</div>
    </div>
    """, unsafe_allow_html=True)


def _render_deal_charts(df: pd.DataFrame, lang: str):
    col1, col2, col3 = st.columns(3)

    with col1:
        if "Deal Stage" in df.columns:
            sc = df["Deal Stage"].value_counts().reset_index()
            sc.columns = ["Stage", "Count"]
            fig = px.bar(sc, x="Stage", y="Count", title="By Stage",
                         color_discrete_sequence=[MISA_GREEN])
            fig.update_layout(margin=dict(t=30, b=10, l=10, r=10),
                               paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        if "Challenge Severity" in df.columns:
            sev_colors = {
                "Critical": "#C0392B",
                "High":     "#C9974A",
                "Medium":   "#1B5C3F",
                "Low":      "#9B9B9B",
            }
            cc = df["Challenge Severity"].value_counts().reset_index()
            cc.columns = ["Severity", "Count"]
            fig2 = px.pie(cc, values="Count", names="Severity",
                          title="Challenge Severity",
                          color="Severity",
                          color_discrete_map=sev_colors)
            fig2.update_layout(margin=dict(t=30, b=10, l=10, r=10),
                                paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig2, use_container_width=True)

    with col3:
        if "Est. Value (SAR)" in df.columns and "Deal Stage" in df.columns:
            sv = df.groupby("Deal Stage")["Est. Value (SAR)"].sum().reset_index()
            sv = sv[sv["Est. Value (SAR)"] > 0]
            if not sv.empty:
                fig3 = px.bar(sv, x="Deal Stage", y="Est. Value (SAR)",
                              title="Value by Stage (SAR)",
                              color_discrete_sequence=[MISA_GOLD])
                fig3.update_layout(margin=dict(t=30, b=10, l=10, r=10),
                                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig3, use_container_width=True)


def _render_deal_expanders(df: pd.DataFrame):
    for _, row in df.iterrows():
        deal_id   = str(row.get("Deal ID", "—"))
        company   = str(row.get("Company Name", "—"))
        deal_name = str(row.get("Deal Name", "—"))
        severity  = str(row.get("Challenge Severity", ""))
        status    = str(row.get("Deal Status", ""))

        flag = "🔴 " if (severity == "Critical" or status == "Blocked") else ("🟠 " if severity == "High" else "")
        label = f"{flag}{deal_id} — {company} | {deal_name}"

        with st.expander(label, expanded=False):
            c1, c2, c3 = st.columns(3)
            c1.markdown(f"**Deal Stage:** {row.get('Deal Stage', '—')}")
            c1.markdown(f"**Deal Status:** {status or '—'}")
            c1.markdown(f"**Est. Value:** SAR {row.get('Est. Value (SAR)', 0):,.0f}" if pd.notna(row.get("Est. Value (SAR)")) else "**Est. Value:** —")

            c2.markdown(f"**Challenge Classification:** {row.get('Challenge Classification', '—')}")
            c2.markdown(f"**Challenge Severity:** {severity or '—'}")
            c2.markdown(f"**Assigned Owner:** {row.get('Assigned Owner', '—')}")

            c3.markdown(f"**Escalation Required:** {row.get('Escalation Required', '—')}")
            c3.markdown(f"**Escalation Level:** {row.get('Escalation Level', '—')}")
            c3.markdown(f"**Escalation Status:** {row.get('Escalation Status', '—')}")

            st.markdown("---")
            challenge_desc = str(row.get("Challenge Description", "") or "")
            proposed_sol   = str(row.get("Proposed Solution", "") or "")
            notes          = str(row.get("Notes", "") or "")

            if challenge_desc:
                st.markdown(f"**Challenge Description:**\n\n{challenge_desc}")
            if proposed_sol:
                st.markdown(f"**Proposed Solution:**\n\n{proposed_sol}")
            if notes:
                st.markdown(f"**Notes:**\n\n{notes}")

            linked_opp = str(row.get("Linked Opportunity ID", "") or "")
            target_res = row.get("Target Resolution Date")
            last_upd   = row.get("Last Updated")
            if linked_opp:
                st.caption(f"Linked Opportunity ID: {linked_opp}")
            if target_res and pd.notna(target_res):
                st.caption(f"Target Resolution Date: {target_res}")
            if last_upd and pd.notna(last_upd):
                st.caption(f"Last Updated: {last_upd}")


def _add_deal_form(dfs: dict, investors: pd.DataFrame, lang: str):
    company_options = [""] + (
        sorted(investors["Company Name"].dropna().unique().tolist())
        if not investors.empty and "Company Name" in investors.columns else []
    )

    with st.form("add_deal_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        company   = c1.selectbox(t("company_name", lang), company_options)
        deal_name = c2.text_input("Deal Name")

        c3, c4 = st.columns(2)
        linked_opp = c3.text_input("Linked Opportunity ID (optional)")
        deal_stage = c4.selectbox("Deal Stage", DEAL_STAGES)

        c5, c6 = st.columns(2)
        deal_status = c5.selectbox("Deal Status", DEAL_STATUSES)
        est_val     = c6.number_input("Est. Value (SAR)", min_value=0.0, step=1_000_000.0)

        st.markdown("##### Challenge Details")
        challenge_desc = st.text_area("Challenge Description")
        c7, c8 = st.columns(2)
        challenge_class = c7.selectbox("Challenge Classification", CHALLENGE_CLASSIFICATIONS)
        challenge_sev   = c8.selectbox("Challenge Severity", CHALLENGE_SEVERITIES)
        proposed_sol    = st.text_area("Proposed Solution")

        st.markdown("##### Escalation")
        c9, c10 = st.columns(2)
        escalate     = c9.selectbox("Escalation Required", ["No", "Yes"])
        esc_level    = c10.selectbox("Escalation Level", ESCALATION_LEVELS)
        esc_status   = st.selectbox("Escalation Status", ESCALATION_STATUSES)

        st.markdown("##### Assignment")
        c11, c12 = st.columns(2)
        assigned_owner = c11.text_input("Assigned Owner")
        target_res_d   = c12.date_input("Target Resolution Date", value=None)

        notes = st.text_area(t("notes", lang))

        if st.form_submit_button("Add Deal"):
            if not company or not deal_name:
                st.warning("Company and Deal Name are required.")
                return

            deals  = dfs.get("Deal Progress", pd.DataFrame())
            inv_id = _get_investor_id(investors, company)
            new_id = _next_id(deals, "Deal ID", "DEAL")

            new_row = {
                "Deal ID":                  new_id,
                "Investor ID":              inv_id,
                "Company Name":             company,
                "Deal Name":                deal_name,
                "Linked Opportunity ID":    linked_opp or None,
                "Deal Stage":               deal_stage,
                "Deal Status":              deal_status,
                "Est. Value (SAR)":         est_val if est_val > 0 else None,
                "Challenge Description":    challenge_desc or None,
                "Challenge Classification": challenge_class,
                "Challenge Severity":       challenge_sev,
                "Proposed Solution":        proposed_sol or None,
                "Escalation Required":      escalate,
                "Escalation Level":         esc_level,
                "Escalation Status":        esc_status,
                "Assigned Owner":           assigned_owner or None,
                "Target Resolution Date":   target_res_d,
                "Last Updated":             date.today(),
                "Notes":                    notes or None,
            }
            dfs["Deal Progress"] = pd.concat(
                [deals, pd.DataFrame([new_row])], ignore_index=True
            )
            save_session(dfs)
            st.success(f"✅ Deal {new_id} added for {company}")
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
    return f"{prefix}-{(max(nums) + 1 if nums else 1):03d}"
