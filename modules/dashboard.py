
# Executive dashboard — all panels for the Minister-grade overview.

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from datetime import date, timedelta

from config.settings import (
    MISA_GREEN, MISA_GOLD, MISA_GREEN_LIGHT, MISA_OFF_WHITE,
    STATUS_COLORS, TIER_COLORS, PRIORITY_COLORS, JOURNEY_STAGES,
    STALE_THRESHOLD_DAYS, DUE_SOON_DAYS,
)
from config.translations import t


CHART_TEMPLATE = "plotly_white"
CHART_FONT     = dict(family="Inter, Tajawal, sans-serif", size=12)


# ── Top KPI cards ────────────────────────────────────────────────────────────

def render_kpi_cards(dfs: dict, lang: str):
    investors    = dfs.get("Investor Master", pd.DataFrame())
    meetings     = dfs.get("Meeting Log",     pd.DataFrame())
    opportunities = dfs.get("Opportunity Pipeline", pd.DataFrame())
    actions      = dfs.get("Action Items",    pd.DataFrame())

    total_inv    = len(investors)
    tier1        = len(investors[investors.get("Investor Tier", pd.Series()) == "Tier 1 — Strategic"]) if not investors.empty and "Investor Tier" in investors.columns else 0
    pipeline_val = _sum_col(investors, "Est. Investment Value (SAR)")
    commitment   = _sum_col(investors, "Actual Commitment (SAR)")
    active_opps  = len(opportunities[opportunities.get("Opportunity Status", pd.Series()) == "Active"]) if not opportunities.empty and "Opportunity Status" in opportunities.columns else 0

    today = date.today()
    mtg_this_month = 0
    if not meetings.empty and "Meeting Date" in meetings.columns:
        m_dates = pd.to_datetime(meetings["Meeting Date"], errors="coerce")
        mtg_this_month = int(m_dates.dt.month.eq(today.month).sum())

    overdue = 0
    if not actions.empty and "Due Date" in actions.columns and "Status" in actions.columns:
        due = pd.to_datetime(actions["Due Date"], errors="coerce")
        overdue = int(((due.dt.date < today) & (~actions["Status"].isin(["Completed", "Cancelled"]))).sum())

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    _kpi_card(c1, t("total_investors", lang),    f"{total_inv:,}",          MISA_GREEN)
    _kpi_card(c2, t("tier1_investors", lang),    f"{tier1}",                MISA_GOLD)
    _kpi_card(c3, t("total_pipeline_value", lang), _fmt_sar(pipeline_val), MISA_GREEN_LIGHT)
    _kpi_card(c4, t("active_opportunities", lang), f"{active_opps:,}",     MISA_GREEN)
    _kpi_card(c5, t("meetings_this_month", lang),  f"{mtg_this_month:,}",  MISA_GOLD)
    _kpi_card(c6, t("overdue_actions", lang),      f"{overdue:,}",
              "#C0392B" if overdue > 0 else MISA_GREEN)


def _kpi_card(col, label: str, value: str, color: str):
    col.markdown(f"""
    <div style="background:{color};padding:18px 14px;border-radius:10px;
                text-align:center;height:110px;display:flex;
                flex-direction:column;justify-content:center;">
      <div style="color:rgba(255,255,255,0.85);font-size:12px;
                  font-weight:500;margin-bottom:6px;">{label}</div>
      <div style="color:#FFFFFF;font-size:26px;font-weight:700;
                  line-height:1.1;">{value}</div>
    </div>
    """, unsafe_allow_html=True)


# ── Pipeline Overview ─────────────────────────────────────────────────────────

def render_pipeline_overview(dfs: dict, lang: str):
    investors = dfs.get("Investor Master", pd.DataFrame())
    if investors.empty or "Relationship Status" not in investors.columns:
        st.info(t("no_data", lang))
        return

    col1, col2 = st.columns(2)
    with col1:
        status_counts = investors["Relationship Status"].value_counts().reset_index()
        status_counts.columns = ["Status", "Count"]
        colors = [STATUS_COLORS.get(s, MISA_GREEN) for s in status_counts["Status"]]
        fig = px.pie(status_counts, values="Count", names="Status",
                     color_discrete_sequence=colors,
                     title=t("chart_by_status", lang))
        fig.update_traces(textposition="inside", textinfo="percent+label")
        fig.update_layout(**_layout(lang))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        if "Investor Tier" in investors.columns:
            tier_counts = investors["Investor Tier"].value_counts().reset_index()
            tier_counts.columns = ["Tier", "Count"]
            colors = [TIER_COLORS.get(t_val, MISA_GREEN) for t_val in tier_counts["Tier"]]
            fig2 = px.bar(tier_counts, x="Tier", y="Count",
                          color="Tier", color_discrete_map=TIER_COLORS,
                          title=t("chart_by_tier", lang))
            fig2.update_layout(**_layout(lang))
            st.plotly_chart(fig2, use_container_width=True)


# ── Meeting Outcomes ──────────────────────────────────────────────────────────

def render_meeting_outcomes(dfs: dict, lang: str):
    meetings  = dfs.get("Meeting Log", pd.DataFrame())
    investors = dfs.get("Investor Master", pd.DataFrame())

    if meetings.empty:
        st.info("No meetings logged yet. Upload an Excel with a Meeting Log sheet or add meetings manually.")
        return

    col1, col2 = st.columns(2)

    with col1:
        if "Meeting Status" in meetings.columns:
            sc = meetings["Meeting Status"].value_counts().reset_index()
            sc.columns = ["Status", "Count"]
            colors = [STATUS_COLORS.get(s, MISA_GREEN) for s in sc["Status"]]
            fig = px.bar(sc, x="Status", y="Count",
                         color="Status",
                         color_discrete_map=STATUS_COLORS,
                         title=t("dash_meeting_tracker", lang))
            fig.update_layout(**_layout(lang))
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        if "Meeting Date" in meetings.columns:
            meetings = meetings.copy()
            meetings["Meeting Date"] = pd.to_datetime(meetings["Meeting Date"], errors="coerce")
            monthly = meetings.groupby(meetings["Meeting Date"].dt.to_period("M")).size().reset_index()
            monthly.columns = ["Month", "Count"]
            monthly["Month"] = monthly["Month"].astype(str)
            fig2 = px.line(monthly, x="Month", y="Count",
                           markers=True,
                           title=t("chart_trend", lang),
                           color_discrete_sequence=[MISA_GREEN])
            fig2.update_layout(**_layout(lang))
            st.plotly_chart(fig2, use_container_width=True)

    # AM activity table
    if "Logged By" in meetings.columns and not meetings.empty:
        st.markdown(f"**{t('chart_meetings_by_am', lang)}**")
        am_counts = meetings["Logged By"].value_counts().reset_index()
        am_counts.columns = ["AM", "Meetings Logged"]
        st.dataframe(am_counts, use_container_width=True, hide_index=True)


# ── Investment Flow Monitor ───────────────────────────────────────────────────

def render_investment_flow(dfs: dict, lang: str):
    investors = dfs.get("Investor Master", pd.DataFrame())
    if investors.empty:
        st.info(t("no_data", lang))
        return

    col1, col2 = st.columns(2)

    with col1:
        if "Actual Commitment (SAR)" in investors.columns and "Company Name" in investors.columns:
            committed = investors[investors["Actual Commitment (SAR)"].notna()].copy()
            committed = committed.sort_values("Actual Commitment (SAR)", ascending=True).tail(10)
            if not committed.empty:
                fig = px.bar(committed,
                             x="Actual Commitment (SAR)", y="Company Name",
                             orientation="h",
                             title=t("commitment_value", lang),
                             color_discrete_sequence=[MISA_GREEN])
                fig.update_layout(**_layout(lang))
                st.plotly_chart(fig, use_container_width=True)

    with col2:
        if "Sector" in investors.columns and "Est. Investment Value (SAR)" in investors.columns:
            sector_val = investors.groupby("Sector")["Est. Investment Value (SAR)"].sum().reset_index()
            sector_val = sector_val[sector_val["Est. Investment Value (SAR)"] > 0]
            if not sector_val.empty:
                fig2 = px.pie(sector_val,
                              values="Est. Investment Value (SAR)",
                              names="Sector",
                              title=f"{t('chart_by_sector', lang)} ({t('value_sar', lang)})",
                              color_discrete_sequence=px.colors.sequential.Greens_r)
                fig2.update_layout(**_layout(lang))
                st.plotly_chart(fig2, use_container_width=True)


# ── Opportunity Pipeline ──────────────────────────────────────────────────────

def render_opportunity_pipeline(dfs: dict, lang: str):
    opps = dfs.get("Opportunity Pipeline", pd.DataFrame())
    if opps.empty:
        st.info("No opportunities in pipeline yet.")
        return

    col1, col2, col3 = st.columns(3)

    with col1:
        if "Opportunity Stage" in opps.columns:
            stage_counts = opps["Opportunity Stage"].value_counts().reset_index()
            stage_counts.columns = ["Stage", "Count"]
            fig = px.funnel(stage_counts, x="Count", y="Stage",
                            title=t("chart_by_stage", lang),
                            color_discrete_sequence=[MISA_GREEN])
            fig.update_layout(**_layout(lang))
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        if "Sector" in opps.columns:
            sec = opps["Sector"].value_counts().reset_index()
            sec.columns = ["Sector", "Count"]
            fig2 = px.pie(sec, values="Count", names="Sector",
                          title=t("chart_by_sector", lang),
                          color_discrete_sequence=px.colors.sequential.Teal)
            fig2.update_layout(**_layout(lang))
            st.plotly_chart(fig2, use_container_width=True)

    with col3:
        if "Opportunity Status" in opps.columns:
            stat = opps["Opportunity Status"].value_counts().reset_index()
            stat.columns = ["Status", "Count"]
            colors = [STATUS_COLORS.get(s, MISA_GREEN) for s in stat["Status"]]
            fig3 = px.bar(stat, x="Status", y="Count",
                          color="Status", color_discrete_map=STATUS_COLORS,
                          title=t("chart_by_status", lang))
            fig3.update_layout(**_layout(lang))
            st.plotly_chart(fig3, use_container_width=True)


# ── Sector & Geography ────────────────────────────────────────────────────────

def render_sector_geography(dfs: dict, lang: str):
    investors = dfs.get("Investor Master", pd.DataFrame())
    if investors.empty:
        st.info(t("no_data", lang))
        return

    col1, col2 = st.columns(2)

    with col1:
        if "Sector" in investors.columns:
            sec = investors["Sector"].value_counts().reset_index()
            sec.columns = ["Sector", "Count"]
            fig = px.pie(sec, values="Count", names="Sector",
                         title=t("chart_by_sector", lang),
                         color_discrete_sequence=px.colors.sequential.Greens_r)
            fig.update_traces(textposition="inside", textinfo="percent+label")
            fig.update_layout(**_layout(lang))
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        if "Country" in investors.columns:
            ctry = investors["Country"].value_counts().reset_index()
            ctry.columns = ["Country", "Count"]
            fig2 = px.bar(ctry.head(12), x="Count", y="Country",
                          orientation="h",
                          title=t("chart_by_country", lang),
                          color="Count",
                          color_continuous_scale=[[0, MISA_OFF_WHITE], [1, MISA_GREEN]])
            fig2.update_layout(**_layout(lang))
            st.plotly_chart(fig2, use_container_width=True)


# ── Strategic Alerts ─────────────────────────────────────────────────────────

def render_alerts(dfs: dict, lang: str):
    today = date.today()
    actions   = dfs.get("Action Items",      pd.DataFrame())
    investors = dfs.get("Investor Master",   pd.DataFrame())
    meetings  = dfs.get("Meeting Log",       pd.DataFrame())

    overdue_items, due_soon_items, stalled, no_engagement = _compute_alerts(
        actions, investors, meetings, today
    )

    a1, a2, a3, a4 = st.columns(4)

    with a1:
        _alert_box(t("alerts_overdue", lang), overdue_items, "#C0392B", "🔴")
    with a2:
        _alert_box(t("alerts_upcoming", lang), due_soon_items, "#C9974A", "🟡")
    with a3:
        _alert_box(t("alerts_stalled", lang), stalled, "#E67E22", "⚠️")
    with a4:
        _alert_box(t("alerts_no_engagement", lang), no_engagement, "#7F8C8D", "⚫")


def _compute_alerts(actions, investors, meetings, today):
    overdue_items  = []
    due_soon_items = []
    stalled        = []
    no_engagement  = []

    # Overdue & due-soon actions
    if not actions.empty and "Due Date" in actions.columns and "Status" in actions.columns:
        for _, row in actions.iterrows():
            if row.get("Status") in ("Completed", "Cancelled"):
                continue
            due = _safe_date(row.get("Due Date"))
            if due is None:
                continue
            label = f"{row.get('Company Name','?')} — {str(row.get('Action Description',''))[:50]}"
            if due < today:
                overdue_items.append(f"🔴 {label} (due {due})")
            elif due <= today + timedelta(days=DUE_SOON_DAYS):
                due_soon_items.append(f"🟡 {label} (due {due})")

    # Stalled investors
    if not investors.empty and "Last Updated" in investors.columns:
        for _, row in investors.iterrows():
            lu = _safe_date(row.get("Last Updated"))
            if lu and (today - lu).days > STALE_THRESHOLD_DAYS:
                stalled.append(f"⚠️ {row.get('Company Name','?')} — {(today-lu).days} {t('days_no_update','en')}")

    # No recent engagement (no meeting in 30 days)
    if not investors.empty and not meetings.empty and "Meeting Date" in meetings.columns:
        meetings_copy = meetings.copy()
        meetings_copy["Meeting Date"] = pd.to_datetime(meetings_copy["Meeting Date"], errors="coerce")
        for _, inv in investors.iterrows():
            if inv.get("Relationship Status") in ("Closed", "On Hold"):
                continue
            inv_id = inv.get("Investor ID", "")
            inv_meetings = meetings_copy[meetings_copy.get("Investor ID", pd.Series()) == inv_id]
            if inv_meetings.empty:
                no_engagement.append(f"⚫ {inv.get('Company Name','?')} — No meetings logged")
            else:
                last_mtg = inv_meetings["Meeting Date"].max()
                if pd.notna(last_mtg) and (today - last_mtg.date()).days > 30:
                    no_engagement.append(f"⚫ {inv.get('Company Name','?')} — Last met {last_mtg.date()}")

    return overdue_items, due_soon_items, stalled, no_engagement


def _alert_box(title: str, items: list, color: str, icon: str):
    count = len(items)
    badge_color = color if count > 0 else "#1B5C3F"
    st.markdown(f"""
    <div style="border:2px solid {badge_color};border-radius:8px;padding:12px;min-height:120px;">
      <div style="color:{badge_color};font-weight:700;font-size:14px;margin-bottom:8px;">
        {icon} {title} ({count})
      </div>
    """, unsafe_allow_html=True)
    if items:
        for item in items[:5]:
            st.markdown(f"<div style='font-size:12px;color:#333;margin-bottom:4px;'>{item}</div>",
                        unsafe_allow_html=True)
        if len(items) > 5:
            st.markdown(f"<div style='font-size:11px;color:#888;'>+{len(items)-5} more</div>",
                        unsafe_allow_html=True)
    else:
        st.markdown("<div style='font-size:12px;color:#888;'>No items</div>",
                    unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _layout(lang: str) -> dict:
    return dict(
        template=CHART_TEMPLATE,
        font=CHART_FONT,
        margin=dict(t=40, b=20, l=20, r=20),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )


def _fmt_sar(val: float) -> str:
    if val is None or pd.isna(val):
        return "—"
    if val >= 1e9:
        return f"SAR {val/1e9:.1f}B"
    if val >= 1e6:
        return f"SAR {val/1e6:.0f}M"
    return f"SAR {val:,.0f}"


def _sum_col(df: pd.DataFrame, col: str) -> float:
    if df.empty or col not in df.columns:
        return 0.0
    return pd.to_numeric(df[col], errors="coerce").sum()


def _safe_date(val) -> date | None:
    if val is None:
        return None
    if isinstance(val, date):
        return val
    try:
        return pd.to_datetime(val).date()
    except Exception:
        return None
