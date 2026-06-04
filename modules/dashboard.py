
# Executive dashboard — Minister-grade overview with international IR metrics.

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from datetime import date, timedelta

from config.settings import (
    MISA_GREEN, MISA_GOLD, MISA_GREEN_LIGHT, MISA_OFF_WHITE, MISA_GREEN_DARK,
    STATUS_COLORS, TIER_COLORS, PRIORITY_COLORS, JOURNEY_STAGES,
    STALE_THRESHOLD_DAYS, DUE_SOON_DAYS, IR_BENCHMARKS,
)
from config.translations import t


CHART_TEMPLATE = "plotly_white"
CHART_FONT     = dict(family="Inter, Tajawal, sans-serif", size=12)
_PALETTE       = [MISA_GREEN, MISA_GOLD, "#2D7A54", "#E4B96A", "#0F3D2A", "#C0392B"]


# ── Top KPI strip ────────────────────────────────────────────────────────────

def render_kpi_cards(dfs: dict, lang: str):
    investors     = dfs.get("Investor Master",      pd.DataFrame())
    meetings      = dfs.get("Meeting Log",          pd.DataFrame())
    opportunities = dfs.get("Opportunity Pipeline", pd.DataFrame())
    actions       = dfs.get("Action Items",         pd.DataFrame())
    today         = date.today()

    # ── Row 1: Opportunity KPIs ──────────────────────────────────────────────
    total_opps   = len(opportunities)
    active_opps  = _count_col(opportunities, "Opportunity Status", "Active")
    opp_value    = _sum_col(opportunities, "Est. Value (SAR)")
    if opp_value == 0:
        opp_value = _sum_col(opportunities, "Est. Investment Value (SAR)")

    # Sector breakdown for opps
    sector_pcts: list[tuple[str, int]] = []
    if not opportunities.empty and "Sector" in opportunities.columns:
        sec_counts = opportunities["Sector"].dropna().value_counts()
        for sec, cnt in sec_counts.head(4).items():
            sector_pcts.append((str(sec), round(cnt / len(opportunities) * 100)))

    c1, c2, c3 = st.columns(3)
    _big_kpi(c1, "No# Opportunities", f"{total_opps}",
             sub=f"{active_opps} Active  ·  {total_opps - active_opps} Other",
             color=MISA_GREEN)
    _big_kpi(c2, "Opportunities Total Value", _fmt_sar(opp_value),
             sub="Estimated pipeline value",
             color=MISA_GOLD)
    _sector_kpi(c3, "% Opportunities per Sector", sector_pcts)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    # ── Row 2: Investor KPIs ─────────────────────────────────────────────────
    total_inv   = len(investors)
    n_countries = (investors["Country"].dropna().nunique()
                   if not investors.empty and "Country" in investors.columns else 0)
    n_sectors   = (investors["Sector"].dropna().nunique()
                   if not investors.empty and "Sector" in investors.columns else 0)

    inv_with_opps = 0
    if not investors.empty and not opportunities.empty and "Company Name" in opportunities.columns and "Company Name" in investors.columns:
        co_with = set(opportunities["Company Name"].dropna().unique())
        inv_with_opps = int(investors["Company Name"].isin(co_with).sum())

    mtg_month = 0
    if not meetings.empty and "Meeting Date" in meetings.columns:
        m_dates = pd.to_datetime(meetings["Meeting Date"], errors="coerce")
        mtg_month = int((m_dates.dt.month == today.month).sum())

    overdue = _overdue_count(actions, today)

    r1, r2, r3, r4, r5, r6 = st.columns(6)
    _kpi(r1, "Total Investors",        f"{total_inv}",       MISA_GREEN)
    _kpi(r2, "Investor Countries",     f"{n_countries}",     MISA_GREEN_DARK)
    _kpi(r3, "Investor Sectors",       f"{n_sectors}",       MISA_GREEN)
    _kpi(r4, "Investors with Opps",    f"{inv_with_opps}",   MISA_GOLD)
    _kpi(r5, "Meetings This Month",    f"{mtg_month}",       MISA_GREEN)
    _kpi(r6, "Overdue Actions",        f"{overdue}",
         "#C0392B" if overdue > 0 else MISA_GREEN)


def _big_kpi(col, label: str, value: str, sub: str, color: str):
    col.markdown(f"""
    <div style="background:{color};padding:18px 16px 14px 16px;border-radius:10px;
                min-height:110px;display:flex;flex-direction:column;justify-content:center;">
      <div style="color:rgba(255,255,255,0.75);font-size:10px;font-weight:600;
                  letter-spacing:.6px;text-transform:uppercase;margin-bottom:6px;">{label}</div>
      <div style="color:#fff;font-size:32px;font-weight:700;line-height:1.1;">{value}</div>
      <div style="color:rgba(255,255,255,0.75);font-size:11px;margin-top:6px;">{sub}</div>
    </div>""", unsafe_allow_html=True)


def _sector_kpi(col, label: str, items: list[tuple[str, int]]):
    bars_html = ""
    for sec, pct in items:
        bars_html += f"""
        <div style="display:flex;align-items:center;margin-bottom:4px;gap:6px;">
          <span style="width:72px;font-size:10px;color:#374151;white-space:nowrap;
                       overflow:hidden;text-overflow:ellipsis;">{sec}</span>
          <div style="flex:1;background:#e5e7eb;border-radius:3px;height:7px;overflow:hidden;">
            <div style="background:{MISA_GREEN};height:100%;width:{pct}%;"></div>
          </div>
          <span style="width:30px;font-size:10px;font-weight:600;color:{MISA_GREEN};
                       text-align:right;">{pct}%</span>
        </div>"""
    if not items:
        bars_html = "<div style='font-size:11px;color:#9ca3af;'>No sector data</div>"
    col.markdown(f"""
    <div style="background:#fff;border:2px solid {MISA_GREEN};padding:14px 12px;
                border-radius:10px;min-height:110px;">
      <div style="color:#6b7280;font-size:10px;font-weight:600;letter-spacing:.6px;
                  text-transform:uppercase;margin-bottom:10px;">{label}</div>
      {bars_html}
    </div>""", unsafe_allow_html=True)


def _kpi(col, label, value, color):
    col.markdown(f"""
    <div style="background:{color};padding:12px 8px;border-radius:8px;
                text-align:center;min-height:76px;display:flex;
                flex-direction:column;justify-content:center;">
      <div style="color:rgba(255,255,255,0.8);font-size:9px;font-weight:600;
                  letter-spacing:.5px;text-transform:uppercase;margin-bottom:3px;">{label}</div>
      <div style="color:#fff;font-size:20px;font-weight:700;line-height:1.1;">{value}</div>
    </div>""", unsafe_allow_html=True)


# ── Pipeline Overview ─────────────────────────────────────────────────────────

def render_pipeline_overview(dfs: dict, lang: str):
    investors = dfs.get("Investor Master",      pd.DataFrame())
    opps      = dfs.get("Opportunity Pipeline", pd.DataFrame())
    if investors.empty:
        st.info(t("no_data", lang)); return

    col1, col2 = st.columns(2)

    with col1:
        # Journey stage funnel — shows the investor lifecycle
        stage_order = ["Awareness", "Initial Contact", "Engagement",
                        "Opportunity Matching", "Active Negotiation",
                        "Committed", "Post-Investment"]
        if "Journey Stage" in investors.columns:
            jc = investors["Journey Stage"].value_counts().reset_index()
            jc.columns = ["Stage", "Count"]
            # preserve meaningful order where possible
            jc["_ord"] = jc["Stage"].apply(
                lambda s: stage_order.index(s) if s in stage_order else 99)
            jc = jc.sort_values("_ord").drop(columns="_ord")
            fig = px.bar(jc, x="Count", y="Stage", orientation="h",
                         title="Investor Journey Stages",
                         color_discrete_sequence=[MISA_GREEN])
            fig.update_layout(**_layout())
            st.plotly_chart(fig, use_container_width=True)
        elif "Relationship Status" in investors.columns:
            sc = investors["Relationship Status"].value_counts().reset_index()
            sc.columns = ["Status", "Count"]
            fig = px.pie(sc, values="Count", names="Status",
                         color="Status", color_discrete_map=STATUS_COLORS,
                         title=t("chart_by_status", lang), hole=0.35)
            fig.update_traces(textposition="inside", textinfo="percent+label")
            fig.update_layout(**_layout())
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        # Opportunity pipeline stages funnel
        if not opps.empty and "Opportunity Stage" in opps.columns:
            oc = opps["Opportunity Stage"].value_counts().reset_index()
            oc.columns = ["Stage", "Count"]
            fig2 = px.funnel(oc, x="Count", y="Stage",
                             title="Opportunity Stages",
                             color_discrete_sequence=[MISA_GOLD])
            fig2.update_layout(**_layout())
            st.plotly_chart(fig2, use_container_width=True)
        elif "Investor Tier" in investors.columns:
            tc = investors["Investor Tier"].value_counts().reset_index()
            tc.columns = ["Tier", "Count"]
            fig2 = px.bar(tc, x="Count", y="Tier", orientation="h",
                          color="Tier", color_discrete_map=TIER_COLORS,
                          title=t("chart_by_tier", lang))
            fig2.update_layout(**_layout())
            st.plotly_chart(fig2, use_container_width=True)


# ── Meeting Outcomes ──────────────────────────────────────────────────────────

def render_meeting_outcomes(dfs: dict, lang: str):
    meetings = dfs.get("Meeting Log", pd.DataFrame())

    if meetings.empty:
        st.info("No meetings logged yet."); return

    col1, col2 = st.columns(2)

    with col1:
        if "Meeting Status" in meetings.columns:
            sc = meetings["Meeting Status"].value_counts().reset_index()
            sc.columns = ["Status", "Count"]
            fig = px.bar(sc, x="Status", y="Count",
                         color="Status", color_discrete_map=STATUS_COLORS,
                         title=t("dash_meeting_tracker", lang))
            fig.update_layout(**_layout())
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        if "Meeting Date" in meetings.columns:
            m2 = meetings.copy()
            m2["Meeting Date"] = pd.to_datetime(m2["Meeting Date"], errors="coerce")
            monthly = m2.groupby(m2["Meeting Date"].dt.to_period("M")).size().reset_index()
            monthly.columns = ["Month", "Count"]
            monthly["Month"] = monthly["Month"].astype(str)
            fig2 = px.line(monthly, x="Month", y="Count", markers=True,
                           title=t("chart_trend", lang),
                           color_discrete_sequence=[MISA_GREEN])
            fig2.update_layout(**_layout())
            st.plotly_chart(fig2, use_container_width=True)

    if "Logged By" in meetings.columns:
        am_c = meetings["Logged By"].value_counts().reset_index()
        am_c.columns = ["AM", "Meetings"]
        st.markdown(f"**{t('chart_meetings_by_am', lang)}**")
        st.dataframe(am_c, use_container_width=True, hide_index=True)


# ── Investment Flow Monitor ───────────────────────────────────────────────────

def render_investment_flow(dfs: dict, lang: str):
    investors = dfs.get("Investor Master", pd.DataFrame())
    if investors.empty:
        st.info(t("no_data", lang)); return

    col1, col2 = st.columns(2)

    with col1:
        committed = investors[investors.get("Actual Commitment (SAR)", pd.Series()).notna()].copy() if "Actual Commitment (SAR)" in investors.columns else pd.DataFrame()
        if not committed.empty and "Company Name" in committed.columns:
            committed = committed.sort_values("Actual Commitment (SAR)", ascending=True).tail(10)
            fig = px.bar(committed, x="Actual Commitment (SAR)", y="Company Name",
                         orientation="h", title=t("commitment_value", lang),
                         color_discrete_sequence=[MISA_GREEN])
            fig.update_layout(**_layout())
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No commitment data yet.")

    with col2:
        if "Sector" in investors.columns and "Est. Investment Value (SAR)" in investors.columns:
            sv = investors.groupby("Sector")["Est. Investment Value (SAR)"].sum().reset_index()
            sv = sv[sv["Est. Investment Value (SAR)"] > 0]
            if not sv.empty:
                fig2 = px.pie(sv, values="Est. Investment Value (SAR)", names="Sector",
                              title=f"{t('chart_by_sector', lang)} ({t('value_sar', lang)})",
                              color_discrete_sequence=px.colors.sequential.Greens_r)
                fig2.update_layout(**_layout())
                st.plotly_chart(fig2, use_container_width=True)


# ── Opportunity Pipeline ──────────────────────────────────────────────────────

def render_opportunity_pipeline(dfs: dict, lang: str):
    opps = dfs.get("Opportunity Pipeline", pd.DataFrame())
    if opps.empty:
        st.info("No opportunities yet."); return

    col1, col2, col3 = st.columns(3)

    with col1:
        if "Opportunity Stage" in opps.columns:
            sc = opps["Opportunity Stage"].value_counts().reset_index()
            sc.columns = ["Stage", "Count"]
            fig = px.funnel(sc, x="Count", y="Stage",
                            title=t("chart_by_stage", lang),
                            color_discrete_sequence=[MISA_GREEN])
            fig.update_layout(**_layout())
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        if "Confidence Level" in opps.columns:
            cc = opps["Confidence Level"].value_counts().reset_index()
            cc.columns = ["Confidence", "Count"]
            conf_colors = {"High": MISA_GREEN, "Medium": MISA_GOLD,
                           "Low": "#C0392B", "Speculative": "#9B9B9B"}
            fig2 = px.pie(cc, values="Count", names="Confidence",
                          color="Confidence", color_discrete_map=conf_colors,
                          title=t("confidence", lang))
            fig2.update_layout(**_layout())
            st.plotly_chart(fig2, use_container_width=True)

    with col3:
        if "Opportunity Status" in opps.columns:
            stat = opps["Opportunity Status"].value_counts().reset_index()
            stat.columns = ["Status", "Count"]
            fig3 = px.bar(stat, x="Status", y="Count",
                          color="Status", color_discrete_map=STATUS_COLORS,
                          title=t("chart_by_status", lang))
            fig3.update_layout(**_layout())
            st.plotly_chart(fig3, use_container_width=True)


# ── Sector & Geography ────────────────────────────────────────────────────────

def render_sector_geography(dfs: dict, lang: str):
    investors = dfs.get("Investor Master", pd.DataFrame())
    if investors.empty:
        st.info(t("no_data", lang)); return

    col1, col2 = st.columns(2)

    with col1:
        if "Sector" in investors.columns:
            sec = investors["Sector"].value_counts().reset_index()
            sec.columns = ["Sector", "Count"]
            fig = px.pie(sec, values="Count", names="Sector",
                         title=t("chart_by_sector", lang),
                         color_discrete_sequence=px.colors.sequential.Greens_r)
            fig.update_traces(textposition="inside", textinfo="percent+label")
            fig.update_layout(**_layout())
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        if "Country" in investors.columns:
            ctry = investors["Country"].value_counts().reset_index()
            ctry.columns = ["Country", "Count"]
            fig2 = px.bar(ctry.head(12), x="Count", y="Country", orientation="h",
                          title=t("chart_by_country", lang), color="Count",
                          color_continuous_scale=[[0, "#D4EAE0"], [1, MISA_GREEN]])
            fig2.update_layout(**_layout())
            st.plotly_chart(fig2, use_container_width=True)


# ── Minister Decision Panel ───────────────────────────────────────────────────

def render_minister_decision_panel(dfs: dict, lang: str):
    """
    Highlights investors that require the Minister's direct attention —
    decisions pending, escalations, high strategic priority scores.
    """
    investors = dfs.get("Investor Master", pd.DataFrame())
    if investors.empty:
        return

    action_items = []

    # Minister Action Required
    if "Minister Action Required" in investors.columns:
        needs_action = investors[
            investors["Minister Action Required"].notna() &
            (investors["Minister Action Required"] != "None Required") &
            (investors["Minister Action Required"] != "")
        ]
        for _, row in needs_action.iterrows():
            action_items.append({
                "company": row.get("Company Name", "?"),
                "action":  row.get("Minister Action Required", ""),
                "deadline":row.get("Decision Required By", ""),
                "tier":    row.get("Investor Tier", ""),
                "priority":row.get("Strategic Priority Score", ""),
                "color":   "#C0392B",
                "icon":    "🏛️",
            })

    # Tier 1 investors with escalated blockers
    if "Blocker Level" in investors.columns and "Investor Tier" in investors.columns:
        escalated = investors[
            (investors["Investor Tier"] == "Tier 1 — Strategic") &
            (investors["Blocker Level"].isin([
                "Ministerial — requires HE intervention",
                "Cabinet — inter-ministerial coordination required"
            ]))
        ]
        for _, row in escalated.iterrows():
            already = any(i["company"] == row.get("Company Name") for i in action_items)
            if not already:
                action_items.append({
                    "company": row.get("Company Name", "?"),
                    "action":  f"Blocker: {row.get('Blocker Level','')}",
                    "deadline":"",
                    "tier":    row.get("Investor Tier", ""),
                    "priority":row.get("Strategic Priority Score", ""),
                    "color":   "#E67E22",
                    "icon":    "⚠️",
                })

    if not action_items:
        st.success("✓ No items requiring Minister attention at this time.")
        return

    for item in action_items:
        deadline_str = f" — **Due: {item['deadline']}**" if item.get("deadline") else ""
        st.markdown(
            f"""<div style="background:#FFF8F0;border-left:4px solid {item['color']};
            padding:12px 16px;border-radius:0 8px 8px 0;margin-bottom:8px;">
            <span style="font-size:16px;">{item['icon']}</span>
            <strong> {item['company']}</strong>
            <span style="background:{item['color']};color:white;padding:2px 8px;
            border-radius:10px;font-size:11px;margin-left:8px;">{item['tier']}</span>
            <br><span style="font-size:13px;color:#4A4A4A;">{item['action']}{deadline_str}</span>
            </div>""", unsafe_allow_html=True
        )


# ── Vision 2030 Alignment Panel ───────────────────────────────────────────────

def render_vision2030_panel(dfs: dict, lang: str):
    investors = dfs.get("Investor Master", pd.DataFrame())
    opps      = dfs.get("Opportunity Pipeline", pd.DataFrame())

    col1, col2 = st.columns(2)

    with col1:
        if not investors.empty and "Vision 2030 Pillar" in investors.columns:
            vc = investors["Vision 2030 Pillar"].dropna().value_counts().reset_index()
            vc.columns = ["Pillar", "Count"]
            if not vc.empty:
                fig = px.bar(vc, x="Count", y="Pillar", orientation="h",
                             title="Investors by Vision 2030 Pillar",
                             color_discrete_sequence=[MISA_GOLD])
                fig.update_layout(**_layout())
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No Vision 2030 pillar data yet.")
        else:
            st.info("Add 'Vision 2030 Pillar' field to investor records to see alignment.")

    with col2:
        if not investors.empty and "Deal Classification" in investors.columns:
            dc = investors["Deal Classification"].dropna().value_counts().reset_index()
            dc.columns = ["Classification", "Count"]
            if not dc.empty:
                fig2 = px.pie(dc, values="Count", names="Classification",
                              title="Deal Classification Mix",
                              color_discrete_sequence=_PALETTE)
                fig2.update_layout(**_layout())
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.info("No deal classification data yet.")


# ── Economic Impact Panel ─────────────────────────────────────────────────────

def render_economic_impact(dfs: dict, lang: str):
    investors = dfs.get("Investor Master", pd.DataFrame())

    total_pipeline  = _sum_col(investors, "Est. Investment Value (SAR)")
    total_committed = _sum_col(investors, "Actual Commitment (SAR)")
    total_jobs      = _sum_col(investors, "Est. Jobs Created")

    # Saudi content breakdown
    col1, col2, col3 = st.columns(3)

    col1.markdown(f"""
    <div style="background:{MISA_GREEN};padding:16px;border-radius:10px;text-align:center;">
      <div style="color:rgba(255,255,255,0.8);font-size:11px;">Total Pipeline</div>
      <div style="color:#fff;font-size:24px;font-weight:700;">{_fmt_sar(total_pipeline)}</div>
    </div>""", unsafe_allow_html=True)

    col2.markdown(f"""
    <div style="background:{MISA_GOLD};padding:16px;border-radius:10px;text-align:center;">
      <div style="color:rgba(255,255,255,0.8);font-size:11px;">Actual Commitments</div>
      <div style="color:#fff;font-size:24px;font-weight:700;">{_fmt_sar(total_committed)}</div>
    </div>""", unsafe_allow_html=True)

    col3.markdown(f"""
    <div style="background:{MISA_GREEN_LIGHT};padding:16px;border-radius:10px;text-align:center;">
      <div style="color:rgba(255,255,255,0.8);font-size:11px;">Est. Jobs Created</div>
      <div style="color:#fff;font-size:24px;font-weight:700;">
        {f"{int(total_jobs):,}" if total_jobs > 0 else "—"}
      </div>
    </div>""", unsafe_allow_html=True)

    # Saudi content & tech transfer breakdown
    if not investors.empty:
        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
        c4, c5 = st.columns(2)
        with c4:
            if "Saudi Content %" in investors.columns:
                sc = investors["Saudi Content %"].dropna().value_counts().reset_index()
                sc.columns = ["Range", "Count"]
                if not sc.empty:
                    fig = px.pie(sc, values="Count", names="Range",
                                 title="Saudi Content % Distribution",
                                 color_discrete_sequence=px.colors.sequential.Greens_r)
                    fig.update_layout(**_layout())
                    st.plotly_chart(fig, use_container_width=True)
        with c5:
            if "Technology Transfer" in investors.columns:
                tt = investors["Technology Transfer"].dropna().value_counts().reset_index()
                tt.columns = ["Tech Transfer", "Count"]
                if not tt.empty:
                    fig2 = px.bar(tt, x="Tech Transfer", y="Count",
                                  title="Technology Transfer",
                                  color_discrete_sequence=[MISA_GOLD])
                    fig2.update_layout(**_layout())
                    st.plotly_chart(fig2, use_container_width=True)


# ── Strategic Alerts ─────────────────────────────────────────────────────────

def render_alerts(dfs: dict, lang: str):
    today     = date.today()
    actions   = dfs.get("Action Items",    pd.DataFrame())
    investors = dfs.get("Investor Master", pd.DataFrame())
    meetings  = dfs.get("Meeting Log",     pd.DataFrame())

    overdue_items, due_soon_items, stalled, no_engagement = _compute_alerts(
        actions, investors, meetings, today
    )

    # HE / Minister attention items
    he_items = _compute_he_alerts(investors, actions)

    a1, a2, a3, a4, a5 = st.columns(5)
    _alert_box(a1, "Overdue Actions",          overdue_items,  "#C0392B", "🔴")
    _alert_box(a2, "Due Soon",                 due_soon_items, "#C9974A", "🟡")
    _alert_box(a3, "Stalled Relationships",    stalled,        "#E67E22", "⚠")
    _alert_box(a4, "No Recent Engagement",     no_engagement,  "#7F8C8D", "●")
    _alert_box(a5, "Needs HE Attention",       he_items,       "#7C3AED", "★")


def _compute_he_alerts(investors: pd.DataFrame, actions: pd.DataFrame) -> list[str]:
    items = []
    if not investors.empty:
        if "Minister Action Required" in investors.columns:
            for _, row in investors.iterrows():
                v = str(row.get("Minister Action Required", "") or "")
                if v and v.lower() not in ("none required", "none", ""):
                    items.append(f"{row.get('Company Name','?')} — {v[:60]}")
        if "Blocker Level" in investors.columns:
            for _, row in investors.iterrows():
                bl = str(row.get("Blocker Level", "") or "")
                if "ministerial" in bl.lower() or "cabinet" in bl.lower():
                    co = row.get("Company Name", "?")
                    if not any(co in i for i in items):
                        items.append(f"{co} — {bl[:60]}")
        if "Strategic Priority Score" in investors.columns:
            scores = pd.to_numeric(investors["Strategic Priority Score"], errors="coerce")
            hi = investors[scores >= 4]
            for _, row in hi.iterrows():
                co = row.get("Company Name", "?")
                if not any(co in i for i in items):
                    items.append(f"{co} — Strategic priority {row.get('Strategic Priority Score')}")
    if not actions.empty and "Priority" in actions.columns and "Status" in actions.columns:
        hi_acts = actions[
            (actions["Priority"].isin(["Very High"])) &
            (~actions["Status"].isin(["Completed", "Cancelled"]))
        ]
        for _, row in hi_acts.iterrows():
            label = f"{row.get('Company Name','?')} — {str(row.get('Action Description',''))[:50]}"
            if label not in items:
                items.append(label)
    return items


def _compute_alerts(actions, investors, meetings, today):
    overdue, due_soon, stalled, no_eng = [], [], [], []

    if not actions.empty and "Due Date" in actions.columns and "Status" in actions.columns:
        due_dates = pd.to_datetime(actions["Due Date"], errors="coerce")
        for idx, row in actions.iterrows():
            if row.get("Status") in ("Completed", "Cancelled"):
                continue
            d = _safe_date(due_dates.iloc[idx] if idx < len(due_dates) else None)
            label = f"{row.get('Company Name','?')} — {str(row.get('Action Description',''))[:50]}"
            if d and d < today:
                overdue.append(f"🔴 {label} (due {d})")
            elif d and d <= today + timedelta(days=DUE_SOON_DAYS):
                due_soon.append(f"🟡 {label} (due {d})")

    if not investors.empty and "Last Updated" in investors.columns:
        for _, row in investors.iterrows():
            lu = _safe_date(row.get("Last Updated"))
            if lu and (today - lu).days > STALE_THRESHOLD_DAYS:
                stalled.append(f"⚠️ {row.get('Company Name','?')} — {(today-lu).days}d no update")

    if not investors.empty and not meetings.empty and "Meeting Date" in meetings.columns:
        m2 = meetings.copy()
        m2["Meeting Date"] = pd.to_datetime(m2["Meeting Date"], errors="coerce")
        for _, inv in investors.iterrows():
            if inv.get("Relationship Status") in ("Closed", "On Hold"):
                continue
            inv_id = inv.get("Investor ID", "")
            inv_m  = m2[m2.get("Investor ID", pd.Series()) == inv_id] if "Investor ID" in m2.columns else pd.DataFrame()
            if inv_m.empty:
                no_eng.append(f"⚫ {inv.get('Company Name','?')} — No meetings logged")
            else:
                last = inv_m["Meeting Date"].max()
                if pd.notna(last) and (today - last.date()).days > 30:
                    no_eng.append(f"⚫ {inv.get('Company Name','?')} — Last met {last.date()}")

    return overdue, due_soon, stalled, no_eng


def _alert_box(col, title, items, color, icon):
    count = len(items)
    badge = color if count > 0 else MISA_GREEN
    col.markdown(f"""
    <div style="border:2px solid {badge};border-radius:8px;
                padding:12px;min-height:130px;">
      <div style="color:{badge};font-weight:700;font-size:13px;margin-bottom:8px;">
        {icon} {title} ({count})
      </div>""", unsafe_allow_html=True)
    for item in items[:5]:
        col.markdown(f"<div style='font-size:11px;color:#333;margin-bottom:3px;'>{item}</div>",
                     unsafe_allow_html=True)
    if len(items) > 5:
        col.markdown(f"<div style='font-size:10px;color:#888;'>+{len(items)-5} more</div>",
                     unsafe_allow_html=True)
    if not items:
        col.markdown("<div style='font-size:11px;color:#888;'>No items</div>",
                     unsafe_allow_html=True)
    col.markdown("</div>", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _layout() -> dict:
    return dict(
        template=CHART_TEMPLATE,
        font=CHART_FONT,
        margin=dict(t=40, b=20, l=20, r=20),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )


def _safe_date(val):
    if val is None:
        return None
    try:
        ts = pd.to_datetime(val, errors="coerce")
        return ts.date() if pd.notna(ts) else None
    except Exception:
        return None


def _fmt_sar(val: float) -> str:
    if not val or (isinstance(val, float) and pd.isna(val)):
        return "—"
    if val >= 1e9:
        return f"SAR {val/1e9:.1f}B"
    if val >= 1e6:
        return f"SAR {val/1e6:.0f}M"
    return f"SAR {val:,.0f}"


def _sum_col(df: pd.DataFrame, col: str) -> float:
    if df.empty or col not in df.columns:
        return 0.0
    return pd.to_numeric(df[col], errors="coerce").fillna(0).sum()


def _count_col(df: pd.DataFrame, col: str, val: str) -> int:
    if df.empty or col not in df.columns:
        return 0
    return int((df[col] == val).sum())


def _overdue_count(actions: pd.DataFrame, today: date) -> int:
    if actions.empty or "Due Date" not in actions.columns or "Status" not in actions.columns:
        return 0
    due = pd.to_datetime(actions["Due Date"], errors="coerce")
    statuses = actions["Status"].values
    count = 0
    for i, d in enumerate(due):
        if pd.notna(d) and d.date() < today and statuses[i] not in ("Completed", "Cancelled"):
            count += 1
    return count


def _avg_stage_days(investors: pd.DataFrame) -> int:
    if investors.empty or "Last Updated" not in investors.columns:
        return 0
    today = date.today()
    days  = pd.to_datetime(investors["Last Updated"], errors="coerce")
    valid = days.dropna()
    if valid.empty:
        return 0
    diffs = [(today - d.date()).days for d in valid]
    return int(sum(diffs) / len(diffs)) if diffs else 0
