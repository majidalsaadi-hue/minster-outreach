
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
    investors     = dfs.get("Investor Master",    pd.DataFrame())
    meetings      = dfs.get("Meeting Log",        pd.DataFrame())
    opportunities = dfs.get("Opportunity Pipeline", pd.DataFrame())
    actions       = dfs.get("Action Items",       pd.DataFrame())
    today         = date.today()

    total_inv     = len(investors)
    tier1         = _count_col(investors, "Investor Tier", "Tier 1 — Strategic")
    pipeline_val  = _sum_col(investors, "Est. Investment Value (SAR)")
    commitment    = _sum_col(investors, "Actual Commitment (SAR)")
    active_opps   = _count_col(opportunities, "Opportunity Status", "Active")
    overdue       = _overdue_count(actions, today)
    jobs_est      = _sum_col(investors, "Est. Jobs Created")

    mtg_month = 0
    if not meetings.empty and "Meeting Date" in meetings.columns:
        m = pd.to_datetime(meetings["Meeting Date"], errors="coerce")
        mtg_month = int(m.dt.month.eq(today.month).sum())

    # Row 1 — pipeline metrics
    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
    _kpi(c1, t("total_investors", lang),      f"{total_inv}",          MISA_GREEN)
    _kpi(c2, t("tier1_investors", lang),      f"{tier1}",              MISA_GOLD)
    _kpi(c3, t("total_pipeline_value", lang), _fmt_sar(pipeline_val),  MISA_GREEN)
    _kpi(c4, t("commitment_value", lang),     _fmt_sar(commitment),    MISA_GOLD)
    _kpi(c5, t("active_opportunities", lang), f"{active_opps}",        MISA_GREEN)
    _kpi(c6, t("meetings_this_month", lang),  f"{mtg_month}",          MISA_GREEN_LIGHT)
    _kpi(c7, t("overdue_actions", lang),      f"{overdue}",
         "#C0392B" if overdue > 0 else MISA_GREEN)

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    # Row 2 — international IR metrics
    _render_ir_metrics_row(investors, meetings, opportunities, actions, jobs_est, lang)


def _render_ir_metrics_row(investors, meetings, opportunities, actions, jobs_est, lang):
    """International investor relations benchmark KPIs — second strip."""
    today = date.today()

    # Conversion rate: meetings → opportunities (approx via counts)
    total_meetings = len(meetings)
    total_opps     = len(opportunities)
    mtg_to_opp     = round(total_opps / total_meetings * 100, 1) if total_meetings > 0 else 0.0
    bench_mto      = IR_BENCHMARKS["meeting_to_opp_conversion"] * 100

    # Opportunity → commitment conversion
    converted = _count_col(opportunities, "Opportunity Status", "Converted to Deal")
    opp_to_commit = round(converted / total_opps * 100, 1) if total_opps > 0 else 0.0
    bench_otc = IR_BENCHMARKS["opp_to_commitment"] * 100

    # Pipeline coverage ratio vs SAR targets (use actual commitment as denominator)
    commitment_val = _sum_col(investors, "Actual Commitment (SAR)")
    pipeline_val   = _sum_col(investors, "Est. Investment Value (SAR)")
    coverage_ratio = round(pipeline_val / commitment_val, 1) if commitment_val > 0 else 0.0
    bench_cov      = IR_BENCHMARKS["pipeline_coverage_ratio"]

    # Avg days in current stage (proxy via last updated)
    avg_stage_days = _avg_stage_days(investors)

    # Escalation resolution: blocked actions vs total
    blocked = _count_col(actions, "Status", "Blocked") if not actions.empty and "Status" in actions.columns else 0
    escalation_rate = round(blocked / max(len(actions), 1) * 100, 1)

    # Minister attention items (strategic priority 4-5)
    minister_items = 0
    if not investors.empty and "Strategic Priority Score" in investors.columns:
        scores = pd.to_numeric(investors["Strategic Priority Score"], errors="coerce")
        minister_items = int((scores >= 4).sum())
    elif not investors.empty and "Minister Action Required" in investors.columns:
        minister_items = int(
            (investors["Minister Action Required"] != "None Required") &
            (investors["Minister Action Required"].notna())
        )

    r1, r2, r3, r4, r5, r6 = st.columns(6)

    _ir_kpi(r1, "Mtg→Opp Rate",    f"{mtg_to_opp}%",
            f"Benchmark {bench_mto:.0f}%", mtg_to_opp >= bench_mto)
    _ir_kpi(r2, "Deal Win Rate",   f"{opp_to_commit}%",
            f"Benchmark {bench_otc:.0f}%", opp_to_commit >= bench_otc)
    _ir_kpi(r3, "Pipeline Coverage", f"{coverage_ratio:.1f}x",
            f"Benchmark {bench_cov}x", coverage_ratio >= bench_cov)
    _ir_kpi(r4, "Avg Days in Stage", f"{avg_stage_days}d",
            f"Target <{IR_BENCHMARKS['avg_days_to_close']}d", avg_stage_days <= IR_BENCHMARKS["avg_days_to_close"])
    _ir_kpi(r5, "Blocked Actions",  f"{blocked} ({escalation_rate}%)",
            "Target 0%", blocked == 0)
    _ir_kpi(r6, "Minister Items",   f"{minister_items}",
            "Needs HE attention", minister_items == 0)


def _kpi(col, label, value, color):
    col.markdown(f"""
    <div style="background:{color};padding:14px 10px;border-radius:8px;
                text-align:center;height:90px;display:flex;
                flex-direction:column;justify-content:center;">
      <div style="color:rgba(255,255,255,0.8);font-size:10px;font-weight:500;
                  margin-bottom:4px;">{label}</div>
      <div style="color:#fff;font-size:22px;font-weight:700;line-height:1.1;">{value}</div>
    </div>""", unsafe_allow_html=True)


def _ir_kpi(col, label, value, bench_label, on_target: bool):
    border = MISA_GREEN if on_target else "#C0392B"
    icon   = "✓" if on_target else "⚠"
    col.markdown(f"""
    <div style="background:#fff;border:2px solid {border};padding:10px 8px;
                border-radius:8px;text-align:center;height:80px;">
      <div style="font-size:10px;color:#6B6B6B;margin-bottom:2px;">{label}</div>
      <div style="font-size:18px;font-weight:700;color:{border};">{value}</div>
      <div style="font-size:9px;color:{border};">{icon} {bench_label}</div>
    </div>""", unsafe_allow_html=True)


# ── Pipeline Overview ─────────────────────────────────────────────────────────

def render_pipeline_overview(dfs: dict, lang: str):
    investors = dfs.get("Investor Master", pd.DataFrame())
    if investors.empty:
        st.info(t("no_data", lang)); return

    col1, col2 = st.columns(2)

    with col1:
        if "Relationship Status" in investors.columns:
            sc = investors["Relationship Status"].value_counts().reset_index()
            sc.columns = ["Status", "Count"]
            fig = px.pie(sc, values="Count", names="Status",
                         color="Status", color_discrete_map=STATUS_COLORS,
                         title=t("chart_by_status", lang), hole=0.35)
            fig.update_traces(textposition="inside", textinfo="percent+label")
            fig.update_layout(**_layout())
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        if "Investor Tier" in investors.columns:
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

    a1, a2, a3, a4 = st.columns(4)
    _alert_box(a1, t("alerts_overdue",     lang), overdue_items,  "#C0392B", "🔴")
    _alert_box(a2, t("alerts_upcoming",    lang), due_soon_items, "#C9974A", "🟡")
    _alert_box(a3, t("alerts_stalled",     lang), stalled,        "#E67E22", "⚠️")
    _alert_box(a4, t("alerts_no_engagement",lang), no_engagement, "#7F8C8D", "⚫")


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
