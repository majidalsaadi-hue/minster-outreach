
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

_URGENCY_ORDER = {"critical": 0, "today": 1, "soon": 2, "strategic": 3, "watch": 4}

_STRATEGY_PILLARS = {
    "Attract Investment": ["invest", "capital", "fund", "pipeline", "value", "narrative",
                           "promotion", "roadshow", "airport", "infrastructure", "strategy"],
    "Matchmaking":        ["meeting", "introduction", "introductory", "connect", "sector",
                           "ict", "health", "fintech", "data center", "humain", "modon",
                           "stc", "interhealth", "engage", "explore"],
    "Resolve Challenges": ["nda", "residency", "regulatory", "guidance", "energy",
                           "challenge", "resolve", "land", "lease", "fiber", "request",
                           "support", "premium", "cost", "manufacturing"],
}


def _map_to_pillar(text: str) -> str:
    t_lower = text.lower()
    scores = {p: sum(1 for kw in kws if kw in t_lower) for p, kws in _STRATEGY_PILLARS.items()}
    best = max(scores, key=lambda p: scores[p])
    return best if scores[best] > 0 else "Matchmaking"


# ── Progress Overview Chart ───────────────────────────────────────────────────

def render_progress_chart(dfs: dict, lang: str):
    """Compact progress donut chart — Done / In Progress / Due percentages."""
    actions = dfs.get("Action Items", pd.DataFrame())
    if actions.empty or "Status" not in actions.columns:
        return

    total    = len(actions)
    if total == 0:
        return

    n_done   = int(actions["Status"].str.lower().str.contains("complet", na=False).sum())
    n_prog   = int(actions["Status"].isin(["In Progress", "Inprogress"]).sum())
    n_due    = total - n_done - n_prog

    pct_done = round(n_done / total * 100)
    pct_prog = round(n_prog / total * 100)
    pct_due  = 100 - pct_done - pct_prog

    fig = go.Figure(go.Pie(
        labels=["Completed", "In Progress", "Due"],
        values=[n_done, n_prog, max(n_due, 0)],
        hole=0.55,
        marker_colors=[MISA_GREEN, MISA_GOLD, "#9CA3AF"],
        textinfo="percent",
        textfont_size=11,
        showlegend=True,
    ))
    fig.add_annotation(
        text=f"<b>{total}</b><br><span style='font-size:10px'>Actions</span>",
        x=0.5, y=0.5, showarrow=False, font_size=13,
    )
    fig.update_layout(
        margin=dict(t=20, b=10, l=10, r=10),
        height=220,
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5,
                    font=dict(size=10)),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif"),
    )

    c_chart, c_stats = st.columns([2, 1])
    with c_chart:
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    with c_stats:
        st.markdown(f"""
        <div style="padding:12px 0;font-family:'Segoe UI',sans-serif;">
          <div style="margin-bottom:10px;">
            <div style="font-size:10px;color:#6B7280;text-transform:uppercase;letter-spacing:.5px;">Completed</div>
            <div style="font-size:24px;font-weight:700;color:{MISA_GREEN};">{pct_done}%</div>
            <div style="font-size:11px;color:#4B5563;">{n_done} of {total} actions</div>
          </div>
          <div style="margin-bottom:10px;">
            <div style="font-size:10px;color:#6B7280;text-transform:uppercase;letter-spacing:.5px;">In Progress</div>
            <div style="font-size:24px;font-weight:700;color:{MISA_GOLD};">{pct_prog}%</div>
            <div style="font-size:11px;color:#4B5563;">{n_prog} actions</div>
          </div>
          <div>
            <div style="font-size:10px;color:#6B7280;text-transform:uppercase;letter-spacing:.5px;">Due / Pending</div>
            <div style="font-size:24px;font-weight:700;color:#6B7280;">{pct_due}%</div>
            <div style="font-size:11px;color:#4B5563;">{max(n_due,0)} actions</div>
          </div>
        </div>
        """, unsafe_allow_html=True)


# ── Today's Briefing — Action Advisor ────────────────────────────────────────

def render_action_advisor(dfs: dict, lang: str):
    """Ranked 'what should I do today' panel derived from live CRM data."""
    today       = date.today()
    investors   = dfs.get("Investor Master",      pd.DataFrame())
    meetings    = dfs.get("Meeting Log",          pd.DataFrame())
    actions     = dfs.get("Action Items",         pd.DataFrame())
    opps        = dfs.get("Opportunity Pipeline", pd.DataFrame())

    items: list[dict] = []

    # ── 1. Action items: overdue + due this week ──────────────────────────────
    if not actions.empty and "Due Date" in actions.columns and "Status" in actions.columns:
        due_dates = pd.to_datetime(actions["Due Date"], errors="coerce")
        for idx, row in actions.iterrows():
            if str(row.get("Status", "")).strip() in ("Completed", "Cancelled"):
                continue
            d = _safe_date(due_dates.iloc[idx] if idx < len(due_dates) else None)
            if not d:
                continue
            company = str(row.get("Company Name", "?"))
            desc    = str(row.get("Action Description", ""))[:90]
            owner   = str(row.get("Assigned To", "—"))
            prio    = str(row.get("Priority", ""))
            days_late = (today - d).days
            if days_late > 0:
                items.append({
                    "score":    120 + days_late * 3,
                    "urgency":  "critical",
                    "tag":      f"DUE {d.strftime('%d %b').upper()}",
                    "tag_color":"#D97706",
                    "company":  company,
                    "action":   desc or "Complete pending action",
                    "detail":   f"Due {d.strftime('%d %b %Y')} · {owner}",
                    "icon":     "🟠",
                })
            elif (d - today).days == 0:
                items.append({
                    "score":    100,
                    "urgency":  "critical",
                    "tag":      "DUE TODAY",
                    "tag_color":"#B45309",
                    "company":  company,
                    "action":   desc or "Complete action",
                    "detail":   f"Due today · {owner}",
                    "icon":     "⚡",
                })
            elif (d - today).days <= 7:
                items.append({
                    "score":    60 + max(0, 7 - (d - today).days) * 4,
                    "urgency":  "soon",
                    "tag":      f"DUE {d.strftime('%d %b').upper()}",
                    "tag_color":"#D97706",
                    "company":  company,
                    "action":   desc or "Complete action",
                    "detail":   f"In {(d-today).days}d · {owner}",
                    "icon":     "🟡",
                })

    # ── 2. Meetings: today + next 5 days ─────────────────────────────────────
    if not meetings.empty and "Meeting Date" in meetings.columns:
        mtg = meetings.copy()
        mtg["Meeting Date"] = pd.to_datetime(mtg["Meeting Date"], errors="coerce")
        for _, row in mtg.iterrows():
            d = _safe_date(row.get("Meeting Date"))
            if not d:
                continue
            delta   = (d - today).days
            company = str(row.get("Company Name", "?"))
            mtype   = str(row.get("Meeting Type",  "Meeting"))
            loc     = str(row.get("Location",      ""))
            obj     = str(row.get("Meeting Objective", ""))[:80]
            status  = str(row.get("Meeting Status", "")).lower()
            if status in ("cancelled", "completed"):
                continue
            if delta == 0:
                items.append({
                    "score":    110,
                    "urgency":  "today",
                    "tag":      "TODAY",
                    "tag_color":"#1D4ED8",
                    "company":  company,
                    "action":   f"{mtype} — confirm attendance & prepare briefing",
                    "detail":   loc or obj or "Check meeting details",
                    "icon":     "📅",
                })
            elif 1 <= delta <= 3:
                items.append({
                    "score":    80,
                    "urgency":  "soon",
                    "tag":      d.strftime("%a %d %b").upper(),
                    "tag_color":"#0891B2",
                    "company":  company,
                    "action":   f"Prepare for {mtype}",
                    "detail":   obj or f"In {delta} day{'s' if delta!=1 else ''}",
                    "icon":     "📋",
                })
            elif 4 <= delta <= 7:
                items.append({
                    "score":    50,
                    "urgency":  "soon",
                    "tag":      d.strftime("%a %d %b").upper(),
                    "tag_color":"#6B7280",
                    "company":  company,
                    "action":   f"Upcoming {mtype} — review account notes",
                    "detail":   obj or f"In {delta} days",
                    "icon":     "🗓",
                })

    # ── 3. Recent meeting follow-ups (Next Steps not blank) ───────────────────
    if not meetings.empty and "Meeting Date" in meetings.columns and "Next Steps" in meetings.columns:
        mtg2 = meetings.copy()
        mtg2["Meeting Date"] = pd.to_datetime(mtg2["Meeting Date"], errors="coerce")
        recent = mtg2[
            (mtg2["Next Steps"].notna()) &
            (mtg2["Next Steps"].astype(str).str.strip() != "") &
            (mtg2["Next Steps"].astype(str).str.strip().str.lower() != "nan")
        ].sort_values("Meeting Date", ascending=False)
        seen_co: set[str] = set()
        for _, row in recent.iterrows():
            d = _safe_date(row.get("Meeting Date"))
            if not d:
                continue
            days_ago = (today - d).days
            if days_ago > 21:
                continue
            company = str(row.get("Company Name", "?"))
            if company in seen_co:
                continue
            seen_co.add(company)
            ns = str(row.get("Next Steps", ""))[:100]
            items.append({
                "score":    70 - days_ago,
                "urgency":  "soon",
                "tag":      f"FOLLOW UP",
                "tag_color":"#7C3AED",
                "company":  company,
                "action":   ns,
                "detail":   f"From meeting {days_ago} day{'s' if days_ago!=1 else ''} ago",
                "icon":     "↩",
            })

    # ── 4. Strategic: high-priority investors with no recent contact ──────────
    if not investors.empty:
        last_mtg_map: dict[str, date] = {}
        if not meetings.empty and "Meeting Date" in meetings.columns and "Company Name" in meetings.columns:
            m3 = meetings.copy()
            m3["Meeting Date"] = pd.to_datetime(m3["Meeting Date"], errors="coerce")
            for co, ts in m3.groupby("Company Name")["Meeting Date"].max().items():
                d = _safe_date(ts)
                if d:
                    last_mtg_map[str(co)] = d

        for _, row in investors.iterrows():
            co     = str(row.get("Company Name", "?"))
            status = str(row.get("Relationship Status", "")).lower()
            if status in ("churned", "inactive", "closed"):
                continue
            last  = last_mtg_map.get(co)
            score = float(row.get("Strategic Priority Score", 0) or 0)
            days_since = (today - last).days if last else 999

            if score >= 4 and days_since > 30:
                items.append({
                    "score":    55 + int(score) * 6,
                    "urgency":  "strategic",
                    "tag":      "STRATEGIC",
                    "tag_color":"#7C3AED",
                    "company":  co,
                    "action":   "Schedule outreach — high-priority investor needs engagement",
                    "detail":   f"Last contact: {days_since}d ago · Priority {score}",
                    "icon":     "★",
                })
            elif not last and status not in ("churned", "closed", "inactive"):
                journey = str(row.get("Journey Stage", ""))
                items.append({
                    "score":    30,
                    "urgency":  "watch",
                    "tag":      "NO CONTACT",
                    "tag_color":"#6B7280",
                    "company":  co,
                    "action":   "Initiate first contact",
                    "detail":   f"No meetings on record · {journey or 'Stage unknown'}",
                    "icon":     "○",
                })
            elif last and days_since > 60 and status == "active":
                items.append({
                    "score":    35,
                    "urgency":  "watch",
                    "tag":      "STALLED",
                    "tag_color":"#E67E22",
                    "company":  co,
                    "action":   "Re-engage — no contact for over 60 days",
                    "detail":   f"Last meeting: {last.strftime('%d %b %Y')} ({days_since}d ago)",
                    "icon":     "⚠",
                })

    # ── 5. Stalled opportunities ──────────────────────────────────────────────
    if not opps.empty:
        for _, row in opps.iterrows():
            st_val = str(row.get("Opportunity Status", "")).lower()
            if st_val in ("closed", "won", "lost", "cancelled", "completed"):
                continue
            lu = _safe_date(row.get("Last Updated") or row.get("Start Date"))
            if lu and (today - lu).days > 60:
                val = row.get("Est. Value (SAR)", 0) or 0
                opp_name = str(row.get("Opportunity Name", ""))[:70]
                items.append({
                    "score":    40,
                    "urgency":  "strategic",
                    "tag":      "OPP STALLED",
                    "tag_color":"#E67E22",
                    "company":  str(row.get("Company Name", "?")),
                    "action":   f"Re-engage on opportunity: {opp_name}",
                    "detail":   f"No update in {(today-lu).days}d · {_fmt_sar(val)}",
                    "icon":     "💼",
                })

    # ── 6. Minister / HE decisions pending ───────────────────────────────────
    if not investors.empty and "Minister Action Required" in investors.columns:
        for _, row in investors.iterrows():
            v = str(row.get("Minister Action Required", "") or "").strip()
            if v and v.lower() not in ("none required", "none", ""):
                dl = _safe_date(row.get("Decision Required By"))
                detail = f"Deadline: {dl.strftime('%d %b %Y')}" if dl else "No deadline set"
                items.append({
                    "score":    130,
                    "urgency":  "critical",
                    "tag":      "HE ACTION",
                    "tag_color":"#DC2626",
                    "company":  str(row.get("Company Name", "?")),
                    "action":   v[:100],
                    "detail":   detail,
                    "icon":     "🏛",
                })

    # ── Render — single self-contained HTML block ─────────────────────────────
    items.sort(key=lambda x: x["score"], reverse=True)
    n_crit  = sum(1 for i in items if i["urgency"] in ("critical", "today"))
    n_soon  = sum(1 for i in items if i["urgency"] == "soon")
    n_strat = sum(1 for i in items if i["urgency"] in ("strategic", "watch"))
    n_done  = sum(1 for i in items if i.get("tag", "").startswith("COMPLETED"))
    day_label = today.strftime("%A, %d %B %Y")

    # ── Header bar (always render as one complete block) ──────────────────────
    st.markdown(
        f'<div style="background:linear-gradient(135deg,#0f2d1e 0%,#1B5C3F 100%);'
        f'border-radius:12px 12px 0 0;padding:16px 20px 14px 20px;">'
        f'<div style="display:flex;align-items:center;justify-content:space-between;">'
        f'<div>'
        f'<span style="color:#C9974A;font-size:10px;font-weight:700;letter-spacing:.8px;text-transform:uppercase;">Today\'s Briefing</span>'
        f'<div style="color:#fff;font-size:17px;font-weight:700;margin-top:2px;">{day_label}</div>'
        f'</div>'
        f'<div style="display:flex;gap:10px;">'
        f'<div style="text-align:center;background:rgba(220,38,38,0.25);border:1px solid rgba(220,38,38,0.5);border-radius:8px;padding:6px 14px;">'
        f'<div style="color:#FCA5A5;font-size:18px;font-weight:700;">{n_crit}</div>'
        f'<div style="color:rgba(255,255,255,0.6);font-size:9px;letter-spacing:.4px;">CRITICAL</div></div>'
        f'<div style="text-align:center;background:rgba(217,119,6,0.25);border:1px solid rgba(217,119,6,0.5);border-radius:8px;padding:6px 14px;">'
        f'<div style="color:#FCD34D;font-size:18px;font-weight:700;">{n_soon}</div>'
        f'<div style="color:rgba(255,255,255,0.6);font-size:9px;letter-spacing:.4px;">THIS WEEK</div></div>'
        f'<div style="text-align:center;background:rgba(124,58,237,0.25);border:1px solid rgba(124,58,237,0.5);border-radius:8px;padding:6px 14px;">'
        f'<div style="color:#C4B5FD;font-size:18px;font-weight:700;">{n_strat}</div>'
        f'<div style="color:rgba(255,255,255,0.6);font-size:9px;letter-spacing:.4px;">STRATEGIC</div></div>'
        f'</div></div></div>',
        unsafe_allow_html=True,
    )

    if not items:
        st.markdown(
            '<div style="background:#0f2d1e;border-radius:0 0 12px 12px;padding:16px 20px;'
            'text-align:center;color:rgba(255,255,255,0.6);font-size:13px;margin-bottom:16px;">'
            'All clear — no pending actions or alerts at this time.</div>',
            unsafe_allow_html=True,
        )
        return

    # ── Group items by company ────────────────────────────────────────────────
    from collections import defaultdict
    company_groups: dict[str, list] = defaultdict(list)
    for item in items:
        company_groups[item["company"]].append(item)

    # Sort companies by highest score within the group
    sorted_companies = sorted(
        company_groups.keys(),
        key=lambda c: max(i["score"] for i in company_groups[c]),
        reverse=True,
    )

    # ── Company boxes in a 2-column Streamlit layout ──────────────────────────
    cols = st.columns(2)
    for ci, company in enumerate(sorted_companies):
        co_items = company_groups[company]
        n_co_crit = sum(1 for i in co_items if i["urgency"] in ("critical", "today"))
        worst_color = co_items[0]["tag_color"]

        rows_html = ""
        for item in co_items:
            tc = item["tag_color"]
            rows_html += (
                f'<div style="display:flex;align-items:flex-start;gap:8px;'
                f'padding:6px 8px;margin-bottom:4px;border-radius:6px;'
                f'background:#f8f9fa;border-left:3px solid {tc};">'
                f'<span style="background:{tc};color:#fff;padding:1px 5px;border-radius:3px;'
                f'font-size:8px;font-weight:700;white-space:nowrap;flex-shrink:0;">{item["tag"]}</span>'
                f'<div style="min-width:0;">'
                f'<div style="color:#1F2937;font-size:11px;font-weight:500;line-height:1.3;">{item["action"]}</div>'
                f'<div style="color:#6B7280;font-size:9px;margin-top:2px;">{item["detail"]}</div>'
                f'</div></div>'
            )

        if n_co_crit:
            badge = (f'<span style="background:#FEE2E2;color:#991B1B;'
                     f'padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;">'
                     f'{n_co_crit} critical</span>')
        else:
            badge = (f'<span style="background:#FEF3C7;color:#92400E;'
                     f'padding:2px 8px;border-radius:10px;font-size:10px;">'
                     f'{len(co_items)} item{"s" if len(co_items)!=1 else ""}</span>')

        box_html = (
            f'<div style="border:1px solid {worst_color}50;border-radius:10px;'
            f'padding:12px 14px;margin-bottom:12px;background:#fff;'
            f'box-shadow:0 1px 4px rgba(0,0,0,0.08);">'
            f'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;">'
            f'<span style="color:{MISA_GREEN};font-size:13px;font-weight:700;">{company}</span>'
            f'{badge}'
            f'</div>'
            f'{rows_html}'
            f'</div>'
        )
        with cols[ci % 2]:
            st.markdown(box_html, unsafe_allow_html=True)

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)


# ── Active Opportunities — Minister View ─────────────────────────────────────

def render_active_opportunities_panel(dfs: dict, lang: str):
    """
    Minister-grade investment opportunity panel shown on the dashboard.
    Shows approved active opportunities grouped by company with total value.
    Falls back to Action Items with Opportunity type if pipeline is empty.
    """
    _STAGE_C = {
        "Exploration":        "#6B7280",
        "Due Diligence":      "#1D4ED8",
        "Active Negotiation": "#D97706",
        "Committed":          "#059669",
        "Post-Investment":    MISA_GREEN,
    }

    opps    = dfs.get("Opportunity Pipeline", pd.DataFrame())
    actions = dfs.get("Action Items",         pd.DataFrame())

    # Fallback: synthesise from action items when pipeline is empty
    if opps.empty and not actions.empty and "Type of Engagement" in actions.columns:
        opp_acts = actions[
            actions["Type of Engagement"].str.strip().str.lower() == "opportunity"
        ].copy()
        if not opp_acts.empty:
            opp_acts["Opportunity Status"] = "Active"
            opp_acts["Opportunity Stage"]  = "Exploration"
            opp_acts["Confidence Level"]   = "Medium"
            opps = opp_acts.rename(columns={"Action Description": "Opportunity Name"})

    if opps.empty:
        return

    # Only show active / under-review rows
    if "Opportunity Status" in opps.columns:
        active = opps[
            opps["Opportunity Status"].str.lower().isin(["active", "under review", ""])
        ].copy()
        if active.empty:
            active = opps.copy()
    else:
        active = opps.copy()

    total_val   = _sum_col(active, "Est. Value (SAR)")
    n_active    = len(active)
    n_committed = int((active.get("Opportunity Stage", pd.Series()) == "Committed").sum()) \
                  if "Opportunity Stage" in active.columns else 0
    n_companies = active["Company Name"].dropna().nunique() \
                  if "Company Name" in active.columns else 0

    # ── Gradient header ───────────────────────────────────────────────────────
    st.markdown(
        f'<div style="background:linear-gradient(135deg,#071a0f 0%,{MISA_GREEN} 100%);'
        f'border-radius:14px;padding:20px 26px;margin-bottom:14px;">'
        f'<div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;">'
        f'<div>'
        f'<div style="color:{MISA_GOLD};font-size:10px;font-weight:700;letter-spacing:1px;'
        f'text-transform:uppercase;margin-bottom:5px;">Active Investment Opportunities</div>'
        f'<div style="color:#fff;font-size:26px;font-weight:700;line-height:1.1;">'
        f'{n_active} Opportunities</div>'
        f'<div style="color:rgba(255,255,255,0.6);font-size:12px;margin-top:5px;">'
        f'{n_companies} companies &nbsp;·&nbsp; {n_committed} committed'
        f'{"&nbsp;·&nbsp;" + str(n_active - n_committed) + " in progress" if n_active - n_committed > 0 else ""}'
        f'</div>'
        f'</div>'
        f'<div style="text-align:right;">'
        f'<div style="color:rgba(255,255,255,0.5);font-size:11px;margin-bottom:4px;">Total Pipeline Value</div>'
        f'<div style="color:{MISA_GOLD};font-size:32px;font-weight:700;">'
        f'{_fmt_sar(total_val) if total_val else "—"}</div>'
        f'</div>'
        f'</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    if "Company Name" not in active.columns:
        return

    companies = active["Company Name"].dropna().unique().tolist()
    cols = st.columns(min(3, len(companies)) if companies else 1)
    for i, co in enumerate(companies[:6]):
        co_opps   = active[active["Company Name"] == co]
        co_val    = _sum_col(co_opps, "Est. Value (SAR)")
        n_co      = len(co_opps)
        n_co_comm = int((co_opps.get("Opportunity Stage", pd.Series()) == "Committed").sum()) \
                    if "Opportunity Stage" in co_opps.columns else 0

        opp_rows = ""
        for _, row in co_opps.iterrows():
            stage = str(row.get("Opportunity Stage", "Exploration"))
            name  = str(row.get("Opportunity Name",  "—"))[:70]
            val   = row.get("Est. Value (SAR)")
            sc    = _STAGE_C.get(stage, "#6B7280")
            vs    = _fmt_sar(float(val)) if val and str(val) not in ("nan", "None", "") else ""
            opp_rows += (
                f'<div style="display:flex;align-items:center;gap:6px;'
                f'padding:5px 8px;border-radius:5px;margin-bottom:3px;'
                f'background:#f9fafb;border-left:3px solid {sc};">'
                f'<div style="flex:1;min-width:0;">'
                f'<div style="font-size:11px;font-weight:600;color:#111827;'
                f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">{name}</div>'
                f'<div style="font-size:9px;color:{sc};font-weight:600;">{stage}</div>'
                f'</div>'
                f'{"<span style=font-size:10px;font-weight:700;color:" + MISA_GOLD + ";flex-shrink:0;>" + vs + "</span>" if vs else ""}'
                f'</div>'
            )

        with cols[i % len(cols)]:
            st.markdown(
                f'<div style="border:1px solid #E5E7EB;border-radius:10px;padding:12px;'
                f'margin-bottom:10px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.05);">'
                f'<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px;">'
                f'<div>'
                f'<div style="font-size:13px;font-weight:700;color:{MISA_GREEN};">{co}</div>'
                f'<div style="font-size:9px;color:#9CA3AF;margin-top:2px;">'
                f'{n_co} opp{"s" if n_co!=1 else ""}'
                f'{"  ·  " + str(n_co_comm) + " committed" if n_co_comm else ""}'
                f'</div>'
                f'</div>'
                f'{"<div style=font-size:12px;font-weight:700;color:" + MISA_GOLD + ";>" + _fmt_sar(co_val) + "</div>" if co_val else ""}'
                f'</div>'
                f'{opp_rows}'
                f'</div>',
                unsafe_allow_html=True,
            )


# ── Top KPI strip ────────────────────────────────────────────────────────────

def render_kpi_cards(dfs: dict, lang: str):
    investors     = dfs.get("Investor Master",      pd.DataFrame())
    meetings      = dfs.get("Meeting Log",          pd.DataFrame())
    opportunities = dfs.get("Opportunity Pipeline", pd.DataFrame())
    actions       = dfs.get("Action Items",         pd.DataFrame())
    today         = date.today()

    # ── Row 1: Opportunity KPIs ──────────────────────────────────────────────
    # If Opportunity Pipeline is empty, count from Action Items (Type of Engagement=Opportunity)
    _opps_source = opportunities
    if opportunities.empty and not actions.empty and "Type of Engagement" in actions.columns:
        _opp_acts = actions[
            actions["Type of Engagement"].str.strip().str.lower() == "opportunity"
        ].copy()
        if not _opp_acts.empty:
            _opp_acts["Opportunity Status"] = "Active"
            _opps_source = _opp_acts.rename(columns={"Action Description": "Opportunity Name"})

    total_opps   = len(_opps_source)
    active_opps  = _count_col(_opps_source, "Opportunity Status", "Active")
    opp_value    = _sum_col(_opps_source, "Est. Value (SAR)")
    if opp_value == 0:
        opp_value = _sum_col(investors, "Est. Investment Value (SAR)")

    # Sector breakdown for opps
    sector_pcts: list[tuple[str, int]] = []
    _sec_src = _opps_source if not _opps_source.empty else investors
    if not _sec_src.empty and "Sector" in _sec_src.columns:
        sec_counts = _sec_src["Sector"].dropna().replace("", pd.NA).dropna().value_counts()
        total_sec  = sec_counts.sum() or 1
        for sec, cnt in sec_counts.head(4).items():
            sector_pcts.append((str(sec), round(cnt / total_sec * 100)))

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
    if not investors.empty and "Company Name" in investors.columns:
        # Primary: count from Opportunity Pipeline
        if not opportunities.empty and "Company Name" in opportunities.columns:
            co_with = set(opportunities["Company Name"].dropna().unique())
            inv_with_opps = int(investors["Company Name"].isin(co_with).sum())
        # Fallback: count from Action Items where Type of Engagement = Opportunity
        elif not actions.empty and "Type of Engagement" in actions.columns and "Company Name" in actions.columns:
            opp_cos = set(
                actions[actions["Type of Engagement"].str.strip().str.lower() == "opportunity"]
                ["Company Name"].dropna().unique()
            )
            inv_with_opps = int(investors["Company Name"].isin(opp_cos).sum())

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
    col.markdown(
        f'<div style="background:{color};padding:18px 16px 14px 16px;border-radius:10px;'
        f'min-height:110px;display:flex;flex-direction:column;justify-content:center;">'
        f'<div style="color:rgba(255,255,255,0.75);font-size:10px;font-weight:600;'
        f'letter-spacing:.6px;text-transform:uppercase;margin-bottom:6px;">{label}</div>'
        f'<div style="color:#fff;font-size:32px;font-weight:700;line-height:1.1;">{value}</div>'
        f'<div style="color:rgba(255,255,255,0.75);font-size:11px;margin-top:6px;">{sub}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


def _sector_kpi(col, label: str, items: list[tuple[str, int]]):
    bars_html = ""
    for sec, pct in items:
        bars_html += (
            f'<div style="display:flex;align-items:center;margin-bottom:4px;gap:6px;">'
            f'<span style="width:72px;font-size:10px;color:#374151;white-space:nowrap;'
            f'overflow:hidden;text-overflow:ellipsis;">{sec}</span>'
            f'<div style="flex:1;background:#e5e7eb;border-radius:3px;height:7px;overflow:hidden;">'
            f'<div style="background:{MISA_GREEN};height:100%;width:{pct}%;"></div>'
            f'</div>'
            f'<span style="width:30px;font-size:10px;font-weight:600;color:{MISA_GREEN};'
            f'text-align:right;">{pct}%</span>'
            f'</div>'
        )
    if not items:
        bars_html = "<div style='font-size:11px;color:#9ca3af;'>No sector data</div>"
    col.markdown(
        f'<div style="background:#fff;border:2px solid {MISA_GREEN};padding:14px 12px;'
        f'border-radius:10px;min-height:110px;">'
        f'<div style="color:#6b7280;font-size:10px;font-weight:600;letter-spacing:.6px;'
        f'text-transform:uppercase;margin-bottom:10px;">{label}</div>'
        f'{bars_html}'
        f'</div>',
        unsafe_allow_html=True,
    )


def _kpi(col, label, value, color):
    col.markdown(
        f'<div style="background:{color};padding:12px 8px;border-radius:8px;'
        f'text-align:center;min-height:76px;display:flex;'
        f'flex-direction:column;justify-content:center;">'
        f'<div style="color:rgba(255,255,255,0.8);font-size:9px;font-weight:600;'
        f'letter-spacing:.5px;text-transform:uppercase;margin-bottom:3px;">{label}</div>'
        f'<div style="color:#fff;font-size:20px;font-weight:700;line-height:1.1;">{value}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


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

    col1.markdown(
        f'<div style="background:{MISA_GREEN};padding:16px;border-radius:10px;text-align:center;">'
        f'<div style="color:rgba(255,255,255,0.8);font-size:11px;">Total Pipeline</div>'
        f'<div style="color:#fff;font-size:24px;font-weight:700;">{_fmt_sar(total_pipeline)}</div>'
        f'</div>', unsafe_allow_html=True)

    col2.markdown(
        f'<div style="background:{MISA_GOLD};padding:16px;border-radius:10px;text-align:center;">'
        f'<div style="color:rgba(255,255,255,0.8);font-size:11px;">Actual Commitments</div>'
        f'<div style="color:#fff;font-size:24px;font-weight:700;">{_fmt_sar(total_committed)}</div>'
        f'</div>', unsafe_allow_html=True)

    jobs_str = f"{int(total_jobs):,}" if total_jobs > 0 else "—"
    col3.markdown(
        f'<div style="background:{MISA_GREEN_LIGHT};padding:16px;border-radius:10px;text-align:center;">'
        f'<div style="color:rgba(255,255,255,0.8);font-size:11px;">Est. Jobs Created</div>'
        f'<div style="color:#fff;font-size:24px;font-weight:700;">{jobs_str}</div>'
        f'</div>', unsafe_allow_html=True)

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
    col.markdown(
        f'<div style="border:2px solid {badge};border-radius:8px;padding:12px;min-height:130px;">'
        f'<div style="color:{badge};font-weight:700;font-size:13px;margin-bottom:8px;">'
        f'{icon} {title} ({count})</div>',
        unsafe_allow_html=True,
    )
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
