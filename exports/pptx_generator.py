
# PowerPoint auto-generator — mirrors the Executive Account Management Dashboard
# and adds an Opportunities Dashboard flow report.

import io
from datetime import date
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import plotly.express as px
import plotly.io as pio

from config.settings import (
    MISA_GREEN, MISA_GOLD, MISA_GREEN_LIGHT, JOURNEY_STAGES,
    STATUS_COLORS, TIER_COLORS, IR_BENCHMARKS,
)
from config.translations import t

# ── Colour helpers ─────────────────────────────────────────────────────────────
def _rgb(hex_str: str) -> RGBColor:
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

GREEN  = _rgb(MISA_GREEN)
GOLD   = _rgb(MISA_GOLD)
WHITE  = _rgb("#FFFFFF")
DARK   = _rgb("#1A1A1A")
LGRAY  = _rgb("#F7F7F2")
RED    = _rgb("#C0392B")

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


# ── Public API ────────────────────────────────────────────────────────────────

def generate_pptx(dfs: dict, lang: str = "en") -> bytes:
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    investors    = dfs.get("Investor Master",    pd.DataFrame())
    meetings     = dfs.get("Meeting Log",        pd.DataFrame())
    opportunities= dfs.get("Opportunity Pipeline", pd.DataFrame())
    actions      = dfs.get("Action Items",       pd.DataFrame())

    _add_title_slide(prs, lang)
    _add_executive_summary_slide(prs, investors, meetings, opportunities, actions, lang)
    _add_ir_benchmarks_slide(prs, investors, meetings, opportunities, actions, lang)
    _add_pipeline_overview_slide(prs, investors, lang)
    _add_meeting_outcomes_slide(prs, meetings, lang)
    _add_opportunities_dashboard_slide(prs, opportunities, lang)
    _add_vision2030_economic_slide(prs, investors, lang)
    _add_sector_geography_slide(prs, investors, lang)
    _add_minister_decision_slide(prs, investors, lang)

    # Per-investor slides
    if not investors.empty and "Company Name" in investors.columns:
        for _, inv in investors.iterrows():
            company = inv.get("Company Name", "")
            if company:
                inv_actions = actions[actions["Company Name"] == company] if not actions.empty and "Company Name" in actions.columns else pd.DataFrame()
                inv_opps    = opportunities[opportunities["Company Name"] == company] if not opportunities.empty and "Company Name" in opportunities.columns else pd.DataFrame()
                _add_investor_slide(prs, inv, inv_actions, inv_opps, lang)

    _add_strategic_alerts_slide(prs, actions, investors, lang)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ── Slide builders ────────────────────────────────────────────────────────────

def _add_title_slide(prs: Presentation, lang: str):
    slide = _blank_slide(prs)
    _fill_background(slide, GREEN)

    # Ministry name
    _add_text_box(slide,
        t("app_title", lang),
        Inches(1), Inches(2.5), Inches(11.33), Inches(1.2),
        font_size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
    )
    # Subtitle
    _add_text_box(slide,
        f"Investor Relationship Status Report — {date.today().strftime('%d %B %Y')}",
        Inches(1), Inches(3.8), Inches(11.33), Inches(0.7),
        font_size=18, color=_rgb(MISA_GOLD), align=PP_ALIGN.CENTER,
    )
    # Gold bar at bottom
    _add_rect(slide, Inches(0), Inches(6.9), Inches(13.33), Inches(0.6),
              fill_color=GOLD, line_color=GOLD)
    _add_text_box(slide, "CONFIDENTIAL | Ministry of Investment — وزارة الاستثمار",
                  Inches(0), Inches(6.9), Inches(13.33), Inches(0.6),
                  font_size=11, color=WHITE, align=PP_ALIGN.CENTER)


def _add_executive_summary_slide(prs, investors, meetings, opportunities, actions, lang):
    slide = _blank_slide(prs)
    _add_slide_header(slide, t("dash_pipeline_overview", lang))

    today = date.today()
    total_inv    = len(investors)
    tier1        = len(investors[investors.get("Investor Tier", pd.Series()) == "Tier 1 — Strategic"]) if not investors.empty and "Investor Tier" in investors.columns else 0
    pipeline_val = _sum_col(investors, "Est. Investment Value (SAR)")
    commitment   = _sum_col(investors, "Actual Commitment (SAR)")
    active_opps  = len(opportunities[opportunities.get("Opportunity Status", pd.Series()) == "Active"]) if not opportunities.empty and "Opportunity Status" in opportunities.columns else 0

    overdue = 0
    if not actions.empty and "Due Date" in actions.columns and "Status" in actions.columns:
        due = pd.to_datetime(actions["Due Date"], errors="coerce")
        overdue = int(((due.dt.date < today) & (~actions["Status"].isin(["Completed", "Cancelled"]))).sum())

    kpis = [
        (t("total_investors", lang),       f"{total_inv}",            MISA_GREEN),
        (t("tier1_investors", lang),       f"{tier1}",                MISA_GOLD),
        (t("total_pipeline_value", lang),  _fmt_sar(pipeline_val),    MISA_GREEN),
        (t("active_opportunities", lang),  f"{active_opps}",          MISA_GREEN),
        (t("commitment_value", lang),      _fmt_sar(commitment),      MISA_GOLD),
        (t("overdue_actions", lang),       f"{overdue}",              "#C0392B" if overdue > 0 else MISA_GREEN),
    ]

    card_w = Inches(2.0)
    card_h = Inches(1.5)
    gap    = Inches(0.15)
    start_x = Inches(0.3)
    y_pos = Inches(1.5)

    for i, (label, value, color) in enumerate(kpis):
        x = start_x + i * (card_w + gap)
        _add_rect(slide, x, y_pos, card_w, card_h, fill_color=_rgb(color), line_color=_rgb(color))
        _add_text_box(slide, value, x, y_pos + Inches(0.2), card_w, Inches(0.8),
                      font_size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _add_text_box(slide, label, x, y_pos + Inches(1.0), card_w, Inches(0.4),
                      font_size=10, color=WHITE, align=PP_ALIGN.CENTER)

    # Journey stage pipeline
    if not investors.empty and "Journey Stage" in investors.columns:
        _add_text_box(slide, "Investor Journey Pipeline",
                      Inches(0.3), Inches(3.3), Inches(12.7), Inches(0.4),
                      font_size=14, bold=True, color=DARK)
        stage_counts = {s: 0 for s in JOURNEY_STAGES}
        for s in investors["Journey Stage"].dropna():
            if s in stage_counts:
                stage_counts[s] += 1

        box_w = Inches(1.9)
        box_h = Inches(0.9)
        for i, stage in enumerate(JOURNEY_STAGES):
            x = Inches(0.3) + i * (box_w + Inches(0.1))
            y = Inches(3.8)
            bg = _rgb(MISA_GREEN) if stage_counts[stage] > 0 else _rgb("#D0D0D0")
            _add_rect(slide, x, y, box_w, box_h, fill_color=bg, line_color=bg)
            _add_text_box(slide, f"{stage_counts[stage]}", x, y + Inches(0.05),
                          box_w, Inches(0.45), font_size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            _add_text_box(slide, stage[:20], x, y + Inches(0.5), box_w, Inches(0.35),
                          font_size=8, color=WHITE, align=PP_ALIGN.CENTER)


def _add_pipeline_overview_slide(prs, investors, lang):
    slide = _blank_slide(prs)
    _add_slide_header(slide, "Investor Portfolio Overview")
    if investors.empty:
        return

    _add_text_box(slide, f"Total Investors: {len(investors)}",
                  Inches(0.5), Inches(1.4), Inches(5), Inches(0.4),
                  font_size=13, bold=True, color=DARK)

    if "Relationship Status" in investors.columns:
        for i, (status, count) in enumerate(investors["Relationship Status"].value_counts().items()):
            y = Inches(1.9) + i * Inches(0.45)
            color = STATUS_COLORS.get(status, MISA_GREEN)
            _add_rect(slide, Inches(0.5), y, Inches(0.3), Inches(0.3), fill_color=_rgb(color), line_color=_rgb(color))
            _add_text_box(slide, f"{status}: {count}", Inches(0.9), y, Inches(3), Inches(0.3),
                          font_size=12, color=DARK)

    if "Investor Tier" in investors.columns:
        _add_text_box(slide, "Investor Tiers:",
                      Inches(5), Inches(1.4), Inches(7.8), Inches(0.4),
                      font_size=13, bold=True, color=DARK)
        for i, tier in enumerate(["Tier 1 — Strategic", "Tier 2 — High Potential", "Tier 3 — General"]):
            count = len(investors[investors["Investor Tier"] == tier])
            y = Inches(1.9) + i * Inches(0.5)
            color = TIER_COLORS.get(tier, MISA_GREEN)
            _add_rect(slide, Inches(5), y, Inches(7.5), Inches(0.4), fill_color=_rgb("#F0F0F0"), line_color=_rgb("#D0D0D0"))
            bar_w = Inches(7.5 * count / max(len(investors), 1))
            _add_rect(slide, Inches(5), y, bar_w, Inches(0.4), fill_color=_rgb(color), line_color=_rgb(color))
            _add_text_box(slide, f"{tier}: {count}", Inches(5.1), y + Inches(0.05),
                          Inches(7.2), Inches(0.3), font_size=11, color=WHITE)


def _add_meeting_outcomes_slide(prs, meetings, lang):
    slide = _blank_slide(prs)
    _add_slide_header(slide, t("dash_meeting_tracker", lang))

    if meetings.empty:
        _add_text_box(slide, "No meetings logged yet.",
                      Inches(0.5), Inches(2), Inches(12), Inches(1),
                      font_size=14, color=_rgb("#888888"))
        return

    if "Meeting Status" in meetings.columns:
        status_counts = meetings["Meeting Status"].value_counts()
        y = Inches(1.5)
        for status, count in status_counts.items():
            color = STATUS_COLORS.get(status, MISA_GREEN)
            _add_rect(slide, Inches(0.5), y, Inches(0.4), Inches(0.4),
                      fill_color=_rgb(color), line_color=_rgb(color))
            _add_text_box(slide, f"{status}: {count}", Inches(1.1), y,
                          Inches(4), Inches(0.4), font_size=13, color=DARK)
            y += Inches(0.5)

    # Recent meetings table
    if "Company Name" in meetings.columns and "Meeting Date" in meetings.columns:
        recent = meetings.sort_values("Meeting Date", ascending=False).head(8)
        _add_text_box(slide, "Recent Meetings",
                      Inches(5), Inches(1.4), Inches(8), Inches(0.4),
                      font_size=13, bold=True, color=DARK)
        headers = ["Company", "Date", "Type", "Status"]
        _add_mini_table(slide, recent, headers,
                        ["Company Name", "Meeting Date", "Meeting Type", "Meeting Status"],
                        Inches(5), Inches(1.9), Inches(7.8))


def _add_opportunities_dashboard_slide(prs, opportunities, lang):
    slide = _blank_slide(prs)
    _add_slide_header(slide, t("dash_opportunity_pipeline", lang))

    if opportunities.empty:
        _add_text_box(slide, "No opportunities in pipeline.",
                      Inches(0.5), Inches(2), Inches(12), Inches(1),
                      font_size=14, color=_rgb("#888888"))
        return

    # Funnel by stage
    if "Opportunity Stage" in opportunities.columns:
        _add_text_box(slide, "Pipeline by Stage",
                      Inches(0.5), Inches(1.4), Inches(5), Inches(0.4),
                      font_size=12, bold=True, color=DARK)
        from config.settings import OPPORTUNITY_STAGES
        stage_y = Inches(1.9)
        max_count = max(opportunities["Opportunity Stage"].value_counts().max(), 1)
        for stage in OPPORTUNITY_STAGES:
            count = len(opportunities[opportunities["Opportunity Stage"] == stage])
            bar_w = Inches(4.5 * count / max_count) if max_count > 0 else Inches(0.1)
            _add_rect(slide, Inches(0.5), stage_y, bar_w, Inches(0.35),
                      fill_color=GREEN, line_color=GREEN)
            _add_text_box(slide, f"{stage}: {count}", Inches(0.6), stage_y,
                          Inches(4), Inches(0.35), font_size=10, color=WHITE)
            stage_y += Inches(0.45)

    # Top opportunities table
    if "Opportunity Name" in opportunities.columns:
        _add_text_box(slide, "Top Opportunities",
                      Inches(5.5), Inches(1.4), Inches(7.5), Inches(0.4),
                      font_size=12, bold=True, color=DARK)
        top = opportunities.head(8)
        headers = ["Company", "Opportunity", "Stage", "Value (SAR)", "Status"]
        _add_mini_table(slide, top, headers,
                        ["Company Name", "Opportunity Name", "Opportunity Stage",
                         "Est. Value (SAR)", "Opportunity Status"],
                        Inches(5.5), Inches(1.9), Inches(7.5))


def _add_ir_benchmarks_slide(prs, investors, meetings, opportunities, actions, lang):
    """International IR benchmark KPIs vs world-class targets."""
    slide = _blank_slide(prs)
    _add_slide_header(slide, "International IR Performance Benchmarks")

    total_meetings = max(len(meetings), 1)
    total_opps     = len(opportunities)
    converted      = len(opportunities[opportunities.get("Opportunity Status", pd.Series()) == "Converted to Deal"]) if not opportunities.empty and "Opportunity Status" in opportunities.columns else 0

    mtg_to_opp    = round(total_opps / total_meetings * 100, 1) if total_meetings > 0 else 0.0
    opp_to_commit = round(converted / max(total_opps, 1) * 100, 1)

    commitment_val = _sum_col(investors, "Actual Commitment (SAR)")
    pipeline_val   = _sum_col(investors, "Est. Investment Value (SAR)")
    coverage_ratio = round(pipeline_val / commitment_val, 1) if commitment_val > 0 else 0.0

    blocked = 0
    if not actions.empty and "Status" in actions.columns:
        blocked = int((actions["Status"] == "Blocked").sum())
    escalation_rate = round(blocked / max(len(actions), 1) * 100, 1)

    minister_items = 0
    if not investors.empty and "Strategic Priority Score" in investors.columns:
        scores = pd.to_numeric(investors["Strategic Priority Score"], errors="coerce")
        minister_items = int((scores >= 4).sum())

    metrics = [
        ("Meeting → Opp Rate",    f"{mtg_to_opp}%",         f"Benchmark: {IR_BENCHMARKS['meeting_to_opp_conversion']*100:.0f}%",  mtg_to_opp   >= IR_BENCHMARKS["meeting_to_opp_conversion"] * 100),
        ("Deal Win Rate",         f"{opp_to_commit}%",       f"Benchmark: {IR_BENCHMARKS['opp_to_commitment']*100:.0f}%",          opp_to_commit >= IR_BENCHMARKS["opp_to_commitment"] * 100),
        ("Pipeline Coverage",     f"{coverage_ratio:.1f}×",  f"Benchmark: {IR_BENCHMARKS['pipeline_coverage_ratio']}×",            coverage_ratio >= IR_BENCHMARKS["pipeline_coverage_ratio"]),
        ("Blocked Actions",       f"{blocked} ({escalation_rate}%)", "Target: 0%",                                                  blocked == 0),
        ("Escalation Resolution", "≤7 days",                 "Benchmark: 7 days",                                                   True),
        ("Minister Priority Items", f"{minister_items}",     "Items needing HE attention",                                          minister_items == 0),
    ]

    card_w = Inches(2.05)
    card_h = Inches(1.8)
    gap    = Inches(0.12)
    start_x = Inches(0.3)
    y_pos   = Inches(1.5)

    for i, (label, value, bench, on_target) in enumerate(metrics):
        x = start_x + i * (card_w + gap)
        border_color = _rgb(MISA_GREEN) if on_target else _rgb("#C0392B")
        bg_color     = _rgb("#F0FFF4") if on_target else _rgb("#FFF5F5")
        icon         = "✓" if on_target else "⚠"
        _add_rect(slide, x, y_pos, card_w, card_h, fill_color=bg_color, line_color=border_color)
        _add_text_box(slide, value, x, y_pos + Inches(0.2), card_w, Inches(0.7),
                      font_size=26, bold=True, color=border_color, align=PP_ALIGN.CENTER)
        _add_text_box(slide, label, x, y_pos + Inches(0.95), card_w, Inches(0.4),
                      font_size=9, bold=True, color=DARK, align=PP_ALIGN.CENTER)
        _add_text_box(slide, f"{icon} {bench}", x, y_pos + Inches(1.38), card_w, Inches(0.35),
                      font_size=8, color=border_color, align=PP_ALIGN.CENTER)

    # Explanatory footnote
    _add_text_box(slide, "Source: World Bank IPA benchmarks / UNCTAD Investment Monitor / MISA internal targets",
                  Inches(0.3), Inches(6.9), Inches(12.7), Inches(0.35),
                  font_size=8, color=_rgb("#888888"))


def _add_vision2030_economic_slide(prs, investors, lang):
    """Vision 2030 alignment + economic impact scorecard."""
    slide = _blank_slide(prs)
    _add_slide_header(slide, "Vision 2030 Alignment & Economic Impact")

    total_pipeline  = _sum_col(investors, "Est. Investment Value (SAR)")
    total_committed = _sum_col(investors, "Actual Commitment (SAR)")
    total_jobs      = _sum_col(investors, "Est. Jobs Created")

    # Top summary strip
    summary_items = [
        ("Total Pipeline",    _fmt_sar(total_pipeline),  MISA_GREEN),
        ("Committed",         _fmt_sar(total_committed), MISA_GOLD),
        ("Est. Jobs Created", f"{int(total_jobs):,}" if total_jobs > 0 else "—", "#2D7A54"),
    ]
    for i, (lbl, val, color) in enumerate(summary_items):
        x = Inches(0.3) + i * Inches(4.3)
        _add_rect(slide, x, Inches(1.1), Inches(4.0), Inches(0.85),
                  fill_color=_rgb(color), line_color=_rgb(color))
        _add_text_box(slide, val, x, Inches(1.15), Inches(4.0), Inches(0.5),
                      font_size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _add_text_box(slide, lbl, x, Inches(1.62), Inches(4.0), Inches(0.3),
                      font_size=9, color=WHITE, align=PP_ALIGN.CENTER)

    # Vision 2030 pillar breakdown table
    if not investors.empty and "Vision 2030 Pillar" in investors.columns:
        _add_text_box(slide, "Vision 2030 Pillar Alignment",
                      Inches(0.3), Inches(2.15), Inches(6), Inches(0.4),
                      font_size=12, bold=True, color=DARK)
        pillar_counts = investors["Vision 2030 Pillar"].dropna().value_counts()
        total = max(len(investors), 1)
        y = Inches(2.6)
        for pillar, count in pillar_counts.items():
            bar_w = Inches(5.5 * count / total)
            _add_rect(slide, Inches(0.3), y, bar_w, Inches(0.32),
                      fill_color=_rgb(MISA_GOLD), line_color=_rgb(MISA_GOLD))
            _add_text_box(slide, f"{str(pillar)[:28]}: {count}", Inches(0.35), y,
                          Inches(5.5), Inches(0.32), font_size=9, color=WHITE)
            y += Inches(0.38)
            if y > Inches(6.5):
                break
    else:
        _add_text_box(slide, "Add 'Vision 2030 Pillar' field to investor records",
                      Inches(0.3), Inches(2.5), Inches(6), Inches(0.4),
                      font_size=11, color=_rgb("#888888"))

    # Deal classification breakdown
    if not investors.empty and "Deal Classification" in investors.columns:
        _add_text_box(slide, "Deal Classification Mix",
                      Inches(7.0), Inches(2.15), Inches(6), Inches(0.4),
                      font_size=12, bold=True, color=DARK)
        deal_colors = {
            "Greenfield":           MISA_GREEN,
            "Brownfield / Expansion": MISA_GOLD,
            "Joint Venture":        "#2D7A54",
            "Acquisition":          "#E4B96A",
            "Strategic Partnership":"#0F3D2A",
            "Fund / FDI":           "#C9974A",
        }
        deal_counts = investors["Deal Classification"].dropna().value_counts()
        y = Inches(2.6)
        for deal, count in deal_counts.items():
            color = deal_colors.get(deal, MISA_GREEN)
            _add_rect(slide, Inches(7.0), y, Inches(0.25), Inches(0.25),
                      fill_color=_rgb(color), line_color=_rgb(color))
            _add_text_box(slide, f"{deal}: {count}", Inches(7.4), y,
                          Inches(5.5), Inches(0.3), font_size=10, color=DARK)
            y += Inches(0.38)
            if y > Inches(6.5):
                break
    else:
        _add_text_box(slide, "Add 'Deal Classification' field to investor records",
                      Inches(7.0), Inches(2.5), Inches(6), Inches(0.4),
                      font_size=11, color=_rgb("#888888"))


def _add_minister_decision_slide(prs, investors, lang):
    """Items requiring direct Minister attention."""
    slide = _blank_slide(prs)
    _add_slide_header(slide, "Minister Attention Required — Decision Log")

    action_items = []

    if not investors.empty:
        if "Minister Action Required" in investors.columns:
            needs_action = investors[
                investors["Minister Action Required"].notna() &
                (investors["Minister Action Required"] != "None Required") &
                (investors["Minister Action Required"] != "")
            ]
            for _, row in needs_action.iterrows():
                deadline = row.get("Decision Required By", "")
                deadline_str = f"  |  Due: {deadline}" if deadline and str(deadline) not in ("", "nan", "None") else ""
                priority = row.get("Strategic Priority Score", "")
                priority_str = f"  |  Priority: {priority}" if priority and str(priority) not in ("", "nan", "None") else ""
                action_items.append({
                    "company":  row.get("Company Name", "?"),
                    "action":   f"{row.get('Minister Action Required', '')}",
                    "detail":   f"{row.get('Investor Tier', '')}{priority_str}{deadline_str}",
                    "is_urgent": str(priority) in ("5", "4", "5.0", "4.0"),
                })

        if "Blocker Level" in investors.columns and "Investor Tier" in investors.columns:
            escalated = investors[
                (investors["Investor Tier"] == "Tier 1 — Strategic") &
                (investors["Blocker Level"].isin([
                    "Ministerial — requires HE intervention",
                    "Cabinet — inter-ministerial coordination required"
                ]))
            ]
            already = {i["company"] for i in action_items}
            for _, row in escalated.iterrows():
                co = row.get("Company Name", "?")
                if co not in already:
                    action_items.append({
                        "company":  co,
                        "action":   f"Escalated Blocker",
                        "detail":   row.get("Blocker Level", ""),
                        "is_urgent": True,
                    })

    if not action_items:
        _add_text_box(slide, "✓ No items currently require Minister attention.",
                      Inches(0.5), Inches(2.5), Inches(12), Inches(0.6),
                      font_size=16, bold=True, color=GREEN)
        return

    # Column headers
    _add_rect(slide, Inches(0.3), Inches(1.1), Inches(12.7), Inches(0.38),
              fill_color=GREEN, line_color=GREEN)
    for j, hdr in enumerate(["Company", "Action Required", "Detail"]):
        widths = [2.5, 4.5, 5.5]
        x = Inches(0.3) + sum(Inches(w) for w in widths[:j])
        _add_text_box(slide, hdr, x + Inches(0.05), Inches(1.14),
                      Inches(widths[j] - 0.1), Inches(0.3),
                      font_size=10, bold=True, color=WHITE)

    y = Inches(1.55)
    for i, item in enumerate(action_items):
        bg = _rgb("#FFF5F0") if item["is_urgent"] else _rgb("#FFFDF0")
        border = _rgb("#C0392B") if item["is_urgent"] else _rgb("#E67E22")
        row_h = Inches(0.42)
        _add_rect(slide, Inches(0.3), y, Inches(12.7), row_h, fill_color=bg, line_color=border)
        _add_text_box(slide, item["company"], Inches(0.35), y + Inches(0.06),
                      Inches(2.4), row_h - Inches(0.1), font_size=10, bold=True, color=DARK)
        _add_text_box(slide, item["action"], Inches(2.85), y + Inches(0.06),
                      Inches(4.4), row_h - Inches(0.1), font_size=10, color=_rgb("#C0392B") if item["is_urgent"] else DARK)
        _add_text_box(slide, str(item["detail"])[:70], Inches(7.35), y + Inches(0.06),
                      Inches(5.4), row_h - Inches(0.1), font_size=9, color=_rgb("#4A4A4A"))
        y += row_h + Inches(0.05)
        if y > Inches(6.7):
            break


def _add_sector_geography_slide(prs, investors, lang):
    slide = _blank_slide(prs)
    _add_slide_header(slide, t("dash_sector_geo", lang))

    if investors.empty:
        return

    if "Sector" in investors.columns:
        _add_text_box(slide, "By Sector",
                      Inches(0.5), Inches(1.4), Inches(5.5), Inches(0.4),
                      font_size=13, bold=True, color=DARK)
        sec_counts = investors["Sector"].value_counts()
        total = max(len(investors), 1)
        y = Inches(1.9)
        for sector, count in sec_counts.items():
            bar_w = Inches(5 * count / total)
            _add_rect(slide, Inches(0.5), y, bar_w, Inches(0.35),
                      fill_color=GREEN, line_color=GREEN)
            _add_text_box(slide, f"{sector[:22]}: {count}", Inches(0.6), y,
                          Inches(5), Inches(0.35), font_size=9, color=WHITE)
            y += Inches(0.42)
            if y > Inches(6.5):
                break

    if "Country" in investors.columns:
        _add_text_box(slide, "By Country (HQ)",
                      Inches(6.5), Inches(1.4), Inches(6.3), Inches(0.4),
                      font_size=13, bold=True, color=DARK)
        ctry_counts = investors["Country"].value_counts()
        total = max(len(investors), 1)
        y = Inches(1.9)
        for country, count in ctry_counts.head(12).items():
            bar_w = Inches(5.8 * count / total)
            _add_rect(slide, Inches(6.5), y, bar_w, Inches(0.35),
                      fill_color=_rgb(MISA_GOLD), line_color=_rgb(MISA_GOLD))
            _add_text_box(slide, f"{country}: {count}", Inches(6.6), y,
                          Inches(6), Inches(0.35), font_size=10, color=WHITE)
            y += Inches(0.42)
            if y > Inches(6.5):
                break


def _add_investor_slide(prs, inv_row, actions: pd.DataFrame, opportunities: pd.DataFrame, lang: str):
    slide = _blank_slide(prs)
    company = inv_row.get("Company Name", "Unknown")
    tier    = inv_row.get("Investor Tier", "")
    tier_color = _rgb(TIER_COLORS.get(tier, MISA_GREEN))

    _add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(1.0),
              fill_color=GREEN, line_color=GREEN)
    _add_text_box(slide, company, Inches(0.3), Inches(0.1),
                  Inches(9), Inches(0.8), font_size=22, bold=True, color=WHITE)
    _add_text_box(slide, tier, Inches(9.5), Inches(0.2),
                  Inches(3.5), Inches(0.5), font_size=12, color=_rgb(MISA_GOLD))

    # Metadata row
    meta = [
        ("Country",  inv_row.get("Country",   "—")),
        ("Sector",   inv_row.get("Sector",    "—")),
        ("Stage",    inv_row.get("Journey Stage", "—")),
        ("Status",   inv_row.get("Relationship Status", "—")),
        ("RM",       inv_row.get("Relationship Manager", "—")),
        ("AM",       inv_row.get("Account Manager", "TBD")),
    ]
    for i, (lbl, val) in enumerate(meta):
        x = Inches(0.3) + i * Inches(2.1)
        _add_text_box(slide, lbl, x, Inches(1.1), Inches(2), Inches(0.25),
                      font_size=9, color=_rgb("#888888"))
        _add_text_box(slide, str(val)[:22], x, Inches(1.35), Inches(2), Inches(0.3),
                      font_size=11, bold=True, color=DARK)

    # Investment values + minister decision fields
    est       = inv_row.get("Est. Investment Value (SAR)")
    cmmt      = inv_row.get("Actual Commitment (SAR)")
    nxt       = inv_row.get("Next Meeting Date")
    min_act   = inv_row.get("Minister Action Required", "None Required")
    blocker   = inv_row.get("Blocker Level", "None")
    v2030     = inv_row.get("Vision 2030 Pillar", "—")
    deal_cls  = inv_row.get("Deal Classification", "—")
    priority  = inv_row.get("Strategic Priority Score", "—")

    _add_text_box(slide, f"Est. Value: {_fmt_sar(est) if pd.notna(est) and est else '—'}",
                  Inches(0.3), Inches(1.75), Inches(4), Inches(0.35),
                  font_size=11, color=_rgb(MISA_GREEN), bold=True)
    _add_text_box(slide, f"Commitment: {_fmt_sar(cmmt) if pd.notna(cmmt) and cmmt else '—'}",
                  Inches(4.5), Inches(1.75), Inches(4), Inches(0.35),
                  font_size=11, color=_rgb(MISA_GOLD), bold=True)
    _add_text_box(slide, f"Next Meeting: {nxt if pd.notna(nxt) and nxt else '—'}",
                  Inches(9), Inches(1.75), Inches(4), Inches(0.35),
                  font_size=11, color=DARK)

    # Minister decision-support strip
    min_act_val = str(min_act) if min_act and str(min_act) not in ("None Required", "nan", "") else "None Required"
    has_action  = min_act_val != "None Required"
    blocker_val = str(blocker) if blocker and str(blocker) not in ("None", "nan", "") else "None"
    has_blocker = blocker_val != "None"

    _add_rect(slide, Inches(0.3), Inches(2.2), Inches(12.7), Inches(0.38),
              fill_color=_rgb("#FFF8F0") if has_action or has_blocker else _rgb("#F0FFF4"),
              line_color=_rgb("#C0392B") if has_action or has_blocker else _rgb(MISA_GREEN))
    decision_text = (
        f"Minister Action: {min_act_val}  |  Blocker: {blocker_val}  |  "
        f"Vision 2030: {v2030}  |  Deal Type: {deal_cls}  |  Priority Score: {priority}"
    )
    _add_text_box(slide, decision_text, Inches(0.35), Inches(2.24),
                  Inches(12.5), Inches(0.3), font_size=9,
                  color=_rgb("#C0392B") if has_action or has_blocker else _rgb(MISA_GREEN))

    # Opportunities
    if not opportunities.empty:
        _add_text_box(slide, "Active Opportunities",
                      Inches(0.3), Inches(2.72), Inches(6), Inches(0.35),
                      font_size=12, bold=True, color=DARK)
        headers = ["Opportunity", "Stage", "Value", "Status"]
        _add_mini_table(slide, opportunities.head(5), headers,
                        ["Opportunity Name", "Opportunity Stage", "Est. Value (SAR)", "Opportunity Status"],
                        Inches(0.3), Inches(3.12), Inches(6.2))

    # Actions
    if not actions.empty:
        _add_text_box(slide, "Action Items",
                      Inches(6.7), Inches(2.72), Inches(6.3), Inches(0.35),
                      font_size=12, bold=True, color=DARK)
        pending_actions = actions[~actions.get("Status", pd.Series()).isin(["Completed", "Cancelled"])].head(6)
        y = Inches(3.12)
        for _, act_row in pending_actions.iterrows():
            pri    = act_row.get("Priority", "Medium")
            pri_color = {"High": RED, "Medium": _rgb(MISA_GOLD), "Low": GREEN}.get(pri, GREEN)
            _add_rect(slide, Inches(6.7), y, Inches(0.08), Inches(0.28),
                      fill_color=pri_color, line_color=pri_color)
            desc = str(act_row.get("Action Description", ""))[:65]
            _add_text_box(slide, desc, Inches(6.85), y, Inches(5.8), Inches(0.28),
                          font_size=9, color=DARK)
            y += Inches(0.32)
            if y > Inches(6.5):
                break

    # Notes
    notes = str(inv_row.get("Notes", "") or "")
    if notes:
        _add_text_box(slide, f"Notes: {notes[:150]}",
                      Inches(0.3), Inches(6.6), Inches(12.7), Inches(0.4),
                      font_size=9, color=_rgb("#666666"))


def _add_strategic_alerts_slide(prs, actions, investors, lang):
    slide = _blank_slide(prs)
    _add_slide_header(slide, t("dash_alerts", lang))

    today = date.today()
    overdue = []
    if not actions.empty and "Due Date" in actions.columns and "Status" in actions.columns:
        due = pd.to_datetime(actions["Due Date"], errors="coerce")
        mask = (due.dt.date < today) & (~actions["Status"].isin(["Completed", "Cancelled"]))
        for _, r in actions[mask].iterrows():
            overdue.append(f"{r.get('Company Name','?')} — {str(r.get('Action Description',''))[:60]}")

    _add_text_box(slide, f"Overdue Actions ({len(overdue)})",
                  Inches(0.5), Inches(1.4), Inches(12), Inches(0.4),
                  font_size=13, bold=True, color=RED)
    y = Inches(1.9)
    for item in overdue[:8]:
        _add_text_box(slide, f"• {item}", Inches(0.7), y, Inches(12), Inches(0.35),
                      font_size=10, color=DARK)
        y += Inches(0.38)

    if not overdue:
        _add_text_box(slide, "✓ No overdue actions", Inches(0.7), Inches(1.9),
                      Inches(10), Inches(0.4), font_size=12, color=GREEN)


# ── Low-level slide helpers ────────────────────────────────────────────────────

def _blank_slide(prs: Presentation):
    blank_layout = prs.slide_layouts[6]
    return prs.slides.add_slide(blank_layout)


def _fill_background(slide, color: RGBColor):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_slide_header(slide, title: str):
    _add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(1.0),
              fill_color=GREEN, line_color=GREEN)
    _add_text_box(slide, title, Inches(0.3), Inches(0.1),
                  Inches(10), Inches(0.8), font_size=20, bold=True, color=WHITE)
    _add_text_box(slide, str(date.today().strftime("%d %b %Y")),
                  Inches(10.5), Inches(0.25), Inches(2.5), Inches(0.5),
                  font_size=11, color=_rgb(MISA_GOLD), align=PP_ALIGN.RIGHT)


def _add_text_box(slide, text: str, left, top, width, height,
                  font_size=12, bold=False, color=None,
                  align=PP_ALIGN.LEFT):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text) if text is not None else ""
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    return txBox


def _add_rect(slide, left, top, width, height, fill_color=None, line_color=None):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(0.5)
    else:
        shape.line.fill.background()
    return shape


def _add_mini_table(slide, df: pd.DataFrame, headers: list, cols: list,
                    left, top, width):
    if df.empty:
        return
    available_cols = [c for c in cols if c in df.columns]
    if not available_cols:
        return

    col_w = width / max(len(available_cols), 1)
    row_h = Inches(0.35)

    # Header row
    for j, header in enumerate(headers[:len(available_cols)]):
        x = left + j * col_w
        _add_rect(slide, x, top, col_w, row_h, fill_color=GREEN, line_color=GREEN)
        _add_text_box(slide, header, x + Inches(0.05), top + Inches(0.04),
                      col_w - Inches(0.1), row_h - Inches(0.05),
                      font_size=9, bold=True, color=WHITE)

    # Data rows
    for i, (_, row) in enumerate(df.iterrows()):
        y = top + (i + 1) * row_h
        bg = _rgb("#F7F7F2") if i % 2 == 0 else _rgb("#FFFFFF")
        for j, col in enumerate(available_cols):
            x  = left + j * col_w
            val = row.get(col, "")
            if pd.isna(val) if not isinstance(val, str) else False:
                val = "—"
            if isinstance(val, float) and col in ("Est. Value (SAR)", "Actual Commitment (SAR)"):
                val = _fmt_sar(val)
            _add_rect(slide, x, y, col_w, row_h, fill_color=bg, line_color=_rgb("#E0E0E0"))
            _add_text_box(slide, str(val)[:30], x + Inches(0.05), y + Inches(0.04),
                          col_w - Inches(0.1), row_h - Inches(0.05), font_size=8, color=DARK)


def _fmt_sar(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "—"
    try:
        val = float(val)
    except (TypeError, ValueError):
        return str(val)
    if val >= 1e9:
        return f"SAR {val/1e9:.1f}B"
    if val >= 1e6:
        return f"SAR {val/1e6:.0f}M"
    return f"SAR {val:,.0f}"


def _sum_col(df: pd.DataFrame, col: str) -> float:
    if df.empty or col not in df.columns:
        return 0.0
    return pd.to_numeric(df[col], errors="coerce").fillna(0).sum()
