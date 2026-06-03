
# PowerPoint auto-generator — compact deck with merged sections.
# Full deck: Title + 4 summary slides + per-investor slides.
# Company deck: 3 slides.

import io
import calendar as _calendar
from datetime import date, timedelta
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from config.settings import (
    MISA_GREEN, MISA_GOLD, MISA_GREEN_LIGHT, JOURNEY_STAGES,
    STATUS_COLORS, TIER_COLORS, IR_BENCHMARKS,
)
from config.translations import t

# ── Colours ───────────────────────────────────────────────────────────────────

def _rgb(hex_str: str) -> RGBColor:
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

GREEN  = _rgb(MISA_GREEN)
GOLD   = _rgb(MISA_GOLD)
WHITE  = _rgb("#FFFFFF")
DARK   = _rgb("#1A1A1A")
LGRAY  = _rgb("#F7F7F2")
RED    = _rgb("#C0392B")
MGRAY  = _rgb("#888888")

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


# ── Public API ────────────────────────────────────────────────────────────────

def generate_pptx(dfs: dict, lang: str = "en") -> bytes:
    """
    Full portfolio deck — 5 summary slides + one slide per investor.
    Slide 1: Title
    Slide 2: Executive Summary  (KPIs + Journey funnel + overdue alerts)
    Slide 3: Portfolio Snapshot (Sector/Geo left | Tier/Status/Meetings right)
    Slide 4: Pipeline & Performance (IR benchmarks left | Top opportunities right)
    Slide 5: Minister Log + Vision 2030 (merged)
    Slides 6+: Per-investor (one per company)
    """
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    investors     = dfs.get("Investor Master",      pd.DataFrame())
    meetings      = dfs.get("Meeting Log",          pd.DataFrame())
    opportunities = dfs.get("Opportunity Pipeline", pd.DataFrame())
    actions       = dfs.get("Action Items",         pd.DataFrame())

    _slide_title(prs, lang)
    _slide_executive_summary(prs, investors, meetings, opportunities, actions, lang)
    _slide_portfolio_snapshot(prs, investors, meetings, lang)
    _slide_pipeline_performance(prs, investors, meetings, opportunities, actions, lang)
    _slide_minister_vision(prs, investors, lang)

    # Per-investor slides
    if not investors.empty and "Company Name" in investors.columns:
        for _, inv in investors.iterrows():
            co = inv.get("Company Name", "")
            if not co:
                continue
            inv_acts = actions[actions["Company Name"] == co] if not actions.empty and "Company Name" in actions.columns else pd.DataFrame()
            inv_opps = opportunities[opportunities["Company Name"] == co] if not opportunities.empty and "Company Name" in opportunities.columns else pd.DataFrame()
            _slide_investor(prs, inv, inv_acts, inv_opps, lang)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def generate_pptx_company(dfs: dict, company: str, lang: str = "en") -> bytes:
    """
    Single-company deck — 3 slides.
    Slide 1: Cover + Profile (metadata + key metrics)
    Slide 2: Meetings (left) + Opportunities (right)
    Slide 3: Action Items (left) + Deals (right)
    """
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    investors     = dfs.get("Investor Master",      pd.DataFrame())
    meetings      = dfs.get("Meeting Log",          pd.DataFrame())
    opportunities = dfs.get("Opportunity Pipeline", pd.DataFrame())
    actions       = dfs.get("Action Items",         pd.DataFrame())
    deals         = dfs.get("Deal Progress",        pd.DataFrame())

    inv_row  = investors[investors["Company Name"] == company].iloc[0] if not investors.empty and "Company Name" in investors.columns and company in investors["Company Name"].values else pd.Series()
    inv_mtgs = meetings[meetings["Company Name"] == company]       if not meetings.empty      and "Company Name" in meetings.columns      else pd.DataFrame()
    inv_opps = opportunities[opportunities["Company Name"] == company] if not opportunities.empty and "Company Name" in opportunities.columns else pd.DataFrame()
    inv_acts = actions[actions["Company Name"] == company]         if not actions.empty       and "Company Name" in actions.columns       else pd.DataFrame()
    inv_dls  = deals[deals["Company Name"] == company]             if not deals.empty         and "Company Name" in deals.columns         else pd.DataFrame()

    _co_slide_cover_profile(prs, company, inv_row, inv_opps, inv_acts, inv_mtgs, lang)
    _co_slide_meetings_opps(prs, company, inv_mtgs, inv_opps, lang)
    _co_slide_actions_deals(prs, company, inv_acts, inv_dls, lang)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# FULL DECK — slide builders
# ══════════════════════════════════════════════════════════════════════════════

def _slide_title(prs, lang):
    slide = _blank_slide(prs)
    _fill_background(slide, GREEN)
    _add_text_box(slide, t("app_title", lang),
                  Inches(1), Inches(2.3), Inches(11.33), Inches(1.2),
                  font_size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    _add_text_box(slide,
                  f"Investor Relationship Status Report — {date.today().strftime('%d %B %Y')}",
                  Inches(1), Inches(3.7), Inches(11.33), Inches(0.7),
                  font_size=18, color=GOLD, align=PP_ALIGN.CENTER)
    _add_rect(slide, Inches(0), Inches(6.9), Inches(13.33), Inches(0.6),
              fill_color=GOLD, line_color=GOLD)
    _add_text_box(slide, "CONFIDENTIAL | Ministry of Investment — وزارة الاستثمار",
                  Inches(0), Inches(6.9), Inches(13.33), Inches(0.6),
                  font_size=11, color=WHITE, align=PP_ALIGN.CENTER)


def _slide_executive_summary(prs, investors, meetings, opportunities, actions, lang):
    """KPI strip + Journey funnel + Overdue alerts — all on one slide."""
    slide = _blank_slide(prs)
    _add_slide_header(slide, "Executive Summary")

    today = date.today()

    # ── KPI strip (6 cards) ─────────────────────────────────────────
    total_inv    = len(investors)
    tier1        = len(investors[investors["Investor Tier"] == "Tier 1 — Strategic"]) if not investors.empty and "Investor Tier" in investors.columns else 0
    pipeline_val = _sum_col(investors, "Est. Investment Value (SAR)")
    commitment   = _sum_col(investors, "Actual Commitment (SAR)")
    active_opps  = len(opportunities[opportunities["Opportunity Status"] == "Active"]) if not opportunities.empty and "Opportunity Status" in opportunities.columns else 0

    overdue = 0
    if not actions.empty and "Due Date" in actions.columns and "Status" in actions.columns:
        due = pd.to_datetime(actions["Due Date"], errors="coerce")
        overdue = int(((due.dt.date < today) & (~actions["Status"].isin(["Completed", "Cancelled"]))).sum())

    kpis = [
        ("Investors",         f"{total_inv}",            MISA_GREEN),
        ("Tier 1 Strategic",  f"{tier1}",                MISA_GOLD),
        ("Pipeline Value",    _fmt_sar(pipeline_val),    MISA_GREEN),
        ("Commitments",       _fmt_sar(commitment),      MISA_GOLD),
        ("Active Opps",       f"{active_opps}",          MISA_GREEN),
        ("Overdue Actions",   f"{overdue}",              "#C0392B" if overdue > 0 else MISA_GREEN),
    ]
    card_w, card_h, gap = Inches(2.05), Inches(1.3), Inches(0.12)
    for i, (label, value, color) in enumerate(kpis):
        x = Inches(0.3) + i * (card_w + gap)
        _add_rect(slide, x, Inches(1.1), card_w, card_h, fill_color=_rgb(color), line_color=_rgb(color))
        _add_text_box(slide, value, x, Inches(1.15), card_w, Inches(0.7),
                      font_size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _add_text_box(slide, label, x, Inches(1.82), card_w, Inches(0.3),
                      font_size=9, color=WHITE, align=PP_ALIGN.CENTER)

    # ── Journey funnel ──────────────────────────────────────────────
    _add_text_box(slide, "Investor Journey Pipeline",
                  Inches(0.3), Inches(2.6), Inches(8), Inches(0.35),
                  font_size=12, bold=True, color=DARK)
    if not investors.empty and "Journey Stage" in investors.columns:
        stage_counts = {s: 0 for s in JOURNEY_STAGES}
        for s in investors["Journey Stage"].dropna():
            if s in stage_counts:
                stage_counts[s] += 1
        box_w = Inches(1.9)
        for i, stage in enumerate(JOURNEY_STAGES):
            x = Inches(0.3) + i * (box_w + Inches(0.12))
            y = Inches(3.05)
            bg = GREEN if stage_counts[stage] > 0 else _rgb("#D0D0D0")
            _add_rect(slide, x, y, box_w, Inches(0.9), fill_color=bg, line_color=bg)
            _add_text_box(slide, str(stage_counts[stage]),
                          x, y + Inches(0.04), box_w, Inches(0.42),
                          font_size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            _add_text_box(slide, stage[:20], x, y + Inches(0.48), box_w, Inches(0.38),
                          font_size=7, color=WHITE, align=PP_ALIGN.CENTER)

    # ── Overdue alerts (right side) ─────────────────────────────────
    _add_text_box(slide, f"⚠ Overdue Actions ({overdue})",
                  Inches(0.3), Inches(4.2), Inches(12.7), Inches(0.35),
                  font_size=12, bold=True, color=RED if overdue > 0 else GREEN)
    if overdue > 0 and not actions.empty and "Due Date" in actions.columns:
        due = pd.to_datetime(actions["Due Date"], errors="coerce")
        mask = (due.dt.date < today) & (~actions["Status"].isin(["Completed", "Cancelled"]))
        od_rows = actions[mask].head(6)
        y = Inches(4.65)
        cols_left  = od_rows.iloc[:3]
        cols_right = od_rows.iloc[3:]
        for _, r in cols_left.iterrows():
            txt = f"• {r.get('Company Name','?')} — {str(r.get('Action Description',''))[:55]}"
            _add_text_box(slide, txt, Inches(0.3), y, Inches(6.3), Inches(0.32), font_size=9, color=DARK)
            y += Inches(0.35)
        y = Inches(4.65)
        for _, r in cols_right.iterrows():
            txt = f"• {r.get('Company Name','?')} — {str(r.get('Action Description',''))[:55]}"
            _add_text_box(slide, txt, Inches(6.8), y, Inches(6.3), Inches(0.32), font_size=9, color=DARK)
            y += Inches(0.35)
    elif overdue == 0:
        _add_text_box(slide, "✓ No overdue actions — on track",
                      Inches(0.5), Inches(4.65), Inches(12), Inches(0.35),
                      font_size=11, color=GREEN)


def _slide_portfolio_snapshot(prs, investors, meetings, lang):
    """Left: Sector/Geography breakdown. Right: Tier + Relationship status + recent meetings."""
    slide = _blank_slide(prs)
    _add_slide_header(slide, "Portfolio Snapshot")

    # ── LEFT: Sector breakdown ──────────────────────────────────────
    _add_text_box(slide, "By Sector", Inches(0.3), Inches(1.1), Inches(5.5), Inches(0.35),
                  font_size=12, bold=True, color=DARK)
    if not investors.empty and "Sector" in investors.columns:
        sec = investors["Sector"].value_counts()
        total = max(len(investors), 1)
        y = Inches(1.55)
        for sector, count in sec.items():
            bar_w = Inches(5.5 * count / total)
            _add_rect(slide, Inches(0.3), y, bar_w, Inches(0.32), fill_color=GREEN, line_color=GREEN)
            _add_text_box(slide, f"{str(sector)[:22]}: {count}", Inches(0.35), y,
                          Inches(5.4), Inches(0.32), font_size=9, color=WHITE)
            y += Inches(0.4)
            if y > Inches(4.5):
                break

    # ── MIDDLE: Country breakdown ────────────────────────────────────
    _add_text_box(slide, "By Country", Inches(6.0), Inches(1.1), Inches(4.0), Inches(0.35),
                  font_size=12, bold=True, color=DARK)
    if not investors.empty and "Country" in investors.columns:
        ctry = investors["Country"].value_counts()
        total = max(len(investors), 1)
        y = Inches(1.55)
        for country, count in ctry.head(8).items():
            bar_w = Inches(3.8 * count / total)
            _add_rect(slide, Inches(6.0), y, bar_w, Inches(0.32),
                      fill_color=_rgb(MISA_GOLD), line_color=_rgb(MISA_GOLD))
            _add_text_box(slide, f"{str(country)[:18]}: {count}", Inches(6.05), y,
                          Inches(3.8), Inches(0.32), font_size=9, color=WHITE)
            y += Inches(0.4)
            if y > Inches(4.5):
                break

    # ── RIGHT: Tier + Status ────────────────────────────────────────
    _add_text_box(slide, "Investor Tiers & Status", Inches(10.2), Inches(1.1), Inches(2.9), Inches(0.35),
                  font_size=12, bold=True, color=DARK)
    y = Inches(1.55)
    if not investors.empty and "Investor Tier" in investors.columns:
        for tier in ["Tier 1 — Strategic", "Tier 2 — High Potential", "Tier 3 — General"]:
            count = len(investors[investors["Investor Tier"] == tier])
            color = TIER_COLORS.get(tier, MISA_GREEN)
            _add_rect(slide, Inches(10.2), y, Inches(2.9), Inches(0.3),
                      fill_color=_rgb(color), line_color=_rgb(color))
            label = tier.split("—")[0].strip()
            _add_text_box(slide, f"{label}: {count}", Inches(10.25), y,
                          Inches(2.8), Inches(0.3), font_size=9, color=WHITE)
            y += Inches(0.38)

    y += Inches(0.15)
    if not investors.empty and "Relationship Status" in investors.columns:
        for status, count in investors["Relationship Status"].value_counts().items():
            color = STATUS_COLORS.get(status, MISA_GREEN)
            _add_rect(slide, Inches(10.2), y, Inches(2.9), Inches(0.28),
                      fill_color=_rgb(color), line_color=_rgb(color))
            _add_text_box(slide, f"{status}: {count}", Inches(10.25), y,
                          Inches(2.8), Inches(0.28), font_size=8, color=WHITE)
            y += Inches(0.35)

    # ── BOTTOM: Recent meetings mini-table ──────────────────────────
    _add_rect(slide, Inches(0.3), Inches(4.6), Inches(12.73), Inches(0.32),
              fill_color=GREEN, line_color=GREEN)
    _add_text_box(slide, "Recent Meetings", Inches(0.35), Inches(4.62),
                  Inches(12.6), Inches(0.28), font_size=10, bold=True, color=WHITE)
    if not meetings.empty and "Meeting Date" in meetings.columns:
        recent = meetings.sort_values("Meeting Date", ascending=False).head(5)
        hdrs  = ["Company", "Date", "Type", "Status", "Objective"]
        c_map = ["Company Name", "Meeting Date", "Meeting Type", "Meeting Status", "Meeting Objective"]
        _add_mini_table(slide, recent, hdrs, c_map, Inches(0.3), Inches(5.0), Inches(12.73))


def _slide_pipeline_performance(prs, investors, meetings, opportunities, actions, lang):
    """Left: IR benchmarks. Right: top opportunities."""
    slide = _blank_slide(prs)
    _add_slide_header(slide, "Pipeline & Performance")

    # ── LEFT: IR benchmarks ─────────────────────────────────────────
    _add_text_box(slide, "IR Performance vs. Benchmarks",
                  Inches(0.3), Inches(1.1), Inches(6.0), Inches(0.35),
                  font_size=12, bold=True, color=DARK)

    total_meetings = max(len(meetings), 1)
    total_opps     = max(len(opportunities), 1)
    converted = len(opportunities[opportunities["Opportunity Status"] == "Converted to Deal"]) if not opportunities.empty and "Opportunity Status" in opportunities.columns else 0
    mtg_to_opp    = round(len(opportunities) / total_meetings * 100, 1)
    opp_to_commit = round(converted / total_opps * 100, 1)
    pipeline_val  = _sum_col(investors, "Est. Investment Value (SAR)")
    commitment    = _sum_col(investors, "Actual Commitment (SAR)")
    coverage      = round(pipeline_val / commitment, 1) if commitment > 0 else 0.0
    blocked       = int((actions["Status"] == "Blocked").sum()) if not actions.empty and "Status" in actions.columns else 0

    benchmarks = [
        ("Meeting → Opp",   f"{mtg_to_opp}%",    f"Target {IR_BENCHMARKS['meeting_to_opp_conversion']*100:.0f}%", mtg_to_opp   >= IR_BENCHMARKS["meeting_to_opp_conversion"] * 100),
        ("Deal Win Rate",   f"{opp_to_commit}%",  f"Target {IR_BENCHMARKS['opp_to_commitment']*100:.0f}%",        opp_to_commit >= IR_BENCHMARKS["opp_to_commitment"] * 100),
        ("Pipeline Cover",  f"{coverage:.1f}×",   f"Target {IR_BENCHMARKS['pipeline_coverage_ratio']}×",          coverage      >= IR_BENCHMARKS["pipeline_coverage_ratio"]),
        ("Blocked Actions", f"{blocked}",          "Target: 0",                                                    blocked == 0),
    ]
    y = Inches(1.6)
    for label, value, bench, ok in benchmarks:
        color  = _rgb(MISA_GREEN) if ok else RED
        bg     = _rgb("#F0FFF4")  if ok else _rgb("#FFF5F5")
        _add_rect(slide, Inches(0.3), y, Inches(5.8), Inches(0.7), fill_color=bg, line_color=color)
        _add_text_box(slide, value, Inches(0.35), y + Inches(0.04), Inches(1.5), Inches(0.38),
                      font_size=20, bold=True, color=color)
        _add_text_box(slide, label, Inches(1.9), y + Inches(0.04), Inches(2.5), Inches(0.3),
                      font_size=10, bold=True, color=DARK)
        _add_text_box(slide, f"{'✓' if ok else '⚠'} {bench}", Inches(1.9), y + Inches(0.35), Inches(2.5), Inches(0.28),
                      font_size=8, color=color)
        y += Inches(0.82)

    # Stage funnel
    _add_text_box(slide, "Opportunity Stages",
                  Inches(0.3), y + Inches(0.1), Inches(6), Inches(0.3),
                  font_size=10, bold=True, color=DARK)
    if not opportunities.empty and "Opportunity Stage" in opportunities.columns:
        from config.settings import OPPORTUNITY_STAGES
        max_c = max(opportunities["Opportunity Stage"].value_counts().max(), 1)
        y2 = y + Inches(0.5)
        for stage in OPPORTUNITY_STAGES:
            count = len(opportunities[opportunities["Opportunity Stage"] == stage])
            bar_w = Inches(5.5 * count / max_c) if max_c > 0 else Inches(0.1)
            _add_rect(slide, Inches(0.3), y2, bar_w, Inches(0.28), fill_color=GREEN, line_color=GREEN)
            _add_text_box(slide, f"{stage}: {count}", Inches(0.35), y2,
                          Inches(5.5), Inches(0.28), font_size=9, color=WHITE)
            y2 += Inches(0.35)

    # ── RIGHT: Top opportunities table ──────────────────────────────
    _add_text_box(slide, "Opportunity Pipeline",
                  Inches(6.5), Inches(1.1), Inches(6.5), Inches(0.35),
                  font_size=12, bold=True, color=DARK)
    if not opportunities.empty:
        top_opps = opportunities.head(12)
        hdrs  = ["Company", "Opportunity", "Stage", "Value", "Status"]
        c_map = ["Company Name", "Opportunity Name", "Opportunity Stage", "Est. Value (SAR)", "Opportunity Status"]
        _add_mini_table(slide, top_opps, hdrs, c_map, Inches(6.5), Inches(1.55), Inches(6.5))
    else:
        _add_text_box(slide, "No opportunities in pipeline.",
                      Inches(6.5), Inches(2.0), Inches(6.5), Inches(0.4),
                      font_size=12, color=MGRAY)


def _slide_minister_vision(prs, investors, lang):
    """Left: Minister decision log. Right: Vision 2030 + Deal classification."""
    slide = _blank_slide(prs)
    _add_slide_header(slide, "Minister Log & Vision 2030")

    # ── LEFT: Minister items ────────────────────────────────────────
    _add_text_box(slide, "Minister Attention Required",
                  Inches(0.3), Inches(1.1), Inches(6.5), Inches(0.35),
                  font_size=12, bold=True, color=DARK)

    items = []
    if not investors.empty and "Minister Action Required" in investors.columns:
        for _, row in investors.iterrows():
            act = row.get("Minister Action Required", "")
            if act and str(act) not in ("None Required", "nan", ""):
                items.append((
                    row.get("Company Name", "?"),
                    str(act),
                    str(row.get("Strategic Priority Score", "—")),
                    str(row.get("Decision Required By", "—")),
                ))

    if not items:
        _add_text_box(slide, "✓ No items require Minister attention.",
                      Inches(0.3), Inches(1.6), Inches(6.5), Inches(0.4),
                      font_size=11, color=GREEN)
    else:
        _add_rect(slide, Inches(0.3), Inches(1.55), Inches(6.5), Inches(0.3),
                  fill_color=GREEN, line_color=GREEN)
        for j, h in enumerate(["Company", "Action", "Priority", "Due"]):
            ws = [2.0, 2.6, 0.8, 1.0]
            x  = Inches(0.3) + sum(Inches(w) for w in ws[:j])
            _add_text_box(slide, h, x + Inches(0.05), Inches(1.57), Inches(ws[j] - 0.1),
                          Inches(0.26), font_size=8, bold=True, color=WHITE)
        y = Inches(1.95)
        for i, (co, act, pri, due) in enumerate(items[:7]):
            bg = _rgb("#FFF5F5") if i % 2 == 0 else WHITE
            _add_rect(slide, Inches(0.3), y, Inches(6.5), Inches(0.38), fill_color=bg, line_color=_rgb("#DDDDDD"))
            for j, val in enumerate([co[:20], act[:28], pri, str(due)[:10]]):
                ws = [2.0, 2.6, 0.8, 1.0]
                x  = Inches(0.3) + sum(Inches(w) for w in ws[:j])
                _add_text_box(slide, val, x + Inches(0.05), y + Inches(0.05), Inches(ws[j] - 0.1),
                              Inches(0.28), font_size=8, color=DARK)
            y += Inches(0.4)

    # ── RIGHT: Vision 2030 alignment ────────────────────────────────
    _add_text_box(slide, "Vision 2030 Alignment",
                  Inches(7.2), Inches(1.1), Inches(5.8), Inches(0.35),
                  font_size=12, bold=True, color=DARK)

    total_pipeline  = _fmt_sar(_sum_col(investors, "Est. Investment Value (SAR)"))
    total_committed = _fmt_sar(_sum_col(investors, "Actual Commitment (SAR)"))
    total_jobs_raw  = _sum_col(investors, "Est. Jobs Created")
    total_jobs      = f"{int(total_jobs_raw):,}" if total_jobs_raw > 0 else "—"

    for i, (lbl, val, col) in enumerate([
        ("Pipeline",     total_pipeline,  MISA_GREEN),
        ("Committed",    total_committed, MISA_GOLD),
        ("Est. Jobs",    total_jobs,      "#2D7A54"),
    ]):
        x = Inches(7.2) + i * Inches(1.95)
        _add_rect(slide, x, Inches(1.55), Inches(1.85), Inches(0.65), fill_color=_rgb(col), line_color=_rgb(col))
        _add_text_box(slide, val, x, Inches(1.57), Inches(1.85), Inches(0.35),
                      font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _add_text_box(slide, lbl, x, Inches(1.9), Inches(1.85), Inches(0.25),
                      font_size=8, color=WHITE, align=PP_ALIGN.CENTER)

    y = Inches(2.4)
    if not investors.empty and "Vision 2030 Pillar" in investors.columns:
        _add_text_box(slide, "Pillar Breakdown", Inches(7.2), y, Inches(5.8), Inches(0.28),
                      font_size=10, bold=True, color=DARK)
        y += Inches(0.35)
        pillar_counts = investors["Vision 2030 Pillar"].dropna().value_counts()
        total = max(len(investors), 1)
        for pillar, count in pillar_counts.items():
            bar_w = Inches(5.5 * count / total)
            _add_rect(slide, Inches(7.2), y, bar_w, Inches(0.28),
                      fill_color=_rgb(MISA_GOLD), line_color=_rgb(MISA_GOLD))
            _add_text_box(slide, f"{str(pillar)[:26]}: {count}", Inches(7.25), y,
                          Inches(5.5), Inches(0.28), font_size=8, color=WHITE)
            y += Inches(0.35)
            if y > Inches(5.0):
                break

    y += Inches(0.15)
    if not investors.empty and "Deal Classification" in investors.columns:
        _add_text_box(slide, "Deal Types", Inches(7.2), y, Inches(5.8), Inches(0.28),
                      font_size=10, bold=True, color=DARK)
        y += Inches(0.35)
        deal_cols = {
            "Greenfield": MISA_GREEN, "Brownfield / Expansion": MISA_GOLD,
            "Joint Venture": "#2D7A54", "Acquisition": "#E4B96A",
            "Strategic Partnership": "#0F3D2A", "Fund / FDI": "#C9974A",
        }
        for deal, count in investors["Deal Classification"].dropna().value_counts().items():
            col = _rgb(deal_cols.get(deal, MISA_GREEN))
            _add_rect(slide, Inches(7.2), y, Inches(0.22), Inches(0.22),
                      fill_color=col, line_color=col)
            _add_text_box(slide, f"{str(deal)[:28]}: {count}", Inches(7.5), y,
                          Inches(5.5), Inches(0.28), font_size=9, color=DARK)
            y += Inches(0.35)
            if y > Inches(6.8):
                break


def _slide_investor(prs, inv_row, actions, opportunities, lang):
    """Compact per-investor slide — all info on one slide."""
    slide = _blank_slide(prs)
    company = inv_row.get("Company Name", "Unknown")
    tier    = inv_row.get("Investor Tier", "")

    # Green header band
    _add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.9),
              fill_color=GREEN, line_color=GREEN)
    _add_text_box(slide, company, Inches(0.3), Inches(0.07),
                  Inches(8.5), Inches(0.75), font_size=20, bold=True, color=WHITE)
    _add_text_box(slide, tier, Inches(9.2), Inches(0.2),
                  Inches(3.9), Inches(0.5), font_size=11, color=GOLD, align=PP_ALIGN.RIGHT)

    # Metadata row
    meta = [("Country", inv_row.get("Country","—")), ("Sector", inv_row.get("Sector","—")),
            ("Stage", inv_row.get("Journey Stage","—")), ("Status", inv_row.get("Relationship Status","—")),
            ("RM", inv_row.get("Relationship Manager","—")), ("AM", inv_row.get("Account Manager","TBD"))]
    for i, (lbl, val) in enumerate(meta):
        x = Inches(0.3) + i * Inches(2.15)
        _add_text_box(slide, lbl, x, Inches(0.95), Inches(2.1), Inches(0.22),
                      font_size=8, color=MGRAY)
        _add_text_box(slide, str(val)[:22], x, Inches(1.16), Inches(2.1), Inches(0.28),
                      font_size=10, bold=True, color=DARK)

    # Investment + minister strip
    est  = inv_row.get("Est. Investment Value (SAR)")
    cmmt = inv_row.get("Actual Commitment (SAR)")
    min_act = str(inv_row.get("Minister Action Required", "None Required") or "None Required")
    blocker = str(inv_row.get("Blocker Level", "None") or "None")
    nxt = inv_row.get("Next Meeting Date")
    has_flag = min_act not in ("None Required", "nan", "") or blocker not in ("None", "nan", "")

    strip_color = _rgb("#FFF3F0") if has_flag else _rgb("#F0FFF4")
    border_color = RED if has_flag else _rgb(MISA_GREEN)
    _add_rect(slide, Inches(0.3), Inches(1.53), Inches(12.73), Inches(0.35),
              fill_color=strip_color, line_color=border_color)
    strip_text = (
        f"Est: {_fmt_sar(est) if pd.notna(est) and est else '—'}  |  "
        f"Committed: {_fmt_sar(cmmt) if pd.notna(cmmt) and cmmt else '—'}  |  "
        f"Next Meeting: {nxt if pd.notna(nxt) and nxt else '—'}  |  "
        f"Minister Action: {min_act}  |  Blocker: {blocker}"
    )
    _add_text_box(slide, strip_text, Inches(0.35), Inches(1.57),
                  Inches(12.6), Inches(0.28), font_size=8,
                  color=RED if has_flag else _rgb(MISA_GREEN))

    # Opportunities (left)
    _add_text_box(slide, f"Opportunities ({len(opportunities)})",
                  Inches(0.3), Inches(1.98), Inches(6), Inches(0.28),
                  font_size=11, bold=True, color=DARK)
    if not opportunities.empty:
        hdrs  = ["Opportunity", "Stage", "Value", "Status"]
        c_map = ["Opportunity Name", "Opportunity Stage", "Est. Value (SAR)", "Opportunity Status"]
        _add_mini_table(slide, opportunities.head(5), hdrs, c_map, Inches(0.3), Inches(2.32), Inches(6.2))
    else:
        _add_text_box(slide, "No opportunities yet.", Inches(0.3), Inches(2.4), Inches(6), Inches(0.3),
                      font_size=9, color=MGRAY)

    # Action Items (right)
    pending = actions[~actions.get("Status", pd.Series(dtype=str)).isin(["Completed", "Cancelled"])].head(8) if not actions.empty and "Status" in actions.columns else actions.head(8)
    _add_text_box(slide, f"Action Items ({len(pending)} pending)",
                  Inches(6.7), Inches(1.98), Inches(6.3), Inches(0.28),
                  font_size=11, bold=True, color=DARK)
    y = Inches(2.32)
    for _, act in pending.iterrows():
        pri = act.get("Priority", "Medium")
        pc  = {"High": RED, "Medium": _rgb(MISA_GOLD), "Low": GREEN}.get(pri, GREEN)
        _add_rect(slide, Inches(6.7), y + Inches(0.04), Inches(0.08), Inches(0.24),
                  fill_color=pc, line_color=pc)
        desc = str(act.get("Action Description", ""))[:60]
        due  = str(act.get("Due Date", "")) if pd.notna(act.get("Due Date")) else ""
        _add_text_box(slide, f"{desc}  {due}", Inches(6.85), y, Inches(6.1), Inches(0.3),
                      font_size=9, color=DARK)
        y += Inches(0.32)
        if y > Inches(6.5):
            break

    # Notes
    notes = str(inv_row.get("Notes", "") or "")
    if notes:
        _add_text_box(slide, f"Notes: {notes[:140]}", Inches(0.3), Inches(6.65),
                      Inches(12.73), Inches(0.35), font_size=8, color=MGRAY)


# ══════════════════════════════════════════════════════════════════════════════
# COMPANY DECK — slide builders (3 slides)
# ══════════════════════════════════════════════════════════════════════════════

def _co_slide_cover_profile(prs, company, inv_row, opps, acts, meetings, lang):
    """Slide 1: Header + metadata + Brief/Outcome strip + timeline with actions (L) + status (R)."""
    slide = _blank_slide(prs)

    # ── Green header band ────────────────────────────────────────────
    _add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.82),
              fill_color=GREEN, line_color=GREEN)
    _add_text_box(slide, company,
                  Inches(0.3), Inches(0.07), Inches(9.0), Inches(0.68),
                  font_size=26, bold=True, color=WHITE)
    tier = str(inv_row.get("Investor Tier", "") or "") if not inv_row.empty else ""
    _add_text_box(slide, tier, Inches(8.8), Inches(0.14),
                  Inches(4.2), Inches(0.35), font_size=12, color=GOLD, align=PP_ALIGN.RIGHT)
    _add_text_box(slide,
                  f"Investor Status Report — {date.today().strftime('%d %B %Y')}",
                  Inches(8.8), Inches(0.49), Inches(4.2), Inches(0.26),
                  font_size=8, color=_rgb("CCCCCC"), align=PP_ALIGN.RIGHT)

    # ── White metadata band ──────────────────────────────────────────
    _add_rect(slide, Inches(0), Inches(0.82), Inches(13.33), Inches(0.58),
              fill_color=WHITE, line_color=WHITE)
    if not inv_row.empty:
        for i, (lbl, val) in enumerate([
            ("Country",       str(inv_row.get("Country",  "—") or "—")),
            ("Sector",        str(inv_row.get("Sector",   "—") or "—")),
            ("Journey Stage", str(inv_row.get("Journey Stage", "—") or "—")),
            ("Status",        str(inv_row.get("Relationship Status", "—") or "—")),
            ("RM",            str(inv_row.get("Relationship Manager", "—") or "—")),
            ("AM",            str(inv_row.get("Account Manager", "TBD") or "TBD")),
        ]):
            x = Inches(0.4) + i * Inches(2.1)
            _add_text_box(slide, lbl, x, Inches(0.87), Inches(2.0), Inches(0.18),
                          font_size=8, color=MGRAY)
            _add_text_box(slide, val[:22], x, Inches(1.04), Inches(2.0), Inches(0.3),
                          font_size=10, bold=True, color=DARK)

    # Gold accent line under metadata
    _add_rect(slide, Inches(0), Inches(1.40), Inches(13.33), Inches(0.025),
              fill_color=GOLD, line_color=GOLD)

    # ── Brief + Major Outcome strip ──────────────────────────────────
    BRIEF_Y = Inches(1.46)
    BRIEF_H = Inches(0.66)
    _add_rect(slide, Inches(0), BRIEF_Y, Inches(13.33), BRIEF_H,
              fill_color=_rgb("#F5F5F0"), line_color=_rgb("#E4E4DC"))

    brief_text    = "—"
    major_outcome = "—"
    if not meetings.empty and "Meeting Date" in meetings.columns:
        last_mtg  = meetings.sort_values("Meeting Date", ascending=False).iloc[0]
        obj  = str(last_mtg.get("Meeting Objective",     "") or "")
        kd   = str(last_mtg.get("Key Discussion Points", "") or "")
        dm   = str(last_mtg.get("Decisions Made",        "") or "")
        ns   = str(last_mtg.get("Next Steps",            "") or "")
        brief_text    = (obj or kd or "—")[:135]
        major_outcome = (dm or ns or "—")[:135]

    # Fall back to completed actions for major outcome
    if major_outcome in ("—", "") and not acts.empty and "Status" in acts.columns:
        done = acts[acts["Status"].str.lower().str.contains("complet", na=False)]
        if not done.empty:
            major_outcome = "✓ " + str(done.iloc[-1].get("Action Description", "") or "")[:120]

    _add_rect(slide, Inches(6.58), BRIEF_Y + Inches(0.08), Inches(0.02),
              BRIEF_H - Inches(0.16), fill_color=_rgb("#CCCCCC"), line_color=_rgb("#CCCCCC"))

    _add_text_box(slide, "Brief", Inches(0.3), BRIEF_Y + Inches(0.06),
                  Inches(0.62), Inches(0.2), font_size=7, bold=True, color=_rgb(MISA_GREEN))
    _add_text_box(slide, brief_text, Inches(0.95), BRIEF_Y + Inches(0.05),
                  Inches(5.52), Inches(0.52), font_size=7.5, color=DARK)

    _add_text_box(slide, "Major Outcome", Inches(6.68), BRIEF_Y + Inches(0.06),
                  Inches(1.35), Inches(0.2), font_size=7, bold=True, color=_rgb(MISA_GOLD))
    _add_text_box(slide, major_outcome, Inches(8.1), BRIEF_Y + Inches(0.05),
                  Inches(5.05), Inches(0.52), font_size=7.5, color=DARK)

    # ── Vertical divider (timeline | status panels) ──────────────────
    DIVX = Inches(8.85)
    _add_rect(slide, DIVX, Inches(2.17), Inches(0.02), Inches(4.83),
              fill_color=_rgb("#DDDDDD"), line_color=_rgb("#DDDDDD"))

    # ══════════════════════════════════════════════════════════════════
    # LEFT: Engagement Timeline — meetings above, action items below
    # ══════════════════════════════════════════════════════════════════
    _add_text_box(slide, "Engagement Timeline",
                  Inches(0.3), Inches(2.20), Inches(8.3), Inches(0.28),
                  font_size=10, bold=True, color=DARK)

    today_d  = date.today()
    start_dt = today_d - timedelta(days=120)
    end_dt   = today_d + timedelta(days=45)
    tot_days = (end_dt - start_dt).days

    AXIS_Y = Inches(4.1)
    AXIS_L = Inches(0.55)
    AXIS_W = Inches(7.95)
    DOT_R  = Inches(0.13)
    SQ     = Inches(0.12)

    def _dx(d):
        try:
            if hasattr(d, "date"):
                d = d.date()
            elif not isinstance(d, type(today_d)):
                d = pd.to_datetime(d).date()
            frac = max(0.0, min(1.0, (d - start_dt).days / tot_days))
            return AXIS_L + frac * AXIS_W
        except Exception:
            return None

    # Month tick marks + labels
    cur = start_dt.replace(day=1)
    while cur <= end_dt:
        mx = _dx(cur)
        if mx is not None and mx >= AXIS_L:
            _add_rect(slide, mx, AXIS_Y - Inches(0.12), Inches(0.015), Inches(0.14),
                      fill_color=_rgb("#AAAAAA"), line_color=_rgb("#AAAAAA"))
            _add_text_box(
                slide,
                f"{_calendar.month_abbr[cur.month]} '{cur.year % 100:02d}",
                mx - Inches(0.27), AXIS_Y - Inches(0.35), Inches(0.72), Inches(0.22),
                font_size=7, color=MGRAY, align=PP_ALIGN.CENTER,
            )
        if cur.month == 12:
            cur = cur.replace(year=cur.year + 1, month=1)
        else:
            cur = cur.replace(month=cur.month + 1)

    # Main axis line
    _add_rect(slide, AXIS_L, AXIS_Y, AXIS_W, Inches(0.025),
              fill_color=_rgb("#CCCCCC"), line_color=_rgb("#CCCCCC"))

    # "Now" marker
    now_x = _dx(today_d)
    if now_x:
        _add_rect(slide, now_x - Inches(0.01), AXIS_Y - Inches(0.30), Inches(0.02), Inches(0.38),
                  fill_color=GOLD, line_color=GOLD)
        _add_text_box(slide, "Now", now_x - Inches(0.22), AXIS_Y + Inches(0.06),
                      Inches(0.5), Inches(0.18), font_size=7, color=GOLD, align=PP_ALIGN.CENTER)

    # Meetings ABOVE axis — all above, alternating high/low heights to avoid overlap
    if not meetings.empty and "Meeting Date" in meetings.columns:
        recent_mtgs = meetings.sort_values("Meeting Date").tail(10).reset_index(drop=True)
        for idx, (_, mtg) in enumerate(recent_mtgs.iterrows()):
            raw_date = mtg.get("Meeting Date")
            if pd.isna(raw_date):
                continue
            mx = _dx(raw_date)
            if mx is None:
                continue

            status = str(mtg.get("Meeting Status", ""))
            if any(w in status for w in ("Completed", "Done", "Held")):
                dot_c = _rgb(MISA_GREEN)
            elif any(w in status for w in ("Scheduled", "Upcoming", "Planned")):
                dot_c = _rgb(MISA_GOLD)
            elif any(w in status for w in ("Cancelled", "Canceled")):
                dot_c = _rgb("#AAAAAA")
            else:
                dot_c = _rgb(MISA_GREEN)

            mtype = str(mtg.get("Meeting Type", ""))[:14]
            try:
                d_lbl = pd.to_datetime(raw_date).strftime("%d %b")
            except Exception:
                d_lbl = str(raw_date)[:8]

            # Even → high position, odd → lower position (still above axis)
            if idx % 2 == 0:
                text_y   = Inches(2.52)
                conn_top = Inches(2.94)
                conn_h   = AXIS_Y - DOT_R - Inches(2.94)
            else:
                text_y   = Inches(3.12)
                conn_top = Inches(3.54)
                conn_h   = AXIS_Y - DOT_R - Inches(3.54)

            _add_rect(slide, mx - Inches(0.01), conn_top, Inches(0.02), conn_h,
                      fill_color=_rgb("#DDDDDD"), line_color=_rgb("#DDDDDD"))
            _add_rect(slide, mx - DOT_R, AXIS_Y - DOT_R, DOT_R * 2, DOT_R * 2,
                      fill_color=dot_c, line_color=dot_c)
            _add_text_box(
                slide, f"{d_lbl}\n{mtype}",
                mx - Inches(0.55), text_y, Inches(1.1), Inches(0.40),
                font_size=7, color=DARK, align=PP_ALIGN.CENTER,
            )

    # Action Items BELOW axis — squares colored by status, plotted by due date
    _ACT_SC = {
        "Completed":   MISA_GREEN,
        "In Progress": MISA_GOLD,
        "Inprogress":  MISA_GOLD,
        "Not Started": "#888888",
        "Blocked":     "#C0392B",
        "Cancelled":   "#AAAAAA",
    }
    if not acts.empty and "Due Date" in acts.columns:
        act_tl = acts.copy()
        act_tl["_dt"] = pd.to_datetime(act_tl["Due Date"], errors="coerce")
        act_tl = act_tl[act_tl["_dt"].notna()].sort_values("_dt").tail(9)
        ROW_H = Inches(0.40)
        for ai, (_, act) in enumerate(act_tl.iterrows()):
            ax = _dx(act["_dt"])
            if ax is None:
                continue
            row_idx = ai % 3
            sq_top  = AXIS_Y + Inches(0.14) + row_idx * ROW_H
            scol    = _rgb(_ACT_SC.get(str(act.get("Status", "")), "#888888"))
            desc    = str(act.get("Action Description", ""))[:16]
            try:
                d_lbl = act["_dt"].strftime("%d %b")
            except Exception:
                d_lbl = ""
            # Connector from axis down to square
            _add_rect(slide, ax - Inches(0.008), AXIS_Y + Inches(0.025),
                      Inches(0.015), Inches(0.14) + row_idx * ROW_H,
                      fill_color=_rgb("#DDDDDD"), line_color=_rgb("#DDDDDD"))
            # Square marker
            _add_rect(slide, ax - SQ / 2, sq_top, SQ, SQ,
                      fill_color=scol, line_color=scol)
            # Tiny label below square
            _add_text_box(
                slide, f"{desc}\n{d_lbl}",
                ax - Inches(0.52), sq_top + SQ + Inches(0.01), Inches(1.04), Inches(0.34),
                font_size=6, color=DARK, align=PP_ALIGN.CENTER,
            )

    # Timeline legend
    leg_y = Inches(6.40)
    lx = Inches(0.55)
    _add_text_box(slide, "Meetings:", lx, leg_y, Inches(0.72), Inches(0.2),
                  font_size=7, bold=True, color=MGRAY)
    lx += Inches(0.75)
    for lbl, col in [("Completed", MISA_GREEN), ("Scheduled", MISA_GOLD), ("Cancelled", "#AAAAAA")]:
        _add_rect(slide, lx, leg_y + Inches(0.03), Inches(0.13), Inches(0.13),
                  fill_color=_rgb(col), line_color=_rgb(col))
        _add_text_box(slide, lbl, lx + Inches(0.17), leg_y,
                      Inches(0.94), Inches(0.2), font_size=7, color=MGRAY)
        lx += Inches(1.12)
    lx += Inches(0.18)
    _add_text_box(slide, "Actions:", lx, leg_y, Inches(0.65), Inches(0.2),
                  font_size=7, bold=True, color=MGRAY)
    lx += Inches(0.68)
    for lbl, col in [("Done", MISA_GREEN), ("In Progress", MISA_GOLD),
                     ("Pending", "#888888"), ("Blocked", "#C0392B")]:
        _add_rect(slide, lx, leg_y + Inches(0.04), Inches(0.11), Inches(0.11),
                  fill_color=_rgb(col), line_color=_rgb(col))
        _add_text_box(slide, lbl, lx + Inches(0.15), leg_y,
                      Inches(0.84), Inches(0.2), font_size=7, color=MGRAY)
        lx += Inches(0.98)

    # ══════════════════════════════════════════════════════════════════
    # RIGHT: Action status bars + Upcoming actions + Opportunities + KPIs
    # ══════════════════════════════════════════════════════════════════
    RX = DIVX + Inches(0.2)
    RW = Inches(13.33) - RX - Inches(0.15)

    # Action Items status bars
    _add_text_box(slide, "Action Items",
                  RX, Inches(2.22), RW, Inches(0.28),
                  font_size=10, bold=True, color=DARK)
    y_r = Inches(2.58)
    ACT_CFG = [
        ("Completed",   "#2D7A54"),
        ("In Progress", MISA_GOLD),
        ("Not Started", "#888888"),
        ("Blocked",     "#C0392B"),
        ("Cancelled",   "#AAAAAA"),
    ]
    if not acts.empty and "Status" in acts.columns:
        sc      = acts["Status"].value_counts()
        total_a = max(len(acts), 1)
        for sname, scol in ACT_CFG:
            cnt = int(sc.get(sname, 0))
            if cnt == 0:
                continue
            bw = max(Inches(0.15), RW * cnt / total_a)
            _add_rect(slide, RX, y_r, bw, Inches(0.26),
                      fill_color=_rgb(scol), line_color=_rgb(scol))
            _add_text_box(slide, f"{sname}: {cnt}",
                          RX + Inches(0.06), y_r + Inches(0.04),
                          RW - Inches(0.1), Inches(0.18), font_size=8, color=WHITE)
            y_r += Inches(0.31)
    else:
        _add_text_box(slide, "No action items yet.",
                      RX, y_r, RW, Inches(0.25), font_size=8, color=MGRAY)
        y_r += Inches(0.32)

    # Upcoming / overdue actions mini-list
    y_r += Inches(0.12)
    if not acts.empty and "Status" in acts.columns:
        pending_acts = acts[~acts["Status"].isin(["Completed", "Cancelled"])].copy()
        if "Due Date" in pending_acts.columns:
            pending_acts["_due"] = pd.to_datetime(pending_acts["Due Date"], errors="coerce")
            pending_acts = pending_acts.sort_values("_due").head(4)
        else:
            pending_acts = pending_acts.head(4)
        if not pending_acts.empty:
            _add_text_box(slide, "Upcoming Actions",
                          RX, y_r, RW, Inches(0.22), font_size=8, bold=True, color=DARK)
            y_r += Inches(0.26)
            for _, pa in pending_acts.iterrows():
                desc = str(pa.get("Action Description", ""))[:34]
                due  = pa.get("Due Date", "")
                try:
                    due_d  = pd.to_datetime(due)
                    due_s  = due_d.strftime("%d %b") if pd.notna(due) else "—"
                    is_ov  = due_d.date() < today_d
                except Exception:
                    due_s, is_ov = "—", False
                bullet   = "⚠" if is_ov else "•"
                txt_color = RED if is_ov else DARK
                _add_text_box(slide, f"{bullet} {desc}  ({due_s})",
                              RX, y_r, RW, Inches(0.22), font_size=7.5, color=txt_color)
                y_r += Inches(0.24)

    # Opportunities status
    y_r += Inches(0.12)
    total_val = _sum_col(opps, "Est. Value (SAR)")
    _add_text_box(slide, "Opportunities",
                  RX, y_r, RW, Inches(0.28), font_size=10, bold=True, color=DARK)
    y_r += Inches(0.34)
    if not opps.empty:
        if "Opportunity Status" in opps.columns:
            for sname, cnt in opps["Opportunity Status"].value_counts().items():
                _add_rect(slide, RX, y_r + Inches(0.04), Inches(0.13), Inches(0.13),
                          fill_color=GREEN, line_color=GREEN)
                _add_text_box(slide, f"{sname}: {cnt}",
                              RX + Inches(0.2), y_r, RW - Inches(0.2), Inches(0.22),
                              font_size=8, color=DARK)
                y_r += Inches(0.27)
        _add_text_box(slide, f"Pipeline: {_fmt_sar(total_val)}",
                      RX, y_r + Inches(0.06), RW, Inches(0.24),
                      font_size=9, bold=True, color=GOLD)
    else:
        _add_text_box(slide, "No opportunities yet.",
                      RX, y_r, RW, Inches(0.25), font_size=8, color=MGRAY)

    # KPI cards fixed at bottom of right panel
    if not inv_row.empty:
        est    = inv_row.get("Est. Investment Value (SAR)")
        cmmt   = inv_row.get("Actual Commitment (SAR)")
        next_m = str(inv_row.get("Next Meeting Date", "—") or "—")[:12]
        kpi_y  = Inches(5.72)
        kw     = (RW - Inches(0.06)) / 2

        for i, (val, lbl, col) in enumerate([
            (_fmt_sar(est)  if pd.notna(est)  and est  else "—", "Est. Investment", MISA_GREEN),
            (_fmt_sar(cmmt) if pd.notna(cmmt) and cmmt else "—", "Committed",       MISA_GOLD),
        ]):
            kx = RX + i * (kw + Inches(0.06))
            _add_rect(slide, kx, kpi_y, kw, Inches(0.72),
                      fill_color=_rgb(col), line_color=_rgb(col))
            _add_text_box(slide, val, kx, kpi_y + Inches(0.04), kw, Inches(0.34),
                          font_size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            _add_text_box(slide, lbl, kx, kpi_y + Inches(0.41), kw, Inches(0.2),
                          font_size=7, color=WHITE, align=PP_ALIGN.CENTER)

        _add_rect(slide, RX, kpi_y + Inches(0.79), RW, Inches(0.52),
                  fill_color=_rgb("#2D7A54"), line_color=_rgb("#2D7A54"))
        _add_text_box(slide, next_m, RX, kpi_y + Inches(0.83), RW, Inches(0.28),
                      font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _add_text_box(slide, "Next Meeting", RX, kpi_y + Inches(1.09), RW, Inches(0.16),
                      font_size=7, color=WHITE, align=PP_ALIGN.CENTER)

    # ── Gold footer ──────────────────────────────────────────────────
    _add_rect(slide, Inches(0), Inches(7.05), Inches(13.33), Inches(0.45),
              fill_color=GOLD, line_color=GOLD)
    _add_text_box(slide, "CONFIDENTIAL | Ministry of Investment — وزارة الاستثمار",
                  Inches(0), Inches(7.05), Inches(13.33), Inches(0.45),
                  font_size=10, color=WHITE, align=PP_ALIGN.CENTER)


def _co_slide_meetings_opps(prs, company, meetings, opps, lang):
    """Meetings (left) + Opportunities (right)."""
    slide = _blank_slide(prs)
    _add_slide_header(slide, f"{company} — Meetings & Opportunities")

    # LEFT: Meetings
    _add_text_box(slide, f"Meeting Log ({len(meetings)})",
                  Inches(0.3), Inches(1.1), Inches(6.2), Inches(0.32),
                  font_size=12, bold=True, color=DARK)
    if meetings.empty:
        _add_text_box(slide, "No meetings logged yet.",
                      Inches(0.3), Inches(1.55), Inches(6.2), Inches(0.4),
                      font_size=11, color=MGRAY)
    else:
        hdrs  = ["Date", "Type", "Status", "Objective", "Next Steps"]
        c_map = ["Meeting Date", "Meeting Type", "Meeting Status", "Meeting Objective", "Next Steps"]
        data  = meetings.sort_values("Meeting Date", ascending=False).head(8) if "Meeting Date" in meetings.columns else meetings.head(8)
        _add_mini_table(slide, data, hdrs, c_map, Inches(0.3), Inches(1.52), Inches(6.2))

    # RIGHT: Opportunities
    total_val = _sum_col(opps, "Est. Value (SAR)")
    _add_text_box(slide, f"Opportunities ({len(opps)})  —  Pipeline: {_fmt_sar(total_val)}",
                  Inches(6.8), Inches(1.1), Inches(6.2), Inches(0.32),
                  font_size=12, bold=True, color=DARK)
    if opps.empty:
        _add_text_box(slide, "No opportunities in pipeline.",
                      Inches(6.8), Inches(1.55), Inches(6.2), Inches(0.4),
                      font_size=11, color=MGRAY)
    else:
        hdrs  = ["Opportunity", "Stage", "Value", "Confidence", "Status"]
        c_map = ["Opportunity Name", "Opportunity Stage", "Est. Value (SAR)", "Confidence Level", "Opportunity Status"]
        _add_mini_table(slide, opps.head(10), hdrs, c_map, Inches(6.8), Inches(1.52), Inches(6.2))

    # Divider
    _add_rect(slide, Inches(6.6), Inches(1.1), Inches(0.02), Inches(5.8),
              fill_color=_rgb("#DDDDDD"), line_color=_rgb("#DDDDDD"))


def _co_slide_actions_deals(prs, company, actions, deals, lang):
    """Action Items (left) + Deal Progress (right)."""
    slide = _blank_slide(prs)
    _add_slide_header(slide, f"{company} — Action Items & Deal Progress")

    # LEFT: Action Items
    pending = actions[~actions.get("Status", pd.Series(dtype=str)).isin(["Completed","Cancelled"])].head(12) if not actions.empty and "Status" in actions.columns else actions.head(12)
    _add_text_box(slide, f"Action Items ({len(pending)} pending)",
                  Inches(0.3), Inches(1.1), Inches(6.2), Inches(0.32),
                  font_size=12, bold=True, color=DARK)
    if pending.empty:
        _add_text_box(slide, "No pending action items.", Inches(0.3), Inches(1.55),
                      Inches(6.2), Inches(0.4), font_size=11, color=MGRAY)
    else:
        hdrs  = ["Action", "Assigned To", "Priority", "Due", "Status"]
        c_map = ["Action Description", "Assigned To", "Priority", "Due Date", "Status"]
        _add_mini_table(slide, pending, hdrs, c_map, Inches(0.3), Inches(1.52), Inches(6.2))

    # RIGHT: Deal Progress
    _add_text_box(slide, f"Deal Progress ({len(deals)})",
                  Inches(6.8), Inches(1.1), Inches(6.2), Inches(0.32),
                  font_size=12, bold=True, color=DARK)
    if deals.empty:
        _add_text_box(slide, "No deals in progress.", Inches(6.8), Inches(1.55),
                      Inches(6.2), Inches(0.4), font_size=11, color=MGRAY)
    else:
        hdrs  = ["Deal", "Stage", "Status", "Severity", "Escalation"]
        c_map = ["Deal Name", "Deal Stage", "Deal Status", "Challenge Severity", "Escalation Level"]
        _add_mini_table(slide, deals.head(10), hdrs, c_map, Inches(6.8), Inches(1.52), Inches(6.2))

    # Divider
    _add_rect(slide, Inches(6.6), Inches(1.1), Inches(0.02), Inches(5.8),
              fill_color=_rgb("#DDDDDD"), line_color=_rgb("#DDDDDD"))


# ══════════════════════════════════════════════════════════════════════════════
# Low-level helpers
# ══════════════════════════════════════════════════════════════════════════════

def _blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _fill_background(slide, color: RGBColor):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_slide_header(slide, title: str):
    _add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(1.0),
              fill_color=GREEN, line_color=GREEN)
    _add_text_box(slide, title, Inches(0.3), Inches(0.12),
                  Inches(10), Inches(0.76), font_size=20, bold=True, color=WHITE)
    _add_text_box(slide, date.today().strftime("%d %b %Y"),
                  Inches(10.5), Inches(0.25), Inches(2.5), Inches(0.5),
                  font_size=11, color=GOLD, align=PP_ALIGN.RIGHT)


def _add_text_box(slide, text, left, top, width, height,
                  font_size=12, bold=False, color=None, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text) if text is not None else ""
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    return tb


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


def _add_mini_table(slide, df, headers, cols, left, top, width):
    if df.empty:
        return
    available = [c for c in cols if c in df.columns]
    if not available:
        return
    mapped_headers = [headers[cols.index(c)] for c in available]
    col_w = width / max(len(available), 1)
    row_h = Inches(0.32)

    # Header row
    for j, hdr in enumerate(mapped_headers):
        x = left + j * col_w
        _add_rect(slide, x, top, col_w, row_h, fill_color=GREEN, line_color=GREEN)
        _add_text_box(slide, hdr, x + Inches(0.04), top + Inches(0.04),
                      col_w - Inches(0.08), row_h - Inches(0.05),
                      font_size=8, bold=True, color=WHITE)

    # Data rows
    for i, (_, row) in enumerate(df.iterrows()):
        y  = top + (i + 1) * row_h
        bg = _rgb("#F7F7F2") if i % 2 == 0 else WHITE
        for j, col in enumerate(available):
            x   = left + j * col_w
            val = row.get(col, "")
            if not isinstance(val, str) and pd.isna(val):
                val = "—"
            if isinstance(val, float) and "Value" in col or "SAR" in col:
                val = _fmt_sar(val)
            _add_rect(slide, x, y, col_w, row_h, fill_color=bg, line_color=_rgb("#E0E0E0"))
            _add_text_box(slide, str(val)[:28], x + Inches(0.04), y + Inches(0.04),
                          col_w - Inches(0.08), row_h - Inches(0.05),
                          font_size=7, color=DARK)


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


def _sum_col(df, col):
    if df.empty or col not in df.columns:
        return 0.0
    return pd.to_numeric(df[col], errors="coerce").fillna(0).sum()
