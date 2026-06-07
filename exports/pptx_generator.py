
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
    deals         = dfs.get("Deal Progress",        pd.DataFrame())

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
            inv_acts = actions[actions["Company Name"] == co]       if not actions.empty       and "Company Name" in actions.columns       else pd.DataFrame()
            inv_opps = opportunities[opportunities["Company Name"] == co] if not opportunities.empty and "Company Name" in opportunities.columns else pd.DataFrame()
            inv_dls  = deals[deals["Company Name"] == co]           if not deals.empty         and "Company Name" in deals.columns         else pd.DataFrame()
            _slide_investor(prs, inv, inv_acts, inv_opps, inv_dls, lang)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def generate_pptx_company(dfs: dict, company: str, lang: str = "en") -> bytes:
    """
    Single-company deck — 2 slides.
    Slide 1: Cover + Timeline + Full Action Items table
    Slide 2: Opportunities (main) + Deal Progress (merged tracker)
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
    inv_mtgs = meetings[meetings["Company Name"] == company]           if not meetings.empty      and "Company Name" in meetings.columns      else pd.DataFrame()
    inv_opps = opportunities[opportunities["Company Name"] == company] if not opportunities.empty and "Company Name" in opportunities.columns else pd.DataFrame()
    inv_acts = actions[actions["Company Name"] == company]             if not actions.empty       and "Company Name" in actions.columns       else pd.DataFrame()
    inv_dls  = deals[deals["Company Name"] == company]                 if not deals.empty         and "Company Name" in deals.columns         else pd.DataFrame()

    _co_slide_cover_profile(prs, company, inv_row, inv_opps, inv_acts, inv_mtgs, inv_dls, lang)
    _co_slide_opps_deals(prs, company, inv_row, inv_opps, inv_dls, inv_acts, lang)

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


def _slide_investor(prs, inv_row, actions, opportunities, deals, lang):
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
    meta = [
        ("Country", _mv(inv_row, "Country")),
        ("Sector",  _mv(inv_row, "Sector")),
        ("Stage",   _mv(inv_row, "Journey Stage")),
        ("Status",  _mv(inv_row, "Relationship Status")),
        ("RM",      _mv(inv_row, "Relationship Manager")),
        ("AM",      _mv(inv_row, "Account Manager", "TBD")),
    ]
    for i, (lbl, val) in enumerate(meta):
        x = Inches(0.3) + i * Inches(2.15)
        _add_text_box(slide, lbl, x, Inches(0.95), Inches(2.1), Inches(0.22),
                      font_size=8, color=MGRAY)
        _add_text_box(slide, val[:22], x, Inches(1.16), Inches(2.1), Inches(0.28),
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

    # Opportunities (left upper)
    opp_val = _sum_col(opportunities, "Est. Value (SAR)")
    _add_text_box(slide,
                  f"Opportunities ({len(opportunities)})   Pipeline: {_fmt_sar(opp_val)}",
                  Inches(0.3), Inches(1.98), Inches(6.2), Inches(0.28),
                  font_size=11, bold=True, color=DARK)
    if not opportunities.empty:
        hdrs  = ["Opportunity", "Stage", "Value", "Status"]
        c_map = ["Opportunity Name", "Opportunity Stage", "Est. Value (SAR)", "Opportunity Status"]
        _add_mini_table(slide, opportunities.head(4), hdrs, c_map, Inches(0.3), Inches(2.32), Inches(6.2))
    else:
        _add_text_box(slide, "No opportunities yet.", Inches(0.3), Inches(2.4), Inches(6), Inches(0.3),
                      font_size=9, color=MGRAY)

    # Deal Progress (left lower)
    deals_y = Inches(3.88)
    n_blocked = int((deals["Deal Status"] == "Blocked").sum()) if not deals.empty and "Deal Status" in deals.columns else 0
    _DS_COL = {"Critical": "#C0392B", "High": MISA_GOLD, "Medium": "#888888", "Low": MISA_GREEN}
    _add_text_box(slide, f"Deal Progress ({len(deals)})  |  Blocked: {n_blocked}",
                  Inches(0.3), deals_y, Inches(6.2), Inches(0.24),
                  font_size=10, bold=True, color=RED if n_blocked else DARK)
    if not deals.empty:
        dy = deals_y + Inches(0.28)
        for _, dl in deals.head(4).iterrows():
            dname = str(dl.get("Deal Name",  "") or "")[:32]
            dstg  = str(dl.get("Deal Stage", "") or "")[:14]
            dsev  = str(dl.get("Challenge Severity", "") or "")
            dc    = _rgb(_DS_COL.get(dsev, "#888888"))
            _add_rect(slide, Inches(0.3), dy + Inches(0.03), Inches(0.08), Inches(0.08),
                      fill_color=dc, line_color=dc)
            _add_text_box(slide, dname, Inches(0.44), dy,
                          Inches(4.2), Inches(0.18), font_size=8, color=DARK)
            _add_text_box(slide, f"{dstg}  ·  {dsev}", Inches(4.7), dy,
                          Inches(1.5), Inches(0.18), font_size=7.5, color=MGRAY)
            dy += Inches(0.30)
    else:
        _add_text_box(slide, "No deals in progress.", Inches(0.3), deals_y + Inches(0.28), Inches(6.2), Inches(0.25),
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

def _co_slide_cover_profile(prs, company, inv_row, opps, acts, meetings, deals, lang):
    """Slide 1 (merged 1+3): Header + metadata + Brief strip + compact timeline + full action items table (L) + status/deals/KPI (R)."""
    slide = _blank_slide(prs)

    # ── Green header band ────────────────────────────────────────────
    _add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.82),
              fill_color=GREEN, line_color=GREEN)
    # Try to show company logo in header
    _s1_name_x = Inches(0.3)
    if not inv_row.empty:
        _s1_logo_b = _fetch_logo_bytes(company, _mv(inv_row, "Website", ""))
        if _s1_logo_b:
            try:
                slide.shapes.add_picture(
                    io.BytesIO(_s1_logo_b), Inches(0.22), Inches(0.14), Inches(0.52), Inches(0.52))
                _s1_name_x = Inches(0.86)
            except Exception:
                pass
    _add_text_box(slide, company,
                  _s1_name_x, Inches(0.07), Inches(8.3), Inches(0.68),
                  font_size=26, bold=True, color=WHITE)
    _add_text_box(slide,
                  f"Investor Status Report — {date.today().strftime('%d %B %Y')}",
                  Inches(8.8), Inches(0.30), Inches(4.2), Inches(0.40),
                  font_size=8, color=_rgb("CCCCCC"), align=PP_ALIGN.RIGHT)

    # ── White metadata band ──────────────────────────────────────────
    _add_rect(slide, Inches(0), Inches(0.82), Inches(13.33), Inches(0.58),
              fill_color=WHITE, line_color=WHITE)
    if not inv_row.empty:
        for i, (lbl, val) in enumerate([
            ("Country",       _mv(inv_row, "Country")),
            ("Sector",        _mv(inv_row, "Sector")),
            ("Journey Stage", _mv(inv_row, "Journey Stage")),
            ("Status",        _mv(inv_row, "Relationship Status")),
            ("RM",            _mv(inv_row, "Relationship Manager")),
            ("AM",            _mv(inv_row, "Account Manager", "TBD")),
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
    BRIEF_H = Inches(0.90)
    _add_rect(slide, Inches(0), BRIEF_Y, Inches(13.33), BRIEF_H,
              fill_color=_rgb("#F5F5F0"), line_color=_rgb("#E4E4DC"))

    # ── Strategic brief — synthesised from all action items ────────────────────
    goal_text, strat_text, action_text = _build_strategic_brief(acts, opps, meetings, inv_row)

    # Dividers — split brief strip into 3 panels
    _add_rect(slide, Inches(6.40), BRIEF_Y + Inches(0.08), Inches(0.02),
              BRIEF_H - Inches(0.16), fill_color=_rgb("#CCCCCC"), line_color=_rgb("#CCCCCC"))
    _add_rect(slide, Inches(10.10), BRIEF_Y + Inches(0.08), Inches(0.02),
              BRIEF_H - Inches(0.16), fill_color=_rgb("#CCCCCC"), line_color=_rgb("#CCCCCC"))

    # Panel 1: Strategic Goal  (label inline left, text fills rest)
    _add_text_box(slide, "STRATEGIC GOAL", Inches(0.3), BRIEF_Y + Inches(0.07),
                  Inches(1.16), Inches(0.22), font_size=9, bold=True, color=_rgb(MISA_GREEN))
    _add_text_box(slide, goal_text, Inches(1.50), BRIEF_Y + Inches(0.06),
                  Inches(4.78), Inches(0.74), font_size=9, color=DARK)

    # Panel 2: Ministry Strategy
    _add_text_box(slide, "MINISTRY STRATEGY", Inches(6.50), BRIEF_Y + Inches(0.07),
                  Inches(3.45), Inches(0.22), font_size=9, bold=True, color=_rgb(MISA_GOLD))
    _add_text_box(slide, strat_text, Inches(6.50), BRIEF_Y + Inches(0.32),
                  Inches(3.45), Inches(0.52), font_size=9, color=DARK)

    # Panel 3: Minister Action
    _add_text_box(slide, "MINISTER ACTION", Inches(10.20), BRIEF_Y + Inches(0.07),
                  Inches(2.93), Inches(0.22), font_size=9, bold=True, color=RED)
    _add_text_box(slide, action_text, Inches(10.20), BRIEF_Y + Inches(0.32),
                  Inches(2.93), Inches(0.52), font_size=9, color=DARK)

    # ── Vertical divider ────────────────────────────────────────────────────────
    DIVX = Inches(10.2)
    _add_rect(slide, DIVX, Inches(2.41), Inches(0.02), Inches(4.59),
              fill_color=_rgb("#DDDDDD"), line_color=_rgb("#DDDDDD"))

    # ══════════════════════════════════════════════════════════════════════════
    # LEFT: Compact timeline + Full action items table
    # ══════════════════════════════════════════════════════════════════════════
    today_d  = date.today()
    start_dt = today_d - timedelta(days=120)
    end_dt   = today_d + timedelta(days=45)
    tot_days = (end_dt - start_dt).days

    AXIS_Y = Inches(2.84)
    AXIS_L = Inches(0.55)
    AXIS_W = Inches(9.20)
    DOT_R  = Inches(0.09)
    SQ     = Inches(0.10)

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

    _add_text_box(slide, "Engagement Timeline",
                  Inches(0.3), Inches(2.41), Inches(9.8), Inches(0.20),
                  font_size=9, bold=True, color=DARK)

    # Month ticks
    cur = start_dt.replace(day=1)
    while cur <= end_dt:
        mx = _dx(cur)
        if mx is not None and mx >= AXIS_L:
            _add_rect(slide, mx, AXIS_Y - Inches(0.08), Inches(0.012), Inches(0.10),
                      fill_color=_rgb("#BBBBBB"), line_color=_rgb("#BBBBBB"))
            _add_text_box(
                slide, _calendar.month_abbr[cur.month],
                mx - Inches(0.20), AXIS_Y - Inches(0.27), Inches(0.42), Inches(0.17),
                font_size=6, color=MGRAY, align=PP_ALIGN.CENTER,
            )
        if cur.month == 12:
            cur = cur.replace(year=cur.year + 1, month=1)
        else:
            cur = cur.replace(month=cur.month + 1)

    # Axis line
    _add_rect(slide, AXIS_L, AXIS_Y, AXIS_W, Inches(0.022),
              fill_color=_rgb("#CCCCCC"), line_color=_rgb("#CCCCCC"))

    # "Now" marker
    now_x = _dx(today_d)
    if now_x:
        _add_rect(slide, now_x - Inches(0.008), AXIS_Y - Inches(0.18), Inches(0.016), Inches(0.22),
                  fill_color=GOLD, line_color=GOLD)
        _add_text_box(slide, "Now", now_x - Inches(0.18), AXIS_Y - Inches(0.30),
                      Inches(0.38), Inches(0.13), font_size=6, color=GOLD, align=PP_ALIGN.CENTER)

    # Meeting dots ON axis — color by status, tiny date label above
    _MTG_SC = {
        "completed": MISA_GREEN, "done": MISA_GREEN, "held": MISA_GREEN,
        "scheduled": MISA_GOLD,  "upcoming": MISA_GOLD, "planned": MISA_GOLD,
        "cancelled": "#AAAAAA",  "canceled": "#AAAAAA",
    }
    if not meetings.empty and "Meeting Date" in meetings.columns:
        recent_mtgs = meetings.sort_values("Meeting Date").tail(12).reset_index(drop=True)
        for idx, (_, mtg) in enumerate(recent_mtgs.iterrows()):
            raw_date = mtg.get("Meeting Date")
            if pd.isna(raw_date):
                continue
            mx = _dx(raw_date)
            if mx is None:
                continue
            status_lc = str(mtg.get("Meeting Status", "")).lower()
            dot_c = _rgb(next((v for k, v in _MTG_SC.items() if k in status_lc), MISA_GREEN))
            try:
                d_lbl = pd.to_datetime(raw_date).strftime("%d %b")
            except Exception:
                d_lbl = str(raw_date)[:8]
            _add_rect(slide, mx - DOT_R, AXIS_Y - DOT_R, DOT_R * 2, DOT_R * 2,
                      fill_color=dot_c, line_color=dot_c)
            lbl_y = Inches(2.33) if idx % 2 == 0 else Inches(2.41)
            _add_text_box(slide, d_lbl, mx - Inches(0.22), lbl_y,
                          Inches(0.45), Inches(0.13), font_size=6, color=DARK, align=PP_ALIGN.CENTER)

    # Action item due-date markers BELOW axis
    _ACT_SC = {
        "Completed":   MISA_GREEN, "In Progress": MISA_GOLD, "Inprogress": MISA_GOLD,
        "Not Started": "#888888",  "Blocked": "#C0392B",      "Cancelled": "#AAAAAA",
    }
    if not acts.empty and "Due Date" in acts.columns:
        act_tl = acts.copy()
        act_tl["_dt"] = pd.to_datetime(act_tl["Due Date"], errors="coerce")
        act_tl = act_tl[act_tl["_dt"].notna()].sort_values("_dt")
        for ai, (_, act) in enumerate(act_tl.iterrows()):
            ax = _dx(act["_dt"])
            if ax is None:
                continue
            sq_y = AXIS_Y + Inches(0.04) + (ai % 2) * Inches(0.14)
            scol = _rgb(_ACT_SC.get(str(act.get("Status", "")), "#888888"))
            _add_rect(slide, ax - SQ / 2, sq_y, SQ, SQ * 0.85,
                      fill_color=scol, line_color=scol)

    # Compact legend
    leg_y = Inches(3.09)
    lx = Inches(0.55)
    for lbl, col in [("Completed mtg", MISA_GREEN), ("Scheduled", MISA_GOLD), ("Cancelled", "#AAAAAA")]:
        _add_rect(slide, lx, leg_y + Inches(0.02), Inches(0.11), Inches(0.11),
                  fill_color=_rgb(col), line_color=_rgb(col))
        _add_text_box(slide, lbl, lx + Inches(0.14), leg_y,
                      Inches(0.95), Inches(0.16), font_size=6, color=MGRAY)
        lx += Inches(1.10)
    lx += Inches(0.1)
    for lbl, col in [("Done", MISA_GREEN), ("In Progress", MISA_GOLD),
                     ("Pending", "#888888"), ("Blocked", "#C0392B")]:
        _add_rect(slide, lx, leg_y + Inches(0.03), Inches(0.09), Inches(0.09),
                  fill_color=_rgb(col), line_color=_rgb(col))
        _add_text_box(slide, lbl, lx + Inches(0.12), leg_y,
                      Inches(0.80), Inches(0.16), font_size=6, color=MGRAY)
        lx += Inches(0.92)

    # ── Action Items full table ────────────────────────────────────────────────
    n_total   = len(acts)
    n_pending = len(acts[~acts["Status"].isin(["Completed", "Cancelled"])]) if not acts.empty and "Status" in acts.columns else 0
    _add_text_box(slide, f"Action Items  —  {n_total} total  |  {n_pending} pending",
                  Inches(0.3), Inches(3.28), Inches(9.75), Inches(0.20),
                  font_size=9, bold=True, color=DARK)

    TBL_L = Inches(0.3)
    HDR_Y = Inches(3.52)
    HDR_H = Inches(0.30)
    ROW_H = Inches(0.32)
    CW    = [Inches(w) for w in [0.40, 4.85, 1.65, 1.20, 0.80, 0.70]]
    CX    = [TBL_L + sum(CW[:j]) for j in range(len(CW))]
    HDRS  = ["#", "Action Item", "Assigned To", "Status", "Due", "Progress"]

    for hdr, cx, cw in zip(HDRS, CX, CW):
        _add_rect(slide, cx, HDR_Y, cw, HDR_H, fill_color=GREEN, line_color=GREEN)
        _add_text_box(slide, hdr, cx + Inches(0.03), HDR_Y + Inches(0.04),
                      cw - Inches(0.06), HDR_H - Inches(0.06),
                      font_size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Sort: pending by due date first, then completed
    if not acts.empty and "Status" in acts.columns:
        pend_df = acts[~acts["Status"].isin(["Completed", "Cancelled"])].copy()
        done_df = acts[acts["Status"].isin(["Completed", "Cancelled"])].copy()
        if "Due Date" in pend_df.columns:
            pend_df["_due"] = pd.to_datetime(pend_df["Due Date"], errors="coerce")
            pend_df = pend_df.sort_values("_due")
        display_acts = pd.concat([pend_df, done_df], ignore_index=True).head(10)
    else:
        display_acts = acts.head(10) if not acts.empty else pd.DataFrame()

    _SBG = {
        "Completed":   "#E8F5E9", "In Progress": "#FFF8E1", "Inprogress": "#FFF8E1",
        "Not Started": "#F5F5F5", "Blocked":     "#FFEBEE", "Cancelled":  "#EEEEEE",
    }
    _SFG = {
        "Completed":   "#2D7A54", "In Progress": "#B45309", "Inprogress": "#B45309",
        "Not Started": "#666666", "Blocked":     "#C0392B", "Cancelled":  "#999999",
    }

    for i, (_, row) in enumerate(display_acts.iterrows()):
        ry  = HDR_Y + HDR_H + i * ROW_H
        alt = _rgb("#F7F7F2") if i % 2 == 0 else WHITE

        desc   = str(row.get("Action Description", "") or "")
        owner  = str(row.get("Assigned To",        "") or "")[:20]
        status = str(row.get("Status", "Not Started") or "Not Started")
        due    = row.get("Due Date", "")
        remark = str(row.get("Remarks", "") or "")[:42]

        try:
            due_d = pd.to_datetime(due)
            due_s = due_d.strftime("%d %b %y") if pd.notna(due) else "—"
            is_ov = due_d.date() < today_d and status not in ("Completed", "Cancelled")
        except Exception:
            due_s, is_ov = "—", False

        try:
            prog_raw = row.get("Progress", 0)
            prog_val = float(str(prog_raw).replace("%", "")) if prog_raw not in (None, "", "nan") else 0.0
            if prog_val > 1.0:
                prog_val /= 100.0
        except Exception:
            prog_val = 0.0

        sbg = _rgb(_SBG.get(status, "#F5F5F5"))
        sfg = _rgb(_SFG.get(status, "#666666"))

        for j, (val, cx, cw) in enumerate(zip(
            [str(i + 1), desc, owner, status, due_s],
            CX[:5], CW[:5]
        )):
            bg = sbg if j == 3 else (_rgb("#FFEBEE") if (j == 4 and is_ov) else alt)
            _add_rect(slide, cx, ry, cw, ROW_H, fill_color=bg, line_color=_rgb("#DDDDDD"))
            fc = sfg if j == 3 else (RED if (j == 4 and is_ov) else DARK)
            al = PP_ALIGN.CENTER if j in (0, 3, 4) else PP_ALIGN.LEFT
            if j == 1:
                _add_text_box(slide, desc[:58], cx + Inches(0.04), ry + Inches(0.02),
                              cw - Inches(0.08), Inches(0.15), font_size=7.5, color=DARK)
                if remark:
                    _add_text_box(slide, f"↳ {remark}", cx + Inches(0.05), ry + Inches(0.17),
                                  cw - Inches(0.10), Inches(0.12), font_size=6, color=MGRAY)
            else:
                _add_text_box(slide, val, cx + Inches(0.03), ry + Inches(0.05),
                              cw - Inches(0.06), Inches(0.20),
                              font_size=8 if j == 0 else 7.5, color=fc,
                              bold=(j == 3), align=al)

        # Progress bar + % label
        px, pcw = CX[5], CW[5]
        _add_rect(slide, px, ry, pcw, ROW_H, fill_color=alt, line_color=_rgb("#DDDDDD"))
        bx = px + Inches(0.07)
        by = ry + Inches(0.07)
        bw = pcw - Inches(0.14)
        bh = Inches(0.09)
        _add_rect(slide, bx, by, bw, bh, fill_color=_rgb("#E0E0E0"), line_color=_rgb("#E0E0E0"))
        if prog_val > 0:
            fc2 = _rgb(MISA_GREEN) if prog_val >= 1.0 else (_rgb(MISA_GOLD) if prog_val >= 0.5 else _rgb("#888888"))
            _add_rect(slide, bx, by, max(bw * prog_val, Inches(0.02)), bh,
                      fill_color=fc2, line_color=fc2)
        _add_text_box(slide, f"{int(prog_val * 100)}%", px, ry + Inches(0.18),
                      pcw, Inches(0.12), font_size=6.5, color=DARK, align=PP_ALIGN.CENTER)

    # ══════════════════════════════════════════════════════════════════════════
    # RIGHT: Status summary + Deals + Opportunities + KPI cards
    # ══════════════════════════════════════════════════════════════════════════
    RX = DIVX + Inches(0.2)
    RW = Inches(13.33) - RX - Inches(0.15)

    # Action status bars
    _add_text_box(slide, "Action Items", RX, Inches(2.46), RW, Inches(0.24),
                  font_size=9, bold=True, color=DARK)
    y_r = Inches(2.76)
    for sname, scol in [("Completed","#2D7A54"), ("In Progress",MISA_GOLD),
                        ("Not Started","#888888"), ("Blocked","#C0392B"), ("Cancelled","#AAAAAA")]:
        cnt = int(acts["Status"].value_counts().get(sname, 0)) if not acts.empty and "Status" in acts.columns else 0
        if cnt == 0:
            continue
        bw = max(Inches(0.12), RW * cnt / max(len(acts), 1))
        _add_rect(slide, RX, y_r, bw, Inches(0.24), fill_color=_rgb(scol), line_color=_rgb(scol))
        _add_text_box(slide, f"{sname}: {cnt}",
                      RX + Inches(0.05), y_r + Inches(0.04),
                      RW - Inches(0.08), Inches(0.16), font_size=7.5, color=WHITE)
        y_r += Inches(0.28)
    if acts.empty:
        _add_text_box(slide, "No action items.", RX, y_r, RW, Inches(0.22), font_size=8, color=MGRAY)
        y_r += Inches(0.28)

    # Deal Progress
    y_r += Inches(0.14)
    _add_text_box(slide, f"Deal Progress ({len(deals)})",
                  RX, y_r, RW, Inches(0.24), font_size=9, bold=True, color=DARK)
    y_r += Inches(0.30)
    _DS = {"Completed":MISA_GREEN, "In Progress":MISA_GOLD, "Active":MISA_GOLD,
           "On Track":MISA_GREEN,  "Blocked":"#C0392B",     "Delayed":"#E67E22"}
    if not deals.empty:
        for _, dl in deals.head(4).iterrows():
            dname  = str(dl.get("Deal Name",  "") or "")[:28]
            dstage = str(dl.get("Deal Stage", "") or "")[:18]
            dstat  = str(dl.get("Deal Status","") or "")
            dc     = _rgb(_DS.get(dstat, "#888888"))
            _add_rect(slide, RX, y_r + Inches(0.04), Inches(0.09), Inches(0.09),
                      fill_color=dc, line_color=dc)
            _add_text_box(slide, dname, RX + Inches(0.13), y_r,
                          RW - Inches(0.13), Inches(0.18), font_size=7.5, color=DARK)
            _add_text_box(slide, f"{dstage}  •  {dstat}", RX + Inches(0.13), y_r + Inches(0.18),
                          RW - Inches(0.13), Inches(0.14), font_size=6.5, color=MGRAY)
            y_r += Inches(0.36)
    else:
        _add_text_box(slide, "No deals in progress.", RX, y_r, RW, Inches(0.22), font_size=8, color=MGRAY)
        y_r += Inches(0.28)

    # Opportunities quick summary
    y_r += Inches(0.12)
    total_val = _sum_col(opps, "Est. Value (SAR)")
    _add_text_box(slide, f"Opportunities ({len(opps)})",
                  RX, y_r, RW, Inches(0.24), font_size=9, bold=True, color=DARK)
    y_r += Inches(0.28)
    if not opps.empty and "Opportunity Status" in opps.columns:
        for sname, cnt in opps["Opportunity Status"].value_counts().items():
            _add_rect(slide, RX, y_r + Inches(0.04), Inches(0.10), Inches(0.10),
                      fill_color=GREEN, line_color=GREEN)
            _add_text_box(slide, f"{sname}: {cnt}",
                          RX + Inches(0.16), y_r, RW - Inches(0.16), Inches(0.20),
                          font_size=8, color=DARK)
            y_r += Inches(0.24)
        _add_text_box(slide, f"Pipeline: {_fmt_sar(total_val)}",
                      RX, y_r + Inches(0.04), RW, Inches(0.20),
                      font_size=8, bold=True, color=GOLD)
    else:
        _add_text_box(slide, "No opportunities yet.", RX, y_r, RW, Inches(0.20), font_size=8, color=MGRAY)

    # KPI cards at bottom
    if not inv_row.empty:
        est    = inv_row.get("Est. Investment Value (SAR)")
        cmmt   = inv_row.get("Actual Commitment (SAR)")
        next_m = str(inv_row.get("Next Meeting Date", "—") or "—")[:12]
        kpi_y  = Inches(5.78)
        kw     = (RW - Inches(0.06)) / 2

        for i, (val, lbl, col) in enumerate([
            (_fmt_sar(est)  if pd.notna(est)  and est  else "—", "Est. Investment", MISA_GREEN),
            (_fmt_sar(cmmt) if pd.notna(cmmt) and cmmt else "—", "Committed",       MISA_GOLD),
        ]):
            kx = RX + i * (kw + Inches(0.06))
            _add_rect(slide, kx, kpi_y, kw, Inches(0.68),
                      fill_color=_rgb(col), line_color=_rgb(col))
            _add_text_box(slide, val, kx, kpi_y + Inches(0.04), kw, Inches(0.32),
                          font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            _add_text_box(slide, lbl, kx, kpi_y + Inches(0.38), kw, Inches(0.18),
                          font_size=7, color=WHITE, align=PP_ALIGN.CENTER)

        _add_rect(slide, RX, kpi_y + Inches(0.75), RW, Inches(0.50),
                  fill_color=_rgb("#2D7A54"), line_color=_rgb("#2D7A54"))
        _add_text_box(slide, next_m, RX, kpi_y + Inches(0.78), RW, Inches(0.28),
                      font_size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _add_text_box(slide, "Next Meeting", RX, kpi_y + Inches(1.04), RW, Inches(0.16),
                      font_size=7, color=WHITE, align=PP_ALIGN.CENTER)

    # ── Gold footer ───────────────────────────────────────────────────────────
    _add_rect(slide, Inches(0), Inches(7.05), Inches(13.33), Inches(0.45),
              fill_color=GOLD, line_color=GOLD)
    _add_text_box(slide, "CONFIDENTIAL | Ministry of Investment — وزارة الاستثمار",
                  Inches(0), Inches(7.05), Inches(13.33), Inches(0.45),
                  font_size=10, color=WHITE, align=PP_ALIGN.CENTER)


def _co_slide_opps_deals(prs, company, inv_row, opps, deals, acts, lang):
    """Slide 2 — Minister-grade opportunity cards, full-width 2-column grid."""
    slide = _blank_slide(prs)

    # ── Green header band
    _add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.9),
              fill_color=GREEN, line_color=GREEN)
    _s2_name_x = Inches(0.3)
    if not inv_row.empty:
        _s2_logo_b = _fetch_logo_bytes(company, _mv(inv_row, "Website", ""))
        if _s2_logo_b:
            try:
                slide.shapes.add_picture(
                    io.BytesIO(_s2_logo_b), Inches(0.22), Inches(0.14), Inches(0.52), Inches(0.52))
                _s2_name_x = Inches(0.86)
            except Exception:
                pass
    _add_text_box(slide, f"{company} — Investment Opportunities",
                  _s2_name_x, Inches(0.08), Inches(10.0), Inches(0.74),
                  font_size=20, bold=True, color=WHITE)
    _add_text_box(slide, date.today().strftime("%d %b %Y"),
                  Inches(10.5), Inches(0.25), Inches(2.5), Inches(0.5),
                  font_size=11, color=GOLD, align=PP_ALIGN.RIGHT)

    # ── Resolve what to show as cards ─────────────────────────────────────────
    # If no formal opportunities exist, use Type="Opportunity" action items
    use_act_cards = opps.empty and not acts.empty
    if use_act_cards:
        opp_acts = pd.DataFrame()
        if "Type of Engagement" in acts.columns:
            opp_acts = acts[acts["Type of Engagement"].fillna("").str.lower().str.contains("opport", regex=False)]
        card_source = opp_acts if not opp_acts.empty else acts
        n_cards  = len(card_source)
        n_active = int(card_source["Status"].isin(["Inprogress","In Progress","Not Started"]).sum()) if "Status" in card_source.columns else 0
        n_done   = int(card_source["Status"].str.lower().str.contains("complet", na=False).sum())   if "Status" in card_source.columns else 0
        total_val = 0.0
    else:
        card_source = opps
        n_cards     = len(opps)
        n_active    = int((opps["Opportunity Status"] == "Active").sum())    if not opps.empty and "Opportunity Status" in opps.columns else 0
        n_done      = int((opps["Opportunity Stage"]  == "Committed").sum()) if not opps.empty and "Opportunity Stage"  in opps.columns else 0
        total_val   = _sum_col(opps, "Est. Value (SAR)")

    # ── KPI strip (4 cards across full width)
    kw = Inches(3.13)
    if use_act_cards:
        kpi_data = [
            (str(n_cards),  "Engagement Actions", MISA_GREEN),
            (str(n_active), "Active / Pending",   MISA_GREEN),
            (str(n_done),   "Completed",           MISA_GOLD),
            ("—",           "Pipeline Value",      MISA_GREEN),
        ]
    else:
        kpi_data = [
            (str(n_cards),         "Total Opportunities", MISA_GREEN),
            (str(n_active),        "Active",              MISA_GREEN),
            (str(n_done),          "Committed",           MISA_GOLD),
            (_fmt_sar(total_val),  "Pipeline Value",      MISA_GREEN),
        ]
    for i, (val, lbl, col) in enumerate(kpi_data):
        kx = Inches(0.3) + i * (kw + Inches(0.08))
        _add_rect(slide, kx, Inches(0.97), kw, Inches(0.56),
                  fill_color=_rgb(col), line_color=_rgb(col))
        _add_text_box(slide, val, kx, Inches(0.99), kw, Inches(0.30),
                      font_size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _add_text_box(slide, lbl, kx, Inches(1.29), kw, Inches(0.18),
                      font_size=7.5, color=WHITE, align=PP_ALIGN.CENTER)

    # ── Card constants ─────────────────────────────────────────────────────────
    _STAGE_COL = {
        "Committed":   "#065F46", "Negotiation":         "#92400E",
        "Exploration": "#1D4ED8", "Active":              "#1B5C3F",
        "Suspended":   "#6B7280", "Blocked":             "#991B1B",
        "On Track":    "#1B5C3F", "Opportunity Matching":"#1D4ED8",
    }
    _STAGE_LIGHT = {
        "Committed":   "#D1FAE5", "Negotiation":         "#FEF3C7",
        "Exploration": "#DBEAFE", "Active":              "#D1FAE5",
        "Suspended":   "#F3F4F6", "Blocked":             "#FEE2E2",
        "On Track":    "#D1FAE5", "Opportunity Matching":"#DBEAFE",
    }
    _STATUS_BG = {"Active":"#D1FAE5","Inactive":"#F3F4F6","On Hold":"#FEF3C7","Completed":"#D1FAE5",
                  "Inprogress":"#FEF3C7","In Progress":"#FEF3C7","Not Started":"#F3F4F6",
                  "Blocked":"#FEE2E2","Cancelled":"#EEEEEE"}
    _STATUS_FG = {"Active":"#065F46","Inactive":"#374151","On Hold":"#92400E","Completed":"#065F46",
                  "Inprogress":"#92400E","In Progress":"#92400E","Not Started":"#374151",
                  "Blocked":"#991B1B","Cancelled":"#6B7280"}
    _PRI_COL   = {"High":"#991B1B","Very High":"#7F1D1D","Medium":"#1B5C3F","Low":"#6B7280"}
    _PRI_LIGHT = {"High":"#FEE2E2","Very High":"#FEE2E2","Medium":"#D1FAE5","Low":"#F3F4F6"}

    CARD_W  = Inches(6.28)
    CARD_H  = Inches(1.65)
    GAP_X   = Inches(0.13)
    GAP_Y   = Inches(0.12)
    COL_X   = [Inches(0.24), Inches(0.24) + CARD_W + GAP_X]
    START_Y = Inches(1.73)

    if card_source.empty:
        _add_text_box(slide,
                      "No opportunities or action items have been recorded for this investor.",
                      Inches(0.3), Inches(2.5), Inches(12.73), Inches(0.5),
                      font_size=13, color=MGRAY, align=PP_ALIGN.CENTER)

    elif use_act_cards:
        # ── Action items rendered as engagement cards ──────────────────────────
        _ACT_DOT = {
            "Completed":   MISA_GREEN, "Inprogress":  MISA_GOLD,
            "In Progress": MISA_GOLD,  "Not Started": "#AAAAAA",
            "Blocked":     "#C0392B",  "Cancelled":   "#CCCCCC",
        }
        for idx, (_, act) in enumerate(card_source.head(6).iterrows()):
            row_i = idx // 2
            col_i = idx % 2
            cx = COL_X[col_i]
            cy = START_Y + row_i * (CARD_H + GAP_Y)

            act_desc = str(act.get("Action Description", "") or "").strip() or "Action Item"
            act_rem  = str(act.get("Remarks",            "") or "").strip()
            act_stat = str(act.get("Status",             "") or "").strip()
            act_pri  = str(act.get("Priority",           "") or "").strip()
            act_due  = act.get("Due Date", None)
            act_own  = str(act.get("Assigned To",        "") or "").strip()

            accent_hex = _PRI_COL.get(act_pri,   "#1B5C3F")
            light_hex  = _PRI_LIGHT.get(act_pri, "#D1FAE5")

            _add_rect(slide, cx, cy, CARD_W, CARD_H,
                      fill_color=WHITE, line_color=_rgb("#D1D5DB"))
            _add_rect(slide, cx, cy, Inches(0.09), CARD_H,
                      fill_color=_rgb(accent_hex), line_color=_rgb(accent_hex))
            _add_rect(slide, cx + Inches(0.09), cy, CARD_W - Inches(0.09), Inches(0.04),
                      fill_color=_rgb(light_hex), line_color=_rgb(light_hex))

            IX = cx + Inches(0.16)
            IW = CARD_W - Inches(0.24)

            # Title: action description
            _add_text_box(slide, act_desc[:68],
                          IX, cy + Inches(0.08), IW, Inches(0.26),
                          font_size=10, bold=True, color=GREEN)

            # Pills: status + priority
            pill_y = cy + Inches(0.37)
            pill_x = IX
            for pill_txt, pbg, pfg in [
                (act_stat, _STATUS_BG.get(act_stat, "#F3F4F6"), _STATUS_FG.get(act_stat, "#374151")),
                (act_pri,  _PRI_LIGHT.get(act_pri,  "#F3F4F6"), _PRI_COL.get(act_pri,   "#374151")),
            ]:
                if pill_txt and pill_txt not in ("—", ""):
                    pw = Inches(min(1.72, max(0.70, len(pill_txt) * 0.072)))
                    _add_rect(slide, pill_x, pill_y, pw, Inches(0.21),
                              fill_color=_rgb(pbg), line_color=_rgb(pbg))
                    _add_text_box(slide, pill_txt, pill_x + Inches(0.04), pill_y + Inches(0.02),
                                  pw - Inches(0.08), Inches(0.17),
                                  font_size=7.5, bold=True, color=_rgb(pfg))
                    pill_x += pw + Inches(0.06)

            # Remark / update text (main body)
            if act_rem and act_rem not in ("nan", "Key notes", ""):
                _add_text_box(slide, act_rem[:120],
                              IX, cy + Inches(0.62), IW, Inches(0.50),
                              font_size=8, color=DARK)
            else:
                _add_text_box(slide, "No update recorded.",
                              IX, cy + Inches(0.62), IW, Inches(0.22),
                              font_size=8, color=MGRAY)

            # Footer: owner + due date
            foot_parts = []
            if act_own and act_own not in ("nan",):
                foot_parts.append(f"Owner: {act_own[:30]}")
            if act_due is not None and not (isinstance(act_due, float) and pd.isna(act_due)):
                try:
                    foot_parts.append(f"Due: {pd.to_datetime(act_due).strftime('%d %b %Y')}")
                except Exception:
                    pass
            if foot_parts:
                _add_text_box(slide, "  |  ".join(foot_parts),
                              IX, cy + CARD_H - Inches(0.20), IW, Inches(0.18),
                              font_size=6.5, color=MGRAY)

        if len(card_source) > 6:
            _add_text_box(slide,
                          f"Showing 6 of {len(card_source)} action items — remaining visible on slide 1",
                          Inches(0.24), Inches(7.0), Inches(12.73), Inches(0.18),
                          font_size=7, color=MGRAY, align=PP_ALIGN.CENTER)

    else:
        # ── Formal opportunity cards ───────────────────────────────────────────
        _CONF_BG = {"High":"#D1FAE5","Medium":"#FEF3C7","Low":"#FEE2E2"}
        _CONF_FG = {"High":"#065F46","Medium":"#92400E","Low":"#991B1B"}
        for idx, (_, opp) in enumerate(card_source.head(6).iterrows()):
            row_i = idx // 2
            col_i = idx % 2
            cx    = COL_X[col_i]
            cy    = START_Y + row_i * (CARD_H + GAP_Y)

            opp_name   = (str(opp.get("Opportunity Name",   "") or "")).strip() or "Unnamed Opportunity"
            stage      = (str(opp.get("Opportunity Stage",  "") or "")).strip()
            status     = (str(opp.get("Opportunity Status", "") or "")).strip()
            confidence = (str(opp.get("Confidence Level",   "") or "")).strip()
            sector     = (str(opp.get("Sector",             "") or "")).strip()
            est_val    = opp.get("Est. Value (SAR)")
            start_d    = opp.get("Start Date")
            due_d      = opp.get("Due Date")
            notes      = (str(opp.get("Notes", "") or opp.get("Remarks", "") or "")).strip()
            opp_type   = (str(opp.get("Opportunity Type", "") or "")).strip()

            accent_hex = _STAGE_COL.get(stage, "#1B5C3F")
            light_hex  = _STAGE_LIGHT.get(stage, "#D1FAE5")

            # Card background + outer border
            _add_rect(slide, cx, cy, CARD_W, CARD_H,
                      fill_color=WHITE, line_color=_rgb("#D1D5DB"))
            # Left accent bar
            _add_rect(slide, cx, cy, Inches(0.09), CARD_H,
                      fill_color=_rgb(accent_hex), line_color=_rgb(accent_hex))
            # Top light stripe
            _add_rect(slide, cx + Inches(0.09), cy, CARD_W - Inches(0.09), Inches(0.04),
                      fill_color=_rgb(light_hex), line_color=_rgb(light_hex))

            IX = cx + Inches(0.16)
            IW = CARD_W - Inches(0.24)

            # Row 1: Opportunity name — large, bold, green
            _add_text_box(slide, opp_name[:62],
                          IX, cy + Inches(0.08), IW, Inches(0.26),
                          font_size=11, bold=True, color=GREEN)

            # Row 2: Stage / Status / Confidence pills
            pill_y = cy + Inches(0.37)
            pill_x = IX
            pills = []
            if stage:
                pills.append((stage[:16],      accent_hex,                              "#FFFFFF"))
            if status and status not in ("—", ""):
                pills.append((status,          _STATUS_BG.get(status, "#F3F4F6"),       _STATUS_FG.get(status, "#374151")))
            if confidence and confidence not in ("—", ""):
                pills.append((f"{confidence} *", _CONF_BG.get(confidence, "#F3F4F6"),  _CONF_FG.get(confidence, "#374151")))
            for pill_txt, pbg, pfg in pills:
                pw = Inches(min(1.72, max(0.70, len(pill_txt) * 0.072)))
                _add_rect(slide, pill_x, pill_y, pw, Inches(0.21),
                          fill_color=_rgb(pbg), line_color=_rgb(pbg))
                _add_text_box(slide, pill_txt.strip(), pill_x + Inches(0.04), pill_y + Inches(0.02),
                              pw - Inches(0.08), Inches(0.17),
                              font_size=7.5, bold=True, color=_rgb(pfg))
                pill_x += pw + Inches(0.06)

            # Rows 3+: Action history list matched by sector/keyword
            _ACT_DOT = {
                "Completed":   MISA_GREEN, "Inprogress":  MISA_GOLD,
                "In Progress": MISA_GOLD,  "Not Started": "#AAAAAA",
                "Blocked":     "#C0392B",  "Cancelled":   "#CCCCCC",
            }
            matched_acts = _match_acts_to_opp(opp_name, acts)
            act_y = cy + Inches(0.62)
            ACT_ROW = Inches(0.32)
            shown_acts = 0
            for _, ar in matched_acts.head(3).iterrows():
                if act_y + ACT_ROW > cy + CARD_H - Inches(0.04):
                    break
                ad  = str(ar.get("Action Description", "") or "").strip()
                rem = str(ar.get("Remarks",            "") or "").strip()
                st  = str(ar.get("Status",             "") or "").strip()
                if not ad or ad in ("nan",):
                    continue
                dot_col = _rgb(_ACT_DOT.get(st, "#AAAAAA"))
                _add_rect(slide, IX, act_y + Inches(0.04), Inches(0.07), Inches(0.07),
                          fill_color=dot_col, line_color=dot_col)
                _add_text_box(slide, ad[:72], IX + Inches(0.11), act_y,
                              IW - Inches(0.11), Inches(0.17),
                              font_size=7.5, color=DARK)
                if rem and rem not in ("nan", "Key notes", ""):
                    _add_text_box(slide, f"↳ {rem[:72]}", IX + Inches(0.14), act_y + Inches(0.17),
                                  IW - Inches(0.14), Inches(0.14),
                                  font_size=6.5, color=MGRAY)
                act_y += ACT_ROW
                shown_acts += 1

            if shown_acts == 0:
                # No matched actions — show stage context
                stage_map = {
                    "Committed": "secured and committed", "Negotiation": "in active negotiation",
                    "Exploration": "in early-stage exploration", "Active": "actively progressing",
                    "Opportunity Matching": "being matched to MISA priorities",
                }
                stage_desc = stage_map.get(stage, "under development")
                _add_text_box(slide, f"Engagement {stage_desc}. No actions logged yet.",
                              IX, cy + Inches(0.62), IW, Inches(0.24),
                              font_size=8, color=MGRAY)

            # Summary footer inside card
            if not matched_acts.empty and "Status" in matched_acts.columns:
                n_tot  = len(matched_acts)
                n_done = int(matched_acts["Status"].str.lower().str.contains("complet").sum())
                _add_text_box(slide,
                              f"{n_done}/{n_tot} completed",
                              IX, cy + CARD_H - Inches(0.18), IW, Inches(0.16),
                              font_size=6.5, color=MGRAY)

    # ── Gold footer
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


def _match_acts_to_opp(opp_name: str, acts) -> "pd.DataFrame":
    """Return action items related to an opportunity.
    Matches first by Sector column, then by keywords in Action Description."""
    if acts is None or acts.empty:
        return pd.DataFrame()
    keywords = [w for w in opp_name.lower().split() if len(w) > 3]
    if not keywords:
        return pd.DataFrame()
    pattern = "|".join(keywords[:5])
    # Sector match is most reliable (e.g. Sector="Health" → "Enter Health" opp)
    if "Sector" in acts.columns:
        try:
            mask = acts["Sector"].fillna("").str.lower().str.contains(pattern, regex=True)
            result = acts[mask]
            if not result.empty:
                return result
        except Exception:
            pass
    # Fall back: keyword in action description
    if "Action Description" in acts.columns:
        try:
            mask = acts["Action Description"].fillna("").str.lower().str.contains(pattern, regex=True)
            result = acts[mask]
            if not result.empty:
                return result
        except Exception:
            pass
    return pd.DataFrame()


def _fetch_logo_bytes(company: str, website: str = ""):
    """Fetch company logo for PPTX embedding. Returns bytes or None."""
    import urllib.request, urllib.parse
    urls = []
    if website:
        domain = website.replace("https://", "").replace("http://", "").split("/")[0].strip()
        if domain:
            urls.append(f"https://logo.clearbit.com/{domain}")
            urls.append(f"https://www.google.com/s2/favicons?domain={domain}&sz=64")
    if not urls:
        slug = urllib.parse.quote(company.lower().replace(" ", ""))
        urls.append(f"https://logo.clearbit.com/{slug}.com")
    for url in urls:
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=3) as resp:
                data = resp.read()
                if data and len(data) > 200:
                    return data
        except Exception:
            continue
    return None


def _build_opportunity_narrative(opp, acts=None, card_idx: int = 0) -> tuple:
    """Return (description, impact) narrative strings for an opportunity card.
    Description: built from opp name + sector + stage (never from action items,
    so each card is distinct). Impact: pulled from the best matching action
    item remark; falls back to round-robin by card_idx so cards differ."""
    opp_name   = str(opp.get("Opportunity Name",   "") or "").strip()
    stage      = str(opp.get("Opportunity Stage",  "") or "").strip()
    sector     = str(opp.get("Sector",             "") or "").strip()
    confidence = str(opp.get("Confidence Level",   "") or "").strip()
    notes      = str(opp.get("Notes", "") or opp.get("Remarks", "") or "").strip()
    est_val    = opp.get("Est. Value (SAR)")
    blockers   = str(opp.get("Blockers", "") or "").strip()

    # ── Description: from the opportunity itself, not action items ────────────
    stage_map = {
        "Committed":           "secured and committed",
        "Negotiation":         "in active negotiation",
        "Exploration":         "in early-stage exploration",
        "Active":              "actively progressing",
        "Opportunity Matching":"being matched to MISA priorities",
        "Suspended":           "temporarily suspended",
        "Blocked":             "blocked — escalation required",
        "On Track":            "on track for closure",
    }
    stage_desc = stage_map.get(stage, f"at {stage} stage" if stage else "under development")
    sect_part  = f" in {sector}" if sector and sector not in ("—", "") else ""
    desc = f"{opp_name}{sect_part} — {stage_desc}."
    if notes and notes not in ("nan",) and len(notes) > 5:
        desc = f"{desc} {notes[:55]}"

    # ── Impact: best matching action item remark ───────────────────────────────
    impact_remark = ""
    if acts is not None and not acts.empty:
        # Try to find an action item whose description mentions this opp's keywords
        keywords = [w for w in opp_name.lower().split() if len(w) > 3]
        matched = pd.DataFrame()
        if keywords and "Action Description" in acts.columns:
            try:
                pattern = "|".join(keywords[:4])
                mask = acts["Action Description"].fillna("").str.lower().str.contains(pattern, regex=True)
                matched = acts[mask]
            except Exception:
                pass

        # Pick remark: from matched set if found, else round-robin across all
        pool = matched if not matched.empty else acts
        if "Remarks" in pool.columns:
            pool_rem = pool[pool["Remarks"].fillna("").str.strip().str.len() > 5].reset_index(drop=True)
            if not pool_rem.empty:
                pick = card_idx % len(pool_rem)
                impact_remark = str(pool_rem.iloc[pick]["Remarks"]).strip()

    # Build impact line
    impact_parts = []
    if impact_remark and impact_remark not in ("nan",):
        impact_parts.append(impact_remark[:80])
    has_val = est_val is not None and not (isinstance(est_val, float) and pd.isna(est_val))
    if has_val:
        impact_parts.append(f"Est. value: {_fmt_sar(est_val)}")
    if not impact_parts:
        conf_map = {"High": "High closure confidence", "Medium": "Medium confidence", "Low": "Early probability"}
        if confidence in conf_map:
            impact_parts.append(conf_map[confidence])
        if blockers and blockers not in ("nan", "—", ""):
            impact_parts.append(f"Blocker: {blockers[:40]}")
        if not impact_parts:
            impact_parts.append("Engagement ongoing — outcome to be determined")

    impact = "; ".join(impact_parts)
    if not impact.endswith("."):
        impact += "."

    return desc[:105], impact[:90]


def _build_strategic_brief(acts, opps, meetings, inv_row):
    """
    Synthesise a minister-grade strategic brief from all action items.
    Returns (goal_text, ministry_strategy, minister_action).
    """
    company     = _mv(inv_row, "Company Name", "this investor") if not inv_row.empty else "this investor"
    sector      = _mv(inv_row, "Sector", "")                   if not inv_row.empty else ""
    stage       = _mv(inv_row, "Journey Stage", "")            if not inv_row.empty else ""
    blocker_lvl = _mv(inv_row, "Blocker Level", "None")        if not inv_row.empty else "None"
    minister_act = _mv(inv_row, "Minister Action Required", "") if not inv_row.empty else ""

    # ── Parse action items ────────────────────────────────────────────────────
    total = len(acts) if not acts.empty else 0
    completed = in_prog = blocked_n = 0
    high_descs    = []
    blocker_descs = []
    eng_counts: dict = {}

    if not acts.empty and "Status" in acts.columns:
        sl = acts["Status"].fillna("")
        completed = int(sl.str.lower().str.contains("complet").sum())
        blocked_n = int((sl == "Blocked").sum())
        in_prog   = int(sl.isin(["In Progress", "Inprogress"]).sum())

        if "Priority" in acts.columns:
            for _, r in acts[acts["Priority"] == "High"].head(3).iterrows():
                d = str(r.get("Action Description", "")).strip()
                if d and d not in ("nan", ""):
                    high_descs.append(d[:55])

        for _, r in acts[acts["Status"] == "Blocked"].head(2).iterrows():
            d = str(r.get("Action Description", "")).strip()
            if d and d not in ("nan", ""):
                blocker_descs.append(d[:48])

        if "Type of Engagement" in acts.columns:
            for t in acts["Type of Engagement"].dropna():
                ts = str(t).strip()
                if ts and ts not in ("nan", ""):
                    eng_counts[ts] = eng_counts.get(ts, 0) + 1

    top_engs = [e for e, _ in sorted(eng_counts.items(), key=lambda x: -x[1])][:3]

    # ── Parse opportunities ───────────────────────────────────────────────────
    n_opps     = len(opps) if not opps.empty else 0
    opp_val    = _sum_col(opps, "Est. Value (SAR)")
    opp_names  = []
    opp_stages = []
    if not opps.empty:
        if "Opportunity Name" in opps.columns:
            opp_names = [str(n).strip() for n in opps["Opportunity Name"].dropna().head(2)
                         if str(n).strip() not in ("nan", "")]
        if "Opportunity Stage" in opps.columns:
            opp_stages = list(opps["Opportunity Stage"].dropna().unique()[:3])

    pct = round(completed / total * 100) if total > 0 else 0
    sector_str = f" in the {sector} sector" if sector and sector not in ("—", "") else ""

    # ── GOAL ─────────────────────────────────────────────────────────────────
    if opp_names:
        goal = f"Ministry aims to advance {company}'s investment{sector_str} by delivering: {', '.join(opp_names)}."
    elif top_engs:
        goal = f"Secure committed investment{sector_str} through structured {', '.join(top_engs[:2]).lower()} with {company}."
    else:
        goal = f"Build strategic partnership with {company}{sector_str} and progress engagement to commitment stage."
    if opp_val > 0:
        goal += f"  Target pipeline: {_fmt_sar(opp_val)}."

    # ── STRATEGY ─────────────────────────────────────────────────────────────
    if total > 0:
        strat = f"{pct}% of {total} actions complete ({completed} done"
        if in_prog:
            strat += f", {in_prog} in progress"
        strat += "). "
    else:
        strat = "Engagement in early stage. "

    if top_engs:
        strat += f"MISA approach: {', '.join(top_engs[:2])}. "
    if opp_stages:
        strat += f"Opportunities at: {', '.join(opp_stages[:2])}."
    if n_opps == 0:
        strat += "Next step: formalise opportunity pipeline."

    # ── MINISTER ACTION ───────────────────────────────────────────────────────
    if minister_act and minister_act not in ("None Required", "—", ""):
        m_act = f"ACTION REQUIRED: {minister_act}. Immediate senior-level engagement needed."
    elif blocked_n > 0 and blocker_descs:
        m_act = f"ESCALATE: '{blocker_descs[0]}' — ministerial intervention required to unblock and restore momentum."
    elif blocked_n > 0:
        m_act = f"{blocked_n} action{'s' if blocked_n > 1 else ''} blocked. Minister to intervene directly — signal MISA's commitment to resolve."
    elif blocker_lvl not in ("None", "—", "nan", ""):
        m_act = f"Blocker level: {blocker_lvl}. Minister to champion resolution and reassure company at executive level."
    elif high_descs:
        m_act = f"Champion: '{high_descs[0]}'. Ministerial endorsement will accelerate delivery and signal strategic priority."
    elif stage in ("Active Negotiation",):
        m_act = "Deal in negotiation. Minister to maintain executive contact — closing requires a clear political commitment signal from ministry leadership."
    elif stage in ("Committed", "Post-Investment"):
        m_act = "Investment committed. Minister to publicly acknowledge relationship and identify next expansion opportunity."
    elif pct >= 70 and n_opps > 0:
        m_act = "Engagement maturing. Minister to initiate commitment conversation — redirect dialogue to deal closure and headline terms."
    else:
        m_act = "Strengthen partnership: arrange senior bilateral meeting, present MISA's strategic value proposition and Vision 2030 alignment."

    return goal[:220], strat.strip()[:220], m_act[:220]


def _mv(row, key: str, default: str = "—") -> str:
    """Safely extract a metadata value from an investor row, suppressing NaN."""
    val = row.get(key, default)
    if val is None:
        return default
    if isinstance(val, float) and pd.isna(val):
        return default
    s = str(val).strip()
    return s if s and s.lower() not in ("nan", "none", "") else default


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
