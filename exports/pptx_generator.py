
# PowerPoint auto-generator — compact deck with merged sections.
# Full deck: Title + 4 summary slides + per-investor slides.
# Company deck: 3 slides.

import io
from datetime import date
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

    _co_slide_cover_profile(prs, company, inv_row, inv_opps, inv_acts, lang)
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

def _co_slide_cover_profile(prs, company, inv_row, opps, acts, lang):
    """Cover + profile + key metrics."""
    slide = _blank_slide(prs)
    _fill_background(slide, GREEN)

    # Company name block
    _add_text_box(slide, company,
                  Inches(1), Inches(1.5), Inches(11.33), Inches(1.1),
                  font_size=38, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tier = inv_row.get("Investor Tier", "") if not inv_row.empty else ""
    _add_text_box(slide, tier,
                  Inches(1), Inches(2.7), Inches(11.33), Inches(0.5),
                  font_size=16, color=GOLD, align=PP_ALIGN.CENTER)
    _add_text_box(slide, f"Investor Status Report — {date.today().strftime('%d %B %Y')}",
                  Inches(1), Inches(3.3), Inches(11.33), Inches(0.4),
                  font_size=13, color=_rgb("FFFFFFBB"), align=PP_ALIGN.CENTER)

    # White info band
    _add_rect(slide, Inches(0), Inches(4.0), Inches(13.33), Inches(2.6),
              fill_color=WHITE, line_color=WHITE)

    if not inv_row.empty:
        meta = [
            ("Country",         inv_row.get("Country", "—")),
            ("Sector",          inv_row.get("Sector", "—")),
            ("Journey Stage",   inv_row.get("Journey Stage", "—")),
            ("Status",          inv_row.get("Relationship Status", "—")),
            ("RM",              inv_row.get("Relationship Manager", "—")),
            ("AM",              inv_row.get("Account Manager", "TBD")),
        ]
        for i, (lbl, val) in enumerate(meta):
            x = Inches(0.4) + i * Inches(2.1)
            _add_text_box(slide, lbl, x, Inches(4.1), Inches(2.0), Inches(0.25),
                          font_size=9, color=MGRAY)
            _add_text_box(slide, str(val)[:22], x, Inches(4.35), Inches(2.0), Inches(0.35),
                          font_size=12, bold=True, color=DARK)

        est  = inv_row.get("Est. Investment Value (SAR)")
        cmmt = inv_row.get("Actual Commitment (SAR)")
        jobs = inv_row.get("Est. Jobs Created")
        min_act = str(inv_row.get("Minister Action Required", "None Required") or "None Required")
        next_m  = str(inv_row.get("Next Meeting Date", "—") or "—")

        kpis = [
            ("Est. Investment",    _fmt_sar(est) if pd.notna(est) and est else "—",          MISA_GREEN),
            ("Committed",          _fmt_sar(cmmt) if pd.notna(cmmt) and cmmt else "—",       MISA_GOLD),
            ("Est. Jobs",          f"{int(jobs):,}" if pd.notna(jobs) and jobs else "—",     "#2D7A54"),
            ("Open Opportunities", str(len(opps)),                                            MISA_GREEN),
            ("Pending Actions",    str(len(acts[~acts.get("Status", pd.Series(dtype=str)).isin(["Completed","Cancelled"])]) if not acts.empty and "Status" in acts.columns else len(acts)), MISA_GOLD),
            ("Next Meeting",       next_m[:12],                                              MISA_GREEN),
        ]
        for i, (lbl, val, col) in enumerate(kpis):
            x = Inches(0.4) + i * Inches(2.1)
            _add_rect(slide, x, Inches(4.85), Inches(1.95), Inches(1.2),
                      fill_color=_rgb(col), line_color=_rgb(col))
            _add_text_box(slide, val, x, Inches(4.9), Inches(1.95), Inches(0.65),
                          font_size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            _add_text_box(slide, lbl, x, Inches(5.53), Inches(1.95), Inches(0.25),
                          font_size=8, color=WHITE, align=PP_ALIGN.CENTER)

    # Gold footer
    _add_rect(slide, Inches(0), Inches(6.9), Inches(13.33), Inches(0.6),
              fill_color=GOLD, line_color=GOLD)
    _add_text_box(slide, "CONFIDENTIAL | Ministry of Investment — وزارة الاستثمار",
                  Inches(0), Inches(6.9), Inches(13.33), Inches(0.6),
                  font_size=11, color=WHITE, align=PP_ALIGN.CENTER)


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
