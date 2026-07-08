
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
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE

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


def _co_summary_slide(prs, company, inv_row, actions, opportunities, meetings, lang):
    """Lightweight one-slide company summary for the all-companies dashboard.
    Uses only rectangles and text boxes — no embedded charts, no network calls."""
    slide = _blank_slide(prs)
    _mv = lambda row, col, default="": (row.get(col, default) or default) if not (hasattr(row, "empty") and row.empty) else default

    # Header
    _add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(0.82), fill_color=GREEN, line_color=GREEN)
    _add_text_box(slide, company, Inches(0.3), Inches(0.08), Inches(9.5), Inches(0.66),
                  font_size=18, bold=True, color=WHITE)
    sector  = _mv(inv_row, "Sector", "")
    country = _mv(inv_row, "Country", "")
    tier    = _mv(inv_row, "Tier", "")
    meta    = "  |  ".join(x for x in [sector, country, tier] if x)
    if meta:
        _add_text_box(slide, meta, Inches(0.3), Inches(0.60), Inches(9.5), Inches(0.20),
                      font_size=8, color=_rgb("#C8E6D4"))

    # Action items summary (left panel)
    co_acts = (actions[actions["Company Name"] == company]
               if not actions.empty and "Company Name" in actions.columns
               else pd.DataFrame())
    n_total = len(co_acts)
    n_done  = int(co_acts["Status"].str.lower().str.contains("complet").sum()) if not co_acts.empty and "Status" in co_acts.columns else 0
    n_prog  = int(co_acts["Status"].isin(["In Progress", "Inprogress"]).sum()) if not co_acts.empty and "Status" in co_acts.columns else 0
    pct     = round(n_done / n_total * 100) if n_total else 0

    _add_text_box(slide, "Action Items", Inches(0.3), Inches(1.0), Inches(4.5), Inches(0.25),
                  font_size=10, bold=True, color=DARK)
    _add_text_box(slide, f"{n_total} total  |  {n_done} completed  |  {n_prog} in progress  |  {pct}% done",
                  Inches(0.3), Inches(1.28), Inches(6.0), Inches(0.22), font_size=8.5, color=MGRAY)

    # Progress bar
    _add_rect(slide, Inches(0.3), Inches(1.56), Inches(6.0), Inches(0.18),
              fill_color=_rgb("#E0E0E0"), line_color=_rgb("#E0E0E0"))
    if pct > 0:
        _add_rect(slide, Inches(0.3), Inches(1.56), Inches(6.0) * pct / 100, Inches(0.18),
                  fill_color=GREEN, line_color=GREEN)

    # Top action items (up to 8)
    _ay = Inches(1.90)
    pend = co_acts[~co_acts["Status"].isin(["Completed", "Cancelled"])].head(8) if not co_acts.empty else pd.DataFrame()
    for i, (_, row) in enumerate(pend.iterrows()):
        desc   = str(row.get("Action Description", "") or "")[:80]
        status = str(row.get("Status", "") or "")
        owner  = str(row.get("Assigned To", "") or "")[:20]
        alt    = _rgb("#F7F7F2") if i % 2 == 0 else WHITE
        _add_rect(slide, Inches(0.3), _ay, Inches(9.0), Inches(0.38),
                  fill_color=alt, line_color=_rgb("#DDDDDD"))
        _add_text_box(slide, f"{i+1}. {desc}", Inches(0.35), _ay + Inches(0.04),
                      Inches(5.8), Inches(0.28), font_size=8, color=DARK)
        _add_text_box(slide, owner, Inches(6.2), _ay + Inches(0.04),
                      Inches(1.5), Inches(0.28), font_size=8, color=MGRAY)
        _add_text_box(slide, status[:14], Inches(7.75), _ay + Inches(0.04),
                      Inches(1.5), Inches(0.28), font_size=8, color=DARK)
        _ay += Inches(0.40)
        if _ay > Inches(6.8):
            break

    # Opportunities (right panel strip)
    co_opps = (opportunities[opportunities["Company Name"] == company]
               if not opportunities.empty and "Company Name" in opportunities.columns
               else pd.DataFrame())
    n_opps = len(co_opps)
    _add_rect(slide, Inches(9.3), Inches(1.0), Inches(3.8), Inches(5.8),
              fill_color=_rgb("#F5F5F0"), line_color=_rgb("#DDDDDD"))
    _add_text_box(slide, f"Opportunities ({n_opps})",
                  Inches(9.4), Inches(1.05), Inches(3.6), Inches(0.26),
                  font_size=9, bold=True, color=GREEN)
    _oy = Inches(1.38)
    if not co_opps.empty and "Opportunity Name" in co_opps.columns:
        for _, orow in co_opps.head(8).iterrows():
            oname  = str(orow.get("Opportunity Name", "") or "")[:45]
            ostage = str(orow.get("Opportunity Stage", "") or "")[:18]
            if not oname or oname == "nan":
                continue
            _add_text_box(slide, f"• {oname}", Inches(9.4), _oy, Inches(3.6), Inches(0.20),
                          font_size=7.5, color=DARK)
            if ostage and ostage != "nan":
                _add_text_box(slide, ostage, Inches(9.4), _oy + Inches(0.19), Inches(3.6), Inches(0.16),
                              font_size=6.5, color=MGRAY)
                _oy += Inches(0.38)
            else:
                _oy += Inches(0.22)
            if _oy > Inches(6.6):
                break

    # Minister action / blocker
    min_act = _mv(inv_row, "Minister Action Required", "") or _mv(inv_row, "Immediate Action", "")
    if min_act and str(min_act) not in ("nan", ""):
        _add_rect(slide, Inches(0.3), Inches(6.85), Inches(8.9), Inches(0.30),
                  fill_color=_rgb("#C0392B"), line_color=_rgb("#C0392B"))
        _add_text_box(slide, f"Immediate Action: {str(min_act)[:120]}",
                      Inches(0.35), Inches(6.87), Inches(8.8), Inches(0.24),
                      font_size=7.5, bold=True, color=WHITE)

    # Gold footer
    _add_rect(slide, Inches(0), Inches(7.05), SLIDE_W, Inches(0.45),
              fill_color=GOLD, line_color=GOLD)
    _add_text_box(slide, "CONFIDENTIAL | Ministry of Investment — وزارة الاستثمار",
                  Inches(0), Inches(7.05), SLIDE_W, Inches(0.45),
                  font_size=10, color=WHITE, align=PP_ALIGN.CENTER)


def _slide_dashboard_title_cover(prs):
    """Intro cover slide for the All-Companies dashboard deck."""
    slide = _blank_slide(prs)
    _fill_background(slide, _rgb(MISA_GREEN))

    # Decorative gold band at top third
    _add_rect(slide, Inches(0), Inches(2.40), Inches(13.33), Inches(0.05),
              fill_color=_rgb(MISA_GOLD), line_color=_rgb(MISA_GOLD))
    _add_rect(slide, Inches(0), Inches(4.80), Inches(13.33), Inches(0.05),
              fill_color=_rgb(MISA_GOLD), line_color=_rgb(MISA_GOLD))

    # Arabic & English ministry label
    _add_text_box(slide, "وزارة الاستثمار  |  Ministry of Investment",
                  Inches(1), Inches(1.50), Inches(11.33), Inches(0.40),
                  font_size=13, color=_rgb(MISA_GOLD), align=PP_ALIGN.CENTER)

    # Main title
    _add_text_box(slide, "Man-marking Weekly Report",
                  Inches(1), Inches(2.55), Inches(11.33), Inches(0.90),
                  font_size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Subtitle
    _add_text_box(slide, "Executive Outreach  |  Minister Office",
                  Inches(1), Inches(3.52), Inches(11.33), Inches(0.38),
                  font_size=14, color=_rgb(MISA_GOLD), align=PP_ALIGN.CENTER)

    # Date
    _add_text_box(slide, date.today().strftime("%d %B %Y"),
                  Inches(1), Inches(5.00), Inches(11.33), Inches(0.32),
                  font_size=11, color=WHITE, align=PP_ALIGN.CENTER)

    # Confidential note
    _add_text_box(slide, "FOR INTERNAL USE ONLY",
                  Inches(1), Inches(5.36), Inches(11.33), Inches(0.26),
                  font_size=9, color=_rgb("#AADDBB"), align=PP_ALIGN.CENTER)

    # Gold footer
    _add_rect(slide, Inches(0), Inches(7.05), Inches(13.33), Inches(0.45),
              fill_color=_rgb(MISA_GOLD), line_color=_rgb(MISA_GOLD))
    _add_text_box(slide, "CONFIDENTIAL | Ministry of Investment — وزارة الاستثمار",
                  Inches(0), Inches(7.05), Inches(13.33), Inches(0.45),
                  font_size=10, color=WHITE, align=PP_ALIGN.CENTER)


def _slide_end_thankyou(prs):
    slide = _blank_slide(prs)
    _fill_background(slide, _rgb(MISA_GREEN))
    # Decorative gold horizontal band
    _add_rect(slide, Inches(0), Inches(3.30), Inches(13.33), Inches(0.04),
              fill_color=_rgb(MISA_GOLD), line_color=_rgb(MISA_GOLD))
    _add_text_box(slide, "شكراً",
                  Inches(1), Inches(1.40), Inches(11.33), Inches(1.20),
                  font_size=54, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    _add_text_box(slide, "Thank You",
                  Inches(1), Inches(2.60), Inches(11.33), Inches(0.80),
                  font_size=32, bold=True, color=_rgb(MISA_GOLD), align=PP_ALIGN.CENTER)
    _add_text_box(slide, "Minister Office  ·  Executive Outreach",
                  Inches(1), Inches(3.55), Inches(11.33), Inches(0.32),
                  font_size=12, color=_rgb(MISA_GOLD), align=PP_ALIGN.CENTER)
    _add_text_box(slide, "Ministry of Investment — وزارة الاستثمار",
                  Inches(1), Inches(3.90), Inches(11.33), Inches(0.30),
                  font_size=11, color=WHITE, align=PP_ALIGN.CENTER)
    _add_rect(slide, Inches(0), Inches(7.05), Inches(13.33), Inches(0.45),
              fill_color=_rgb(MISA_GOLD), line_color=_rgb(MISA_GOLD))
    _add_text_box(slide, "CONFIDENTIAL | Ministry of Investment — وزارة الاستثمار",
                  Inches(0), Inches(7.05), Inches(13.33), Inches(0.45),
                  font_size=10, color=WHITE, align=PP_ALIGN.CENTER)


def generate_pptx_all_companies_dashboard(dfs: dict, lang: str = "en") -> bytes:
    """
    Strategic all-companies dashboard deck.
    Slide 1: Cover with KPI strip + progress bars + sector/country charts.
    Slides 2+: Per-company slides (cover + opps/deals) for every investor.
    """
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    investors     = dfs.get("Investor Master",      pd.DataFrame())
    meetings      = dfs.get("Meeting Log",          pd.DataFrame())
    opportunities = dfs.get("Opportunity Pipeline", pd.DataFrame())
    actions       = dfs.get("Action Items",         pd.DataFrame())
    deals         = dfs.get("Deal Progress",        pd.DataFrame())

    # ── Slide 1: Title / Intro cover ─────────────────────────────────────────
    _slide_dashboard_title_cover(prs)

    # ── Slide 2: KPI Summary ─────────────────────────────────────────────────
    slide = _blank_slide(prs)

    # Full-width green header
    _add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.96),
              fill_color=GREEN, line_color=GREEN)
    _add_text_box(slide, "Ministry of Investment — All Companies Dashboard",
                  Inches(0.3), Inches(0.04), Inches(9.5), Inches(0.40),
                  font_size=18, bold=True, color=WHITE)
    _add_text_box(slide, "Minister Office  ·  Executive Outreach",
                  Inches(0.3), Inches(0.44), Inches(7.0), Inches(0.22),
                  font_size=9, bold=False, color=GOLD)
    _add_text_box(slide, "Man-marking Weekly Report",
                  Inches(0.3), Inches(0.64), Inches(7.0), Inches(0.20),
                  font_size=8, color=WHITE)
    _add_text_box(slide, date.today().strftime("%d %B %Y"),
                  Inches(10.3), Inches(0.30), Inches(2.8), Inches(0.30),
                  font_size=9, color=GOLD, align=PP_ALIGN.RIGHT)

    # ── KPI strip (5 cards including overall progress) ───────────────────────
    total_cos    = len(investors)
    total_acts   = len(actions)
    n_done_a     = int(actions["Status"].str.lower().str.contains("complet").sum()) if not actions.empty and "Status" in actions.columns else 0
    n_prog_a     = int(actions["Status"].isin(["In Progress", "Inprogress"]).sum()) if not actions.empty and "Status" in actions.columns else 0
    _ovr_pct     = round((n_done_a + n_prog_a) / max(total_acts, 1) * 100)

    kpi_data = [
        ("Total Companies",   str(total_cos),         MISA_GREEN),
        ("Total Actions",     str(total_acts),         MISA_GOLD),
        ("Completed",         str(n_done_a),           MISA_GREEN),
        ("In Progress",       str(n_prog_a),           MISA_GOLD),
        ("Overall Progress",  f"{_ovr_pct}%",          MISA_GREEN),
    ]
    kcard_w = Inches(2.47)
    kcard_h = Inches(0.72)
    for i, (lbl, val, col) in enumerate(kpi_data):
        kx = Inches(0.3) + i * (kcard_w + Inches(0.08))
        _add_rect(slide, kx, Inches(0.90), kcard_w, kcard_h,
                  fill_color=_rgb(col), line_color=_rgb(col))
        _add_text_box(slide, val, kx, Inches(0.92), kcard_w, Inches(0.38),
                      font_size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _add_text_box(slide, lbl, kx, Inches(1.28), kcard_w, Inches(0.26),
                      font_size=9, color=WHITE, align=PP_ALIGN.CENTER)

    # ── Left 60%: Actions Progress by Company ────────────────────────────────
    _add_text_box(slide, "Actions Progress by Company",
                  Inches(0.3), Inches(1.80), Inches(7.8), Inches(0.24),
                  font_size=11, bold=True, color=DARK)

    if not actions.empty and "Company Name" in actions.columns and "Status" in actions.columns:
        co_stats = {}
        for co_name_g, grp in actions.groupby("Company Name"):
            total_g = len(grp)
            done_g  = int(grp["Status"].str.lower().str.contains("complet").sum())
            prog_g  = int(grp["Status"].isin(["In Progress", "Inprogress"]).sum())
            ns_g    = max(total_g - done_g - prog_g, 0)
            pct_g   = round((done_g + prog_g) / total_g * 100) if total_g > 0 else 0
            co_stats[co_name_g] = (pct_g, done_g, prog_g, ns_g, total_g)

        # Average of per-company progress %
        _avg_co_pct = round(sum(v[0] for v in co_stats.values()) / max(len(co_stats), 1)) if co_stats else _ovr_pct
        MAX_BAR_W = Inches(5.4)
        by = Inches(2.06)
        bar_h   = Inches(0.30)
        bar_gap = Inches(0.08)

        # Overall average row
        _add_text_box(slide, "OVERALL AVG", Inches(0.3), by,
                      Inches(2.0), Inches(0.22), font_size=8, bold=True, color=GREEN)
        bx = Inches(2.35)
        _add_rect(slide, bx, by + Inches(0.03), MAX_BAR_W, bar_h,
                  fill_color=_rgb("#E0E0E0"), line_color=_rgb("#E0E0E0"))
        _add_rect(slide, bx, by + Inches(0.03), MAX_BAR_W * _avg_co_pct / 100, bar_h,
                  fill_color=GREEN, line_color=GREEN)
        _add_text_box(slide, f"{_avg_co_pct}%", bx + MAX_BAR_W + Inches(0.08), by,
                      Inches(0.5), Inches(0.22), font_size=8, bold=True, color=GREEN)
        by += bar_h + Inches(0.14)

        sorted_cos = sorted(co_stats.items(), key=lambda x: x[1][0], reverse=True)[:8]

        for co_name_b, (pct_g, done_g, prog_g, ns_g, total_g) in sorted_cos:
            _add_text_box(slide, str(co_name_b)[:24], Inches(0.3), by,
                          Inches(2.0), Inches(0.26), font_size=8, color=DARK)
            bx = Inches(2.35)
            _add_rect(slide, bx, by + Inches(0.04), MAX_BAR_W, bar_h,
                      fill_color=_rgb("#E0E0E0"), line_color=_rgb("#E0E0E0"))
            if done_g > 0 and total_g > 0:
                done_w = MAX_BAR_W * done_g / total_g
                _add_rect(slide, bx, by + Inches(0.04), done_w, bar_h,
                          fill_color=GREEN, line_color=GREEN)
            if prog_g > 0 and total_g > 0:
                prog_off = MAX_BAR_W * done_g / total_g
                prog_w   = MAX_BAR_W * prog_g / total_g
                _add_rect(slide, bx + prog_off, by + Inches(0.04), prog_w, bar_h,
                          fill_color=GOLD, line_color=GOLD)
            _add_text_box(slide, f"{pct_g}%", bx + MAX_BAR_W + Inches(0.08), by,
                          Inches(0.5), Inches(0.26), font_size=8, bold=True, color=DARK)
            by += bar_h + bar_gap
            if by > Inches(6.5):
                break

    # ── Opportunities by Sector (left side bottom) ──────────────────────────
    n_total_opps = len(opportunities) if not opportunities.empty else 0
    n_active_opps = (
        int((opportunities["Opportunity Status"] == "Active").sum())
        if not opportunities.empty and "Opportunity Status" in opportunities.columns
        else n_total_opps
    )
    OPP_SEC_X = Inches(0.3)
    OPP_SEC_Y = Inches(6.08)
    _add_text_box(slide, "Opportunities by Sector",
                  OPP_SEC_X, OPP_SEC_Y, Inches(2.60), Inches(0.22),
                  font_size=9, bold=True, color=GREEN)
    _add_text_box(slide, f"{n_total_opps} total  |  active: {n_active_opps}",
                  OPP_SEC_X + Inches(2.65), OPP_SEC_Y, Inches(1.15), Inches(0.22),
                  font_size=7.5, color=_rgb(MISA_GOLD), align=PP_ALIGN.RIGHT)
    # Build sector counts by joining opportunities → investor sector
    _sec_opp_counts = pd.Series(dtype=int)
    if (not opportunities.empty and not investors.empty
            and "Company Name" in opportunities.columns
            and "Company Name" in investors.columns
            and "Sector" in investors.columns):
        _inv_sec = investors[["Company Name", "Sector"]].rename(columns={"Sector": "_inv_sector"})
        _opp_merged = opportunities.merge(_inv_sec, on="Company Name", how="left")
        _sec_opp_counts = _opp_merged["_inv_sector"].fillna("Unknown").value_counts().head(5)
    elif n_total_opps > 0:
        _sec_opp_counts = pd.Series({"All Sectors": n_total_opps})
    _SEC_OPP_COLORS = [MISA_GREEN, MISA_GOLD, "#2D7A54", "#E0B06A", "#888888"]
    _MAX_BAR_OPP = Inches(1.50)
    _sec_opp_y = OPP_SEC_Y + Inches(0.26)
    for _si, (_sname, _scnt) in enumerate(_sec_opp_counts.items()):
        _col = _SEC_OPP_COLORS[_si % len(_SEC_OPP_COLORS)]
        _bar_frac = _scnt / max(int(_sec_opp_counts.max()), 1)
        _add_rect(slide, OPP_SEC_X, _sec_opp_y + Inches(0.04), Inches(0.08), Inches(0.08),
                  fill_color=_rgb(_col), line_color=_rgb(_col))
        _add_text_box(slide, str(_sname)[:16], OPP_SEC_X + Inches(0.13), _sec_opp_y,
                      Inches(1.35), Inches(0.18), font_size=7.5, color=DARK)
        _add_rect(slide, OPP_SEC_X + Inches(1.55), _sec_opp_y + Inches(0.04),
                  _MAX_BAR_OPP * _bar_frac, Inches(0.12),
                  fill_color=_rgb(_col), line_color=_rgb(_col))
        _add_text_box(slide, str(_scnt),
                      OPP_SEC_X + Inches(1.55) + _MAX_BAR_OPP + Inches(0.06), _sec_opp_y,
                      Inches(0.30), Inches(0.18), font_size=7.5, bold=True, color=DARK)
        _sec_opp_y += Inches(0.18)
        if _sec_opp_y > Inches(6.92):
            break

    # ── Right: By Sector donut chart ──────────────────────────────────────────
    _DONUT_PALETTE = ["#1B5C3F", "#2D7A54", "#3D9068", "#C9974A", "#E0B06A", "#888888"]
    SEC_X = Inches(8.35)
    _add_text_box(slide, "By Sector",
                  SEC_X, Inches(1.80), Inches(4.7), Inches(0.26),
                  font_size=11, bold=True, color=DARK)
    if not investors.empty and "Sector" in investors.columns:
        sec_counts = investors["Sector"].dropna().value_counts().head(6)
        try:
            _sec_data = ChartData()
            _sec_data.categories = [str(s)[:14] for s in sec_counts.index]
            _sec_data.add_series("Sectors", tuple(int(v) for v in sec_counts.values))
            _sec_gfx = slide.shapes.add_chart(
                XL_CHART_TYPE.DOUGHNUT,
                Inches(8.40), Inches(2.10), Inches(2.40), Inches(2.15),
                _sec_data,
            )
            _sec_obj = _sec_gfx.chart
            _sec_obj.has_title = False
            _sec_obj.has_legend = False
            for _pi, _pt in enumerate(_sec_obj.series[0].points):
                _pt.format.fill.solid()
                _pt.format.fill.fore_color.rgb = _rgb(_DONUT_PALETTE[_pi % len(_DONUT_PALETTE)])
        except Exception:
            pass
        # Text legend beside donut
        _leg_x = Inches(10.90)
        _leg_y = Inches(2.15)
        for _si, (_sname, _scnt) in enumerate(sec_counts.items()):
            _col = _DONUT_PALETTE[_si % len(_DONUT_PALETTE)]
            _add_rect(slide, _leg_x, _leg_y + Inches(0.04), Inches(0.09), Inches(0.09),
                      fill_color=_rgb(_col), line_color=_rgb(_col))
            _add_text_box(slide, f"{str(_sname)[:14]}: {_scnt}",
                          _leg_x + Inches(0.13), _leg_y,
                          Inches(2.25), Inches(0.20), font_size=7.5, color=DARK)
            _leg_y += Inches(0.26)

    # ── Right: By Country donut chart ─────────────────────────────────────────
    CTY_X = Inches(8.35)
    _add_text_box(slide, "By Country",
                  CTY_X, Inches(4.45), Inches(4.7), Inches(0.26),
                  font_size=11, bold=True, color=DARK)
    if not investors.empty and "Country" in investors.columns:
        cty_counts = investors["Country"].dropna().value_counts().head(6)
        try:
            _cty_data = ChartData()
            _cty_data.categories = [str(c)[:14] for c in cty_counts.index]
            _cty_data.add_series("Countries", tuple(int(v) for v in cty_counts.values))
            _cty_gfx = slide.shapes.add_chart(
                XL_CHART_TYPE.DOUGHNUT,
                Inches(8.40), Inches(4.72), Inches(2.40), Inches(2.15),
                _cty_data,
            )
            _cty_obj = _cty_gfx.chart
            _cty_obj.has_title = False
            _cty_obj.has_legend = False
            _CTY_PAL = ["#C9974A", "#E0B06A", "#1B5C3F", "#2D7A54", "#888888", "#AAAAAA"]
            for _pi, _pt in enumerate(_cty_obj.series[0].points):
                _pt.format.fill.solid()
                _pt.format.fill.fore_color.rgb = _rgb(_CTY_PAL[_pi % len(_CTY_PAL)])
        except Exception:
            pass
        # Text legend beside donut
        _leg_x = Inches(10.90)
        _leg_y = Inches(4.77)
        _CTY_PAL2 = ["#C9974A", "#E0B06A", "#1B5C3F", "#2D7A54", "#888888", "#AAAAAA"]
        for _ci, (_cname, _ccnt) in enumerate(cty_counts.items()):
            _col = _CTY_PAL2[_ci % len(_CTY_PAL2)]
            _add_rect(slide, _leg_x, _leg_y + Inches(0.04), Inches(0.09), Inches(0.09),
                      fill_color=_rgb(_col), line_color=_rgb(_col))
            _add_text_box(slide, f"{str(_cname)[:14]}: {_ccnt}",
                          _leg_x + Inches(0.13), _leg_y,
                          Inches(2.25), Inches(0.20), font_size=7.5, color=DARK)
            _leg_y += Inches(0.26)

    # Gold footer
    _add_rect(slide, Inches(0), Inches(7.05), Inches(13.33), Inches(0.45),
              fill_color=GOLD, line_color=GOLD)
    _add_text_box(slide, "CONFIDENTIAL | Ministry of Investment — وزارة الاستثمار",
                  Inches(0), Inches(7.05), Inches(13.33), Inches(0.45),
                  font_size=10, color=WHITE, align=PP_ALIGN.CENTER)

    # ── Slides 2+: one per company (full design, logos + charts skipped for speed)
    if not investors.empty and "Company Name" in investors.columns:
        for _, inv in investors.iterrows():
            co = inv.get("Company Name", "")
            if not co:
                continue
            inv_acts = (actions[actions["Company Name"] == co]
                        if not actions.empty and "Company Name" in actions.columns
                        else pd.DataFrame())
            inv_opps = (opportunities[opportunities["Company Name"] == co]
                        if not opportunities.empty and "Company Name" in opportunities.columns
                        else pd.DataFrame())
            inv_mtgs = (meetings[meetings["Company Name"] == co]
                        if not meetings.empty and "Company Name" in meetings.columns
                        else pd.DataFrame())
            inv_dls  = (deals[deals["Company Name"] == co]
                        if not deals.empty and "Company Name" in deals.columns
                        else pd.DataFrame())
            try:
                _co_slide_cover_profile(prs, co, inv, inv_opps, inv_acts, inv_mtgs, inv_dls, lang,
                                        _skip_logos=True, _skip_charts=True)
            except Exception:
                pass

    _slide_end_thankyou(prs)

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
    """Left: Immediate action log. Right: Vision 2030 + Strategy pillars."""
    slide = _blank_slide(prs)
    _add_slide_header(slide, "Immediate Actions & Vision 2030")

    # ── LEFT: Immediate action items ────────────────────────────────
    _add_text_box(slide, "Immediate Actions Required",
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
        f"Immediate Action: {min_act}  |  Blocker: {blocker}"
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

def _co_slide_cover_profile(prs, company, inv_row, opps, acts, meetings, deals, lang, _skip_logos=False, _skip_charts=False):
    """Slide 1 (merged 1+3): Header + metadata + Brief strip + compact timeline + full action items table (L) + status/deals/KPI (R)."""
    slide = _blank_slide(prs)

    # ── Green header band ────────────────────────────────────────────
    _add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.82),
              fill_color=GREEN, line_color=GREEN)
    # Try to show company logo in header
    _s1_name_x = Inches(0.3)
    if not inv_row.empty and not _skip_logos:
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
                  Inches(8.8), Inches(0.20), Inches(4.2), Inches(0.30),
                  font_size=8, color=_rgb("CCCCCC"), align=PP_ALIGN.RIGHT)
    _add_text_box(slide,
                  "Ministry of Investment  ·  Minister Office  ·  Executive Outreach  ·  Man-marking Weekly Report",
                  Inches(0.3), Inches(0.64), Inches(13.0), Inches(0.16),
                  font_size=6.5, color=GOLD)

    # ── White metadata band ──────────────────────────────────────────
    _add_rect(slide, Inches(0), Inches(0.82), Inches(13.33), Inches(0.58),
              fill_color=WHITE, line_color=WHITE)
    if not inv_row.empty:
        for i, (lbl, val) in enumerate([
            ("Country",       _mv(inv_row, "Country")),
            ("Sector",        _mv(inv_row, "Sector")),
            ("Journey Stage", _mv(inv_row, "Journey Stage")),
            ("RM",            _mv(inv_row, "Relationship Manager")),
            ("AM",            _mv(inv_row, "Account Manager", "TBD")),
        ]):
            x = Inches(0.4) + i * Inches(2.52)
            _add_text_box(slide, lbl, x, Inches(0.87), Inches(2.4), Inches(0.18),
                          font_size=8, color=MGRAY)
            _add_text_box(slide, val[:26], x, Inches(1.04), Inches(2.4), Inches(0.3),
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

    # ── 3-panel brief strip: Strategic Goal | Ministry Strategy | Minister Action ──
    P1_X = Inches(0.20);  P1_W = Inches(4.20)
    P2_X = Inches(4.55);  P2_W = Inches(4.10)
    P3_X = Inches(8.80);  P3_W = Inches(4.35)
    DIV_H = BRIEF_H - Inches(0.16)
    for div_x in (Inches(4.47), Inches(8.72)):
        _add_rect(slide, div_x, BRIEF_Y + Inches(0.08), Inches(0.015), DIV_H,
                  fill_color=_rgb("#CCCCCC"), line_color=_rgb("#CCCCCC"))
    # Panel 1 — Strategic Goal
    _add_text_box(slide, "STRATEGIC GOAL", P1_X, BRIEF_Y + Inches(0.07),
                  P1_W, Inches(0.22), font_size=9, bold=True, color=_rgb(MISA_GREEN))
    _add_text_box(slide, goal_text, P1_X, BRIEF_Y + Inches(0.30),
                  P1_W, Inches(0.54), font_size=8.5, color=DARK)
    # Panel 2 — Overall Progress
    _add_text_box(slide, "OVERALL PROGRESS", P2_X, BRIEF_Y + Inches(0.07),
                  P2_W, Inches(0.22), font_size=9, bold=True, color=_rgb(MISA_GOLD))
    _add_text_box(slide, strat_text, P2_X, BRIEF_Y + Inches(0.30),
                  P2_W, Inches(0.54), font_size=8.5, color=DARK)
    # Panel 3 — Immediate Action
    _add_text_box(slide, "IMMEDIATE ACTION", P3_X, BRIEF_Y + Inches(0.07),
                  P3_W, Inches(0.22), font_size=9, bold=True, color=RED)
    _add_text_box(slide, action_text, P3_X, BRIEF_Y + Inches(0.30),
                  P3_W, Inches(0.54), font_size=8.5, color=DARK)

    # ══════════════════════════════════════════════════════════════════════════
    # LEFT: Action items table (Engagement Timeline removed — replaced by Due Dates strip)
    # ══════════════════════════════════════════════════════════════════════════

    # Status colour map — used by vertical due-date timeline
    _ACT_SC = {
        "Completed":   MISA_GREEN, "In Progress": MISA_GOLD, "Inprogress": MISA_GOLD,
        "Not Started": "#888888",  "Blocked": "#C0392B",      "Cancelled": "#AAAAAA",
    }

    # ── Legend row rendered between section heading (y=2.42) and table header ─
    # Placed at y=2.68 as a single horizontal strip so the table never covers it
    _LG_Y  = Inches(2.68)
    _LG_SQ = Inches(0.09)
    _mtg_items = [
        ("Completed mtg", MISA_GREEN), ("Scheduled", MISA_GOLD), ("Cancelled", "#AAAAAA"),
    ]
    for ri, (lbl, col) in enumerate(_mtg_items):
        lx = Inches(5.75) + ri * Inches(1.10)
        _add_rect(slide, lx, _LG_Y + Inches(0.06), _LG_SQ, _LG_SQ,
                  fill_color=_rgb(col), line_color=_rgb(col))
        _add_text_box(slide, lbl, lx + Inches(0.12), _LG_Y,
                      Inches(1.00), Inches(0.20), font_size=6.5, color=MGRAY)
    _add_text_box(slide, "|", Inches(9.06), _LG_Y, Inches(0.15), Inches(0.20),
                  font_size=6.5, color=MGRAY, align=PP_ALIGN.CENTER)
    _act_items = [
        ("Done", MISA_GREEN), ("In Progress", MISA_GOLD),
        ("Pending", "#888888"), ("Blocked", "#C0392B"),
    ]
    for ri, (lbl, col) in enumerate(_act_items):
        lx = Inches(9.22) + ri * Inches(0.82)
        _add_rect(slide, lx, _LG_Y + Inches(0.06), _LG_SQ, _LG_SQ,
                  fill_color=_rgb(col), line_color=_rgb(col))
        _add_text_box(slide, lbl, lx + Inches(0.12), _LG_Y,
                      Inches(0.72), Inches(0.20), font_size=6.5, color=MGRAY)

    # ── Action counts ─────────────────────────────────────────────────────────
    n_total  = len(acts)
    n_done_s = int(acts["Status"].str.lower().str.contains("complet").sum()) if not acts.empty and "Status" in acts.columns else 0
    n_prog_s = int(acts["Status"].isin(["In Progress","Inprogress"]).sum()) if not acts.empty and "Status" in acts.columns else 0
    n_due_s  = max(n_total - n_done_s - n_prog_s, 0)
    # Average of individual action Progress values (0–1 fractions stored per row)
    if not acts.empty and "Progress" in acts.columns and n_total > 0:
        def _parse_prog(x):
            try:
                v = float(str(x).replace("%", "")) if x not in (None, "", "nan") else 0.0
                return v / 100.0 if v > 1.0 else v
            except Exception:
                return 0.0
        # Completed actions count as 100%; use stored Progress for others
        _prog_series = acts.apply(
            lambda r: 1.0 if str(r.get("Status", "")).lower() in ("completed", "cancelled")
            else _parse_prog(r.get("Progress", 0)), axis=1)
        pct_s = round(_prog_series.mean() * 100)
    else:
        pct_s = round((n_done_s + n_prog_s) / n_total * 100) if n_total > 0 else 0

    if not acts.empty and "Status" in acts.columns:
        pend_df = acts[~acts["Status"].isin(["Completed", "Cancelled"])].copy()
        done_df = acts[acts["Status"].isin(["Completed", "Cancelled"])].copy()
        if "Due Date" in pend_df.columns:
            pend_df["_due"] = pd.to_datetime(pend_df["Due Date"], errors="coerce")
            pend_df = pend_df.sort_values("_due")
    else:
        pend_df = acts.copy() if not acts.empty else pd.DataFrame()
        done_df = pd.DataFrame()

    # ════════════════════════════════════════════════════════════════════════
    # LEFT PANEL — doughnut chart + opportunities  (x: 0.20 → 5.00)
    # RIGHT PANEL — two stacked tables             (x: 5.10 → 13.10)
    # ════════════════════════════════════════════════════════════════════════

    # ── LEFT PANEL: Doughnut status chart ────────────────────────────────────
    _donut_vals = [max(n_done_s, 0), max(n_prog_s, 0), max(n_due_s, 0)]
    if sum(_donut_vals) == 0:
        _donut_vals = [0, 0, 1]
    if _skip_charts:
        # Circle chart — concentric ovals (area-proportional, no embedded Excel)
        _tot = max(sum(_donut_vals), 1)
        _pct_any  = min(1.0, (_donut_vals[0] + _donut_vals[1]) / _tot)  # completed+in-progress
        _pct_done = min(1.0, _donut_vals[0] / _tot)                     # completed only
        _ccx = Inches(2.30)   # circle center x
        _ccy = Inches(4.35)   # circle center y
        _CR  = Inches(1.10)   # outer radius

        def _oval_c(cx, cy, r, color):
            r = max(r, Inches(0.05))
            s = slide.shapes.add_shape(9, cx - r, cy - r, r * 2, r * 2)
            s.fill.solid(); s.fill.fore_color.rgb = color
            s.line.fill.background()

        _oval_c(_ccx, _ccy, _CR,                                    _rgb("#E6F2EA"))          # bg soft green
        _oval_c(_ccx, _ccy, _CR * max(_pct_any  ** 0.5, 0.08),     _rgb(MISA_GOLD))          # in-prog+done
        _oval_c(_ccx, _ccy, _CR * max(_pct_done ** 0.5, 0.08),     _rgb(MISA_GREEN))         # done
        _oval_c(_ccx, _ccy, _CR * 0.46,                             _rgb("#FFFFFF"))          # hole
    else:
        try:
            _chart_data = ChartData()
            _chart_data.categories = ["Completed", "In Progress", "Not Started"]
            _chart_data.add_series("Status", tuple(_donut_vals))
            _chart_gfx = slide.shapes.add_chart(
                XL_CHART_TYPE.DOUGHNUT,
                Inches(0.5), Inches(3.05), Inches(2.5), Inches(2.20),
                _chart_data,
            )
            _chart_obj = _chart_gfx.chart
            _chart_obj.has_title = False
            _chart_obj.has_legend = False
            _donut_colors = [_rgb("#1B5C3F"), _rgb("#C9974A"), _rgb("#888888")]
            _chart_series = _chart_obj.series[0]
            for _pi, _pc in enumerate(_donut_colors):
                _pt = _chart_series.points[_pi]
                _pt.format.fill.solid()
                _pt.format.fill.fore_color.rgb = _pc
        except Exception:
            _add_rect(slide, Inches(0.2), Inches(3.3), Inches(2.5), Inches(2.5),
                      fill_color=_rgb("#F0FFF4"), line_color=_rgb("#C8E0D4"))

    # Center label overlaid on doughnut hole
    _add_text_box(slide, f"{pct_s}%",
                  Inches(1.80), Inches(4.10), Inches(1.0), Inches(0.40),
                  font_size=18, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    _add_text_box(slide, "progress",
                  Inches(1.80), Inches(4.47), Inches(1.0), Inches(0.18),
                  font_size=7, color=MGRAY, align=PP_ALIGN.CENTER)

    # ── LEFT PANEL: Opportunities list (x constrained to 0.5"–3.0") ──────────
    _n_opps_left = len(opps) if not opps.empty else 0
    _n_active_opps = int((opps["Opportunity Status"] == "Active").sum()) \
        if not opps.empty and "Opportunity Status" in opps.columns else _n_opps_left
    _add_text_box(slide, f"Opportunities ({_n_opps_left})",
                  Inches(0.5), Inches(6.08), Inches(3.50), Inches(0.28),
                  font_size=11, bold=True, color=GREEN)
    if _n_active_opps > 0:
        _add_text_box(slide, f"Active: {_n_active_opps}",
                      Inches(3.40), Inches(6.10), Inches(1.20), Inches(0.22),
                      font_size=8, color=_rgb(MISA_GOLD), align=PP_ALIGN.RIGHT)
    _opp_item_y = Inches(6.40)
    if not opps.empty and "Opportunity Name" in opps.columns:
        _opp_names_list = [
            (str(r.get("Opportunity Name", "") or "").strip(),
             str(r.get("Opportunity Stage", "") or "").strip())
            for _, r in opps.head(4).iterrows()
            if str(r.get("Opportunity Name", "") or "").strip() not in ("nan", "")
        ]
        # 2-column grid: 2 opportunities per row
        _OPP_COL_W = Inches(2.45)
        _OPP_ROW_H = Inches(0.26)
        for _oi, (_oname, _ostage) in enumerate(_opp_names_list):
            _col_i = _oi % 2
            _row_i = _oi // 2
            _ox = Inches(0.50) + _col_i * (_OPP_COL_W + Inches(0.10))
            _oy = _opp_item_y + _row_i * _OPP_ROW_H
            _add_rect(slide, _ox, _oy + Inches(0.06), Inches(0.09), Inches(0.09),
                      fill_color=_rgb(MISA_GOLD), line_color=_rgb(MISA_GOLD))
            _stage_txt = f"  [{_ostage}]" if _ostage else ""
            _add_text_box(slide, f"{_oname[:22]}{_stage_txt}",
                          _ox + Inches(0.13), _oy, _OPP_COL_W - Inches(0.15), Inches(0.22),
                          font_size=8, color=DARK)
    else:
        _add_text_box(slide, "No opportunities recorded.",
                      Inches(0.5), _opp_item_y, Inches(3.90), Inches(0.22),
                      font_size=8, color=MGRAY)

    # Thin vertical divider between left and right panels
    _add_rect(slide, Inches(5.70), Inches(2.42), Inches(0.015), Inches(4.60),
              fill_color=_rgb("#DDDDDD"), line_color=_rgb("#DDDDDD"))

    # ── RIGHT PANEL: Two stacked action tables ────────────────────────────────
    _n_pend_label = len(pend_df)
    _n_done_label = len(done_df)
    # Heading moved to vertical strip on far right (see below after timeline)
    TBL_X  = Inches(5.72)
    HDR_Y  = Inches(2.94)
    HDR_H  = Inches(0.32)
    ROW_H  = Inches(0.46)
    MAX_R  = 4
    CW_NEW = [Inches(w) for w in [0.30, 2.85, 1.50, 1.50, 0.75]]

    def _render_tbl_new(df, tbl_y, title, hdr_color, max_r=MAX_R):
        CX = [TBL_X + sum(CW_NEW[:j]) for j in range(len(CW_NEW))]
        col_hdrs = ["#", title, "Owner", "Update the Status", "%"]
        for hdr, cx, cw in zip(col_hdrs, CX, CW_NEW):
            _add_rect(slide, cx, tbl_y, cw, HDR_H, fill_color=hdr_color, line_color=hdr_color)
            _add_text_box(slide, hdr, cx + Inches(0.02), tbl_y + Inches(0.04),
                          cw - Inches(0.04), HDR_H - Inches(0.06),
                          font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        for i, (_, row) in enumerate(df.head(max_r).iterrows()):
            ry     = tbl_y + HDR_H + i * ROW_H
            alt    = _rgb("#F7F7F2") if i % 2 == 0 else WHITE
            desc   = str(row.get("Action Description", "") or "")
            owner  = str(row.get("Assigned To", "") or "")[:18]
            status = str(row.get("Status", "") or "")
            remark = str(row.get("Remarks", "") or "").strip()
            am_inp = str(row.get("AM Input", "") or "").strip()
            is_done = status.lower() in ("completed", "cancelled")
            try:
                prog_raw = row.get("Progress", 0)
                prog_val = float(str(prog_raw).replace("%", "")) if prog_raw not in (None, "", "nan") else 0.0
                if prog_val > 1.0:
                    prog_val /= 100.0
            except Exception:
                prog_val = 0.0
            if is_done:
                prog_val = 1.0
            # Col 0: row number
            _add_rect(slide, CX[0], ry, CW_NEW[0], ROW_H, fill_color=alt, line_color=_rgb("#DDDDDD"))
            _add_text_box(slide, str(i + 1), CX[0] + Inches(0.03), ry + Inches(0.07),
                          CW_NEW[0] - Inches(0.05), Inches(0.22),
                          font_size=9, color=DARK, align=PP_ALIGN.CENTER)
            # Col 1: action description only (no note — note moves to Update the Status col)
            _add_rect(slide, CX[1], ry, CW_NEW[1], ROW_H, fill_color=alt, line_color=_rgb("#DDDDDD"))
            _add_text_box(slide, desc[:70], CX[1] + Inches(0.04), ry + Inches(0.12),
                          CW_NEW[1] - Inches(0.08), Inches(0.25), font_size=9, color=DARK)
            note = (remark if remark and remark not in ("nan",) else "") or \
                   (am_inp if am_inp and am_inp not in ("nan",) else "")
            # Col 2: owner
            _add_rect(slide, CX[2], ry, CW_NEW[2], ROW_H, fill_color=alt, line_color=_rgb("#DDDDDD"))
            _add_text_box(slide, owner, CX[2] + Inches(0.03), ry + Inches(0.07),
                          CW_NEW[2] - Inches(0.05), Inches(0.22), font_size=8.5, color=DARK)
            # Col 3: update note only (no status label)
            _add_rect(slide, CX[3], ry, CW_NEW[3], ROW_H, fill_color=alt, line_color=_rgb("#DDDDDD"))
            if note:
                _add_text_box(slide, note[:55], CX[3] + Inches(0.05), ry + Inches(0.07),
                              CW_NEW[3] - Inches(0.08), Inches(0.32),
                              font_size=7.5, color=DARK)
            # Col 4: date labels + progress bar + % label
            px, pcw = CX[4], CW_NEW[4]
            _add_rect(slide, px, ry, pcw, ROW_H, fill_color=alt, line_color=_rgb("#DDDDDD"))
            # Date labels above bar
            def _fmt_dt(v):
                try:
                    d = pd.to_datetime(v, errors="coerce")
                    return d.strftime("%b '%y") if pd.notna(d) else ""
                except Exception:
                    return ""
            _start_raw = row.get("Start Date", row.get("Created Date", None))
            _due_raw   = row.get("Due Date", None)
            _s_lbl = _fmt_dt(_start_raw)
            _d_lbl = _fmt_dt(_due_raw)
            if _s_lbl:
                _add_text_box(slide, _s_lbl, px + Inches(0.02), ry + Inches(0.01),
                              pcw * 0.55, Inches(0.09),
                              font_size=5, color=_rgb("#888888"))
            if _d_lbl:
                _add_text_box(slide, _d_lbl, px + pcw * 0.45, ry + Inches(0.01),
                              pcw * 0.55, Inches(0.09),
                              font_size=5, color=_rgb("#888888"), align=PP_ALIGN.RIGHT)
            # Bar
            bx, by = px + Inches(0.03), ry + Inches(0.12)
            bw, bh = pcw - Inches(0.06), Inches(0.07)
            _add_rect(slide, bx, by, bw, bh, fill_color=_rgb("#E0E0E0"), line_color=_rgb("#E0E0E0"))
            if prog_val > 0:
                fc2 = (_rgb(MISA_GREEN) if prog_val >= 1.0
                       else (_rgb(MISA_GOLD) if prog_val >= 0.5 else _rgb("#888888")))
                _add_rect(slide, bx, by, max(bw * prog_val, Inches(0.02)), bh,
                          fill_color=fc2, line_color=fc2)
            _add_text_box(slide, "100%" if is_done else f"{int(prog_val * 100)}%",
                          px, ry + Inches(0.26), pcw, Inches(0.18),
                          font_size=8, color=DARK, align=PP_ALIGN.CENTER)
        if len(df) > max_r:
            overflow_y = tbl_y + HDR_H + max_r * ROW_H
            _add_text_box(slide, f"+ {len(df) - max_r} more",
                          TBL_X + Inches(0.10), overflow_y, Inches(3.0), Inches(0.18),
                          font_size=7.5, color=MGRAY)
        return tbl_y + HDR_H + min(len(df), max_r) * ROW_H

    tbl1_end = _render_tbl_new(pend_df, HDR_Y, "Pending / In Progress", GREEN, max_r=4)
    if not done_df.empty:
        _render_tbl_new(done_df, tbl1_end + Inches(0.15), "Completed", _rgb("#2D7A54"), max_r=3)

    # ── Horizontal due-date timeline — above opportunities, full left panel width ─
    try:
        if not acts.empty and "Due Date" in acts.columns:
            _act_tl = acts.copy()
            _act_tl["_dt"] = pd.to_datetime(_act_tl["Due Date"], errors="coerce")
            _act_tl = _act_tl[_act_tl["_dt"].notna()].sort_values("_dt")
            if not _act_tl.empty:
                _today = date.today()
                _vt_min_dt = _act_tl["_dt"].min().date()
                _vt_max_dt = _act_tl["_dt"].max().date()
                _vt_min_dt = max(_vt_min_dt, _today.replace(day=1) - timedelta(days=548))
                _vt_max_dt = min(_vt_max_dt, _today.replace(day=28) + timedelta(days=548))
                _vt_min_dt = _vt_min_dt.replace(day=1)
                _ld = _calendar.monthrange(_vt_max_dt.year, _vt_max_dt.month)[1]
                _vt_max_dt = _vt_max_dt.replace(day=_ld)
                if _today < _vt_min_dt:
                    _vt_min_dt = _today.replace(day=1)
                if _today > _vt_max_dt:
                    _ld2 = _calendar.monthrange(_today.year, _today.month)[1]
                    _vt_max_dt = _today.replace(day=_ld2)
                _span = max((_vt_max_dt - _vt_min_dt).days, 1)

                HX_L    = Inches(0.45)
                HX_W    = Inches(5.20)
                HX_BG_Y = Inches(2.42)
                HX_BG_H = Inches(0.56)          # taller strip to fit month labels
                HX_AX_Y = HX_BG_Y + Inches(0.30)  # axis at 2.72" — room above for dots+numbers

                def _hx(d):
                    if hasattr(d, "date") and callable(d.date):
                        d = d.date()
                    frac = max(0.0, min(1.0, (d - _vt_min_dt).days / _span))
                    return HX_L + frac * HX_W

                _add_rect(slide, HX_L, HX_BG_Y, HX_W, HX_BG_H,
                          fill_color=_rgb("#F5F5F0"), line_color=_rgb("#DDDDDD"))
                _add_text_box(slide, "Action Due Dates",
                              HX_L, HX_BG_Y + Inches(0.02), HX_W, Inches(0.14),
                              font_size=7, bold=True, color=DARK)
                _add_rect(slide, HX_L, HX_AX_Y, HX_W, Inches(0.012),
                          fill_color=_rgb("#CCCCCC"), line_color=_rgb("#CCCCCC"))

                _n_months = max(1, (_vt_max_dt.year - _vt_min_dt.year) * 12
                                + (_vt_max_dt.month - _vt_min_dt.month))
                _show_every = 1 if _n_months <= 12 else 2
                _mo = _vt_min_dt.replace(day=1)
                _mo_count = 0
                while _mo <= _vt_max_dt and _mo_count < 24:
                    _mo_count += 1
                    _mx = _hx(_mo)
                    # Tick mark below axis
                    _add_rect(slide, _mx, HX_AX_Y, Inches(0.008), Inches(0.08),
                              fill_color=_rgb("#AAAAAA"), line_color=_rgb("#AAAAAA"))
                    # Month label — show all when ≤12 months, every other when >12
                    if _mo_count % _show_every == 0:
                        _add_text_box(slide, _mo.strftime("%b"), _mx - Inches(0.18),
                                      HX_AX_Y + Inches(0.09), Inches(0.38), Inches(0.16),
                                      font_size=7, color=_rgb("#555555"), align=PP_ALIGN.CENTER)
                    _mo = (_mo.replace(year=_mo.year + 1, month=1)
                           if _mo.month == 12
                           else _mo.replace(month=_mo.month + 1))

                _now_x = _hx(_today)
                _add_rect(slide, _now_x - Inches(0.005), HX_AX_Y - Inches(0.10),
                          Inches(0.010), Inches(0.20),
                          fill_color=GOLD, line_color=GOLD)
                _add_text_box(slide, "NOW", _now_x - Inches(0.15), HX_AX_Y - Inches(0.18),
                              Inches(0.32), Inches(0.10),
                              font_size=5.5, bold=True, color=GOLD, align=PP_ALIGN.CENTER)

                for _ai, (_, _arow) in enumerate(_act_tl.iterrows()):
                    _adx = _hx(_arow["_dt"])
                    _ast = str(_arow.get("Status", ""))
                    _asc = _rgb(_ACT_SC.get(_ast, "#888888"))
                    _dy = Inches(0.09) if _ai % 2 == 0 else Inches(0.04)
                    _add_rect(slide, _adx - Inches(0.035), HX_AX_Y - _dy - Inches(0.07),
                              Inches(0.07), Inches(0.07),
                              fill_color=_asc, line_color=_asc)
                    _add_text_box(slide, str(_ai + 1),
                                  _adx - Inches(0.09), HX_AX_Y - _dy - Inches(0.16),
                                  Inches(0.18), Inches(0.10),
                                  font_size=5, color=DARK, align=PP_ALIGN.CENTER)
    except Exception:
        pass  # timeline is non-critical; never block slide generation

    # ── Vertical action items heading — far right strip (x=13.05" to 13.33") ─
    # Placed after timeline so it draws on top; rotation=90 → reads bottom-to-top
    _vai_text = f"Action Items — {n_total} total | {_n_pend_label} pending / in progress"
    _add_rect(slide, Inches(13.05), Inches(2.42), Inches(0.28), Inches(4.63),
              fill_color=GREEN, line_color=GREEN)
    _vai_tb = slide.shapes.add_textbox(Inches(10.875), Inches(4.595), Inches(4.63), Inches(0.28))
    _vai_tf = _vai_tb.text_frame
    _vai_tf.word_wrap = False
    _vai_tf.auto_size = MSO_AUTO_SIZE.NONE
    _vai_p = _vai_tf.paragraphs[0]
    _vai_p.alignment = PP_ALIGN.CENTER
    _vai_run = _vai_p.add_run()
    _vai_run.text = _vai_text
    _vai_run.font.size = Pt(7.5)
    _vai_run.font.bold = True
    _vai_run.font.color.rgb = WHITE
    _vai_tb.rotation = 90

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
    # Primary: formal Opportunity Pipeline entries for this company.
    # Fallback: action items tagged Type of Engagement = "Opportunity" when pipeline is empty.
    if not opps.empty and "Opportunity Name" in opps.columns:
        card_source = opps[~opps["Opportunity Name"].fillna("").str.strip().str.lower().eq("action item")]
    elif not acts.empty and "Type of Engagement" in acts.columns and "Action Description" in acts.columns:
        opp_acts = acts[acts["Type of Engagement"].str.strip().str.lower() == "opportunity"].copy()
        if not opp_acts.empty:
            opp_acts = opp_acts.rename(columns={"Action Description": "Opportunity Name"})
            opp_acts["Opportunity Status"] = "Active"
            opp_acts["Opportunity Stage"]  = "Exploration"
            opp_acts["Confidence Level"]   = "Suggested"
            card_source = opp_acts
            opps        = opp_acts
        else:
            card_source = opps
    else:
        card_source = opps
    n_cards  = len(card_source)
    n_active = int((opps["Opportunity Status"] == "Active").sum())    if not opps.empty and "Opportunity Status" in opps.columns else 0
    n_done   = int((opps["Opportunity Stage"]  == "Committed").sum()) if not opps.empty and "Opportunity Stage"  in opps.columns else 0
    total_val   = _sum_col(opps, "Est. Value (SAR)")

    # ── KPI strip (4 cards across full width)
    kw = Inches(3.13)
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

    CARD_W  = Inches(4.08)
    CARD_H  = Inches(1.65)
    GAP_X   = Inches(0.13)
    GAP_Y   = Inches(0.12)
    COL_X   = [Inches(0.24), Inches(0.24) + CARD_W + GAP_X]
    START_Y = Inches(1.73)
    SB_X    = Inches(8.70)
    SB_W    = Inches(4.48)

    if card_source.empty:
        _add_text_box(slide,
                      "No opportunities or action items have been recorded for this investor.",
                      Inches(0.3), Inches(2.5), Inches(12.73), Inches(0.5),
                      font_size=13, color=MGRAY, align=PP_ALIGN.CENTER)

    elif False:  # legacy act_cards path — kept for structure
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

            # Row 3: 2-line auto description + impact
            opp_desc, opp_impact = _build_opportunity_narrative(opp, acts, idx)
            _add_text_box(slide, opp_desc[:68],
                          IX, cy + Inches(0.62), IW, Inches(0.14),
                          font_size=7.5, color=DARK)
            _add_text_box(slide, opp_impact[:72],
                          IX, cy + Inches(0.77), IW, Inches(0.12),
                          font_size=6.5, color=MGRAY)

            # Rows 4+: Related action items (up to 2)
            _ACT_DOT = {
                "Completed":   MISA_GREEN, "Inprogress":  MISA_GOLD,
                "In Progress": MISA_GOLD,  "Not Started": "#AAAAAA",
                "Blocked":     "#C0392B",  "Cancelled":   "#CCCCCC",
            }
            matched_acts = _match_acts_to_opp(opp_name, acts)
            act_y = cy + Inches(0.92)
            ACT_ROW = Inches(0.27)
            shown_acts = 0
            for _, ar in matched_acts.head(2).iterrows():
                if act_y + ACT_ROW > cy + CARD_H - Inches(0.18):
                    break
                ad = str(ar.get("Action Description", "") or "").strip()
                st = str(ar.get("Status",             "") or "").strip()
                if not ad or ad in ("nan",):
                    continue
                dot_col = _rgb(_ACT_DOT.get(st, "#AAAAAA"))
                _add_rect(slide, IX, act_y + Inches(0.04), Inches(0.07), Inches(0.07),
                          fill_color=dot_col, line_color=dot_col)
                _add_text_box(slide, ad[:58], IX + Inches(0.11), act_y,
                              IW - Inches(0.11), Inches(0.17),
                              font_size=7, color=DARK)
                act_y += ACT_ROW
                shown_acts += 1

            # Footer: completion count
            if not matched_acts.empty and "Status" in matched_acts.columns:
                n_tot  = len(matched_acts)
                n_done = int(matched_acts["Status"].str.lower().str.contains("complet").sum())
                _add_text_box(slide,
                              f"{n_done}/{n_tot} completed",
                              IX, cy + CARD_H - Inches(0.18), IW, Inches(0.16),
                              font_size=6.5, color=MGRAY)

    # ── Ministry Strategy Sidebar (right of opportunity cards) ───────────────
    _add_rect(slide, SB_X - Inches(0.07), Inches(1.60), Inches(0.02), Inches(5.42),
              fill_color=_rgb("#DDDDDD"), line_color=_rgb("#DDDDDD"))
    # Sidebar title band
    _add_rect(slide, SB_X, Inches(1.73), SB_W, Inches(0.30),
              fill_color=_rgb(MISA_GOLD), line_color=_rgb(MISA_GOLD))
    _add_text_box(slide, "OVERALL PROGRESS",
                  SB_X + Inches(0.10), Inches(1.77),
                  SB_W - Inches(0.20), Inches(0.22),
                  font_size=9, bold=True, color=WHITE)

    _SB_PILLAR_META = {
        "Attract Investment": {
            "alignment": "FDI growth — Vision 2030 pipeline targets",
            "team_aim":  "Secure commitment from target investors",
            "color":     "#065F46",
            "light":     "#D1FAE5",
        },
        "Matchmaking": {
            "alignment": "Connects investors with MISA priority sectors",
            "team_aim":  "Facilitate introductions & sector engagement",
            "color":     "#1D4ED8",
            "light":     "#DBEAFE",
        },
        "Resolve Challenges": {
            "alignment": "Removes barriers to investment facilitation",
            "team_aim":  "Cross-ministry coordination to unblock deals",
            "color":     "#92400E",
            "light":     "#FEF3C7",
        },
    }

    _sb_pillar_acts = {p: [] for p in _SB_PILLAR_META}
    if not acts.empty and "Action Description" in acts.columns:
        for _, _sbr in acts.iterrows():
            _sbp = _map_to_strategy_pillar(
                str(_sbr.get("Action Description", "")) + " " +
                str(_sbr.get("Type of Engagement", "")) + " " +
                str(_sbr.get("Sector", ""))
            )
            if _sbp in _sb_pillar_acts:
                _sb_pillar_acts[_sbp].append(_sbr)

    _sb_y = Inches(2.10)
    for _sb_pillar, _sb_meta in _SB_PILLAR_META.items():
        _sb_items = _sb_pillar_acts[_sb_pillar]
        _sb_cnt   = len(_sb_items)
        _sb_pcol  = _sb_meta["color"]
        _sb_plight = _sb_meta["light"]
        _sb_n_act  = min(_sb_cnt, 3)
        _sb_detail_h = Inches(0.36 + _sb_n_act * 0.20)

        if _sb_y + Inches(0.30) + _sb_detail_h > Inches(7.02):
            break

        # Pillar header
        _add_rect(slide, SB_X, _sb_y, SB_W, Inches(0.28),
                  fill_color=_rgb(_sb_pcol), line_color=_rgb(_sb_pcol))
        _add_text_box(slide,
                      f"{_sb_pillar}  —  {_sb_cnt} action{'s' if _sb_cnt != 1 else ''}",
                      SB_X + Inches(0.08), _sb_y + Inches(0.05),
                      SB_W - Inches(0.16), Inches(0.20),
                      font_size=8, bold=True, color=WHITE)
        _sb_y += Inches(0.28)

        # Light detail block
        _add_rect(slide, SB_X, _sb_y, SB_W, _sb_detail_h,
                  fill_color=_rgb(_sb_plight), line_color=_rgb(_sb_plight))

        # Alignment + team aim
        _add_text_box(slide, f"↗ {_sb_meta['alignment']}",
                      SB_X + Inches(0.08), _sb_y + Inches(0.03),
                      SB_W - Inches(0.16), Inches(0.13),
                      font_size=6.5, color=_rgb(_sb_pcol))
        _add_text_box(slide, f"▸ {_sb_meta['team_aim']}",
                      SB_X + Inches(0.08), _sb_y + Inches(0.17),
                      SB_W - Inches(0.16), Inches(0.13),
                      font_size=6.5, color=DARK)

        # Top action items
        _sb_act_y = _sb_y + Inches(0.33)
        for _sbact in _sb_items[:_sb_n_act]:
            _sb_ad = str(_sbact.get("Action Description", "") or "").strip()[:56]
            if _sb_ad:
                _add_text_box(slide, f"• {_sb_ad}",
                              SB_X + Inches(0.08), _sb_act_y,
                              SB_W - Inches(0.16), Inches(0.16),
                              font_size=6, color=DARK)
                _sb_act_y += Inches(0.19)

        _sb_y += _sb_detail_h + Inches(0.08)

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
    tf.auto_size = MSO_AUTO_SIZE.NONE
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
            with urllib.request.urlopen(req, timeout=1) as resp:
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


_STRATEGY_PILLAR_KEYWORDS = {
    "Attract Investment": ["invest", "capital", "fund", "pipeline", "value", "narrative",
                           "promotion", "roadshow", "airport", "infrastructure", "strategy",
                           "financial", "aum", "commitment", "deal"],
    "Matchmaking":        ["meeting", "introduction", "introductory", "connect", "sector",
                           "ict", "health", "fintech", "data center", "humain", "modon",
                           "stc", "interhealth", "engage", "explore", "outreach", "bilateral"],
    "Resolve Challenges": ["nda", "residency", "regulatory", "guidance", "energy",
                           "challenge", "resolve", "land", "lease", "fiber", "request",
                           "support", "premium", "cost", "manufacturing", "modon", "land"],
}


def _map_to_strategy_pillar(text: str) -> str:
    """Map action description to the nearest MISA strategy pillar."""
    t = text.lower()
    scores = {p: sum(1 for kw in kws if kw in t)
              for p, kws in _STRATEGY_PILLAR_KEYWORDS.items()}
    best = max(scores, key=lambda p: scores[p])
    return best if scores[best] > 0 else "Matchmaking"


def _action_context_line(desc: str, eng_type: str = "", sector: str = "") -> str:
    """Generate a short 1-line context description for an action item."""
    desc_l = desc.lower()
    if "meeting" in desc_l or "introductory" in desc_l:
        return f"Engagement action — {eng_type or 'outreach'} to advance the relationship."
    if "nda" in desc_l:
        return "Formal agreement to protect information exchange and enable deeper discussions."
    if "plan" in desc_l or "submit" in desc_l:
        return f"Deliverable from {sector or 'engagement'} workstream — pending review."
    if "regulatory" in desc_l or "guidance" in desc_l:
        return "Regulatory enablement to support the investor's soft-landing in Saudi Arabia."
    if "energy" in desc_l or "cost" in desc_l:
        return "Infrastructure challenge requiring cross-ministry coordination to resolve."
    if "residency" in desc_l:
        return "Premium residency processing to facilitate the investor's team presence in KSA."
    if "fiber" in desc_l or "stc" in desc_l:
        return "Connectivity facilitation to support data centre operational requirements."
    if "land" in desc_l or "modon" in desc_l:
        return "Site allocation support required to proceed with investment facility setup."
    if "roadshow" in desc_l or "promotion" in desc_l:
        return "Investor promotion activity to build pipeline ahead of formal commitment."
    if eng_type:
        return f"{eng_type} workstream{(' — ' + sector) if sector else ''} engagement."
    return "Coordinated action to advance investment engagement objectives."


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

    # ── IMMEDIATE ACTION (renamed from "Minister Action") ────────────────────
    if minister_act and minister_act not in ("None Required", "—", ""):
        m_act = f"{minister_act}. Immediate senior-level engagement needed."
    elif blocked_n > 0 and blocker_descs:
        m_act = f"Resolve: '{blocker_descs[0]}' — senior intervention needed to unblock and restore momentum."
    elif blocked_n > 0:
        m_act = f"{blocked_n} action{'s' if blocked_n > 1 else ''} pending resolution. Signal MISA's commitment to resolve."
    elif blocker_lvl not in ("None", "—", "nan", ""):
        m_act = f"Blocker: {blocker_lvl}. Senior engagement needed to champion resolution."
    elif high_descs:
        m_act = f"Advance: '{high_descs[0]}'. Endorsement will accelerate delivery and signal strategic priority."
    elif stage in ("Active Negotiation",):
        m_act = "Deal in negotiation — maintain executive contact and signal commitment to close."
    elif stage in ("Committed", "Post-Investment"):
        m_act = "Investment committed. Acknowledge relationship and identify next expansion opportunity."
    elif pct >= 70 and n_opps > 0:
        m_act = "Engagement maturing. Initiate commitment conversation — redirect to deal closure."
    else:
        m_act = "Arrange senior bilateral meeting. Present MISA's strategic value proposition and Vision 2030 alignment."

    # ── Strategy Pillar Summary ───────────────────────────────────────────────
    if not acts.empty and "Action Description" in acts.columns:
        pillar_counts: dict = {}
        for _, r in acts.iterrows():
            desc_v   = str(r.get("Action Description", "") or "")
            eng_v    = str(r.get("Type of Engagement", "") or "")
            sec_v    = str(r.get("Sector", "") or "")
            p        = _map_to_strategy_pillar(desc_v + " " + eng_v + " " + sec_v)
            pillar_counts[p] = pillar_counts.get(p, 0) + 1
        if pillar_counts:
            top_p = max(pillar_counts, key=lambda k: pillar_counts[k])
            strat += f" Primary pillar: {top_p}."

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

