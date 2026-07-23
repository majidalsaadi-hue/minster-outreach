
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

SLIDE_W = Inches(20.00)
SLIDE_H = Inches(11.25)

HDR_GREEN  = _rgb("#0B4A2F")
BEIGE_BG   = _rgb("#FAF7EE")
IVORY      = _rgb("#F9F5EA")
BODY_TXT   = _rgb("#2B2B2B")
ALERT_RED  = _rgb("#C0392B")
PRI_RED    = _rgb("#C00000")
PRI_AMB    = _rgb("#FFC000")
INT_GOLD   = _rgb("#C89B3C")


def _donut_image_buf(counts_dict: dict, palette: list, size: int = 480) -> "io.BytesIO | None":
    """Draw a donut chart as a PIL PNG image and return a BytesIO buffer."""
    try:
        from PIL import Image, ImageDraw
        import math as _math

        total = sum(counts_dict.values())
        if total == 0:
            return None

        img = Image.new("RGBA", (size, size), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)

        cx = cy = size // 2
        r_outer = cx - 6
        r_inner = int(r_outer * 0.52)

        angle = -90.0
        for i, (_name, count) in enumerate(counts_dict.items()):
            sweep = 360.0 * count / total
            h = palette[i % len(palette)].lstrip("#")
            fill = (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16), 255)
            draw.pieslice(
                [(cx - r_outer, cy - r_outer), (cx + r_outer, cy + r_outer)],
                start=angle, end=angle + sweep, fill=fill,
            )
            angle += sweep

        # Hollow centre
        draw.ellipse(
            [(cx - r_inner, cy - r_inner), (cx + r_inner, cy + r_inner)],
            fill=(255, 255, 255, 255),
        )

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return buf
    except Exception:
        return None


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

    # Footer
    _add_rect(slide, Inches(0), Inches(10.833), SLIDE_W, Inches(0.417),
              fill_color=_rgb("#0B4A2F"), line_color=_rgb("#0B4A2F"))
    _add_text_box(slide, "CONFIDENTIAL | Ministry of Investment — وزارة الاستثمار",
                  Inches(0), Inches(10.833), SLIDE_W, Inches(0.417),
                  font_size=11, color=_rgb("#C89B3C"), align=PP_ALIGN.CENTER)


def _slide_dashboard_title_cover(prs):
    """Cover slide matching the colleague template — 20"×11.25" dark-green canvas."""
    slide = _blank_slide(prs)
    _fill_background(slide, _rgb(MISA_GREEN))

    # Two horizontal gold accent bands
    _add_rect(slide, Inches(0), Inches(3.50), Inches(20.0), Inches(0.06),
              fill_color=_rgb(MISA_GOLD), line_color=_rgb(MISA_GOLD))
    _add_rect(slide, Inches(0), Inches(7.00), Inches(20.0), Inches(0.06),
              fill_color=_rgb(MISA_GOLD), line_color=_rgb(MISA_GOLD))

    # Arabic & English ministry label
    _add_text_box(slide, "وزارة الاستثمار  |  Ministry of Investment",
                  Inches(2), Inches(2.00), Inches(16.0), Inches(0.55),
                  font_size=18, color=_rgb(MISA_GOLD), align=PP_ALIGN.CENTER)

    # Main title
    _add_text_box(slide, "Man-marking Weekly Report",
                  Inches(2), Inches(3.70), Inches(16.0), Inches(1.20),
                  font_size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Subtitle
    _add_text_box(slide, "Executive Outreach  |  Minister Office",
                  Inches(2), Inches(5.00), Inches(16.0), Inches(0.50),
                  font_size=18, color=_rgb(MISA_GOLD), align=PP_ALIGN.CENTER)

    # Date
    _add_text_box(slide, date.today().strftime("%d %B %Y"),
                  Inches(2), Inches(7.30), Inches(16.0), Inches(0.40),
                  font_size=14, color=WHITE, align=PP_ALIGN.CENTER)

    # Confidential note
    _add_text_box(slide, "FOR INTERNAL USE ONLY  ·  MISA CONFIDENTIAL",
                  Inches(2), Inches(7.74), Inches(16.0), Inches(0.34),
                  font_size=11, color=_rgb("#AADDBB"), align=PP_ALIGN.CENTER)

    # Dark green footer
    _add_rect(slide, Inches(0), Inches(10.833), Inches(20.0), Inches(0.417),
              fill_color=_rgb("#0B4A2F"), line_color=_rgb("#0B4A2F"))
    _add_text_box(slide, "CONFIDENTIAL | Ministry of Investment — وزارة الاستثمار",
                  Inches(0), Inches(10.833), Inches(20.0), Inches(0.417),
                  font_size=11, color=_rgb("#C89B3C"), align=PP_ALIGN.CENTER)


def _slide_end_thankyou(prs):
    """Thank-you closing slide — 20"×11.25" dark-green canvas."""
    slide = _blank_slide(prs)
    _fill_background(slide, _rgb(MISA_GREEN))

    # Gold divider
    _add_rect(slide, Inches(0), Inches(5.00), Inches(20.0), Inches(0.06),
              fill_color=_rgb(MISA_GOLD), line_color=_rgb(MISA_GOLD))

    _add_text_box(slide, "شكراً",
                  Inches(2), Inches(1.80), Inches(16.0), Inches(1.60),
                  font_size=72, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    _add_text_box(slide, "Thank You",
                  Inches(2), Inches(3.50), Inches(16.0), Inches(1.00),
                  font_size=44, bold=True, color=_rgb(MISA_GOLD), align=PP_ALIGN.CENTER)
    _add_text_box(slide, "Minister Office  ·  Executive Outreach",
                  Inches(2), Inches(5.30), Inches(16.0), Inches(0.40),
                  font_size=16, color=_rgb(MISA_GOLD), align=PP_ALIGN.CENTER)
    _add_text_box(slide, "Ministry of Investment — وزارة الاستثمار",
                  Inches(2), Inches(5.74), Inches(16.0), Inches(0.38),
                  font_size=14, color=WHITE, align=PP_ALIGN.CENTER)

    # Footer
    _add_rect(slide, Inches(0), Inches(10.833), Inches(20.0), Inches(0.417),
              fill_color=_rgb("#0B4A2F"), line_color=_rgb("#0B4A2F"))
    _add_text_box(slide, "CONFIDENTIAL | Ministry of Investment — وزارة الاستثمار",
                  Inches(0), Inches(10.833), Inches(20.0), Inches(0.417),
                  font_size=11, color=_rgb("#C89B3C"), align=PP_ALIGN.CENTER)


def generate_pptx_all_companies_dashboard(dfs: dict, lang: str = "en", ministry_note: str = "") -> bytes:
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
    _add_rect(slide, Inches(0), Inches(0), Inches(20.0), Inches(0.96),
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
                  Inches(16.0), Inches(0.30), Inches(3.6), Inches(0.30),
                  font_size=9, color=GOLD, align=PP_ALIGN.RIGHT)

    # ── KPI strip (4 cards — Overall Progress removed) ───────────────────────
    total_cos  = len(investors)
    total_acts = len(actions)
    n_done_a   = int(actions["Status"].str.lower().str.contains("complet").sum()) if not actions.empty and "Status" in actions.columns else 0
    n_prog_a   = int(actions["Status"].isin(["In Progress", "Inprogress"]).sum()) if not actions.empty and "Status" in actions.columns else 0

    kpi_data = [
        ("Total Companies", str(total_cos),  MISA_GREEN, "⊞"),
        ("Total Actions",   str(total_acts), MISA_GOLD,  "✓"),
        ("Completed",       str(n_done_a),   MISA_GREEN, "★"),
        ("In Progress",     str(n_prog_a),   MISA_GOLD,  "↺"),
    ]
    kcard_w = Inches(4.72)
    kcard_h = Inches(1.00)
    _dsq    = Inches(0.54)
    for i, (lbl, val, col, icon) in enumerate(kpi_data):
        kx = Inches(0.30) + i * (kcard_w + Inches(0.15))
        _add_rect(slide, kx, Inches(0.96), kcard_w, kcard_h,
                  fill_color=_rgb(col), line_color=_rgb(col))
        _ksq = slide.shapes.add_shape(5, kx + Inches(0.12), Inches(0.995), _dsq, _dsq)
        _ksq.fill.solid()
        _ksq.fill.fore_color.rgb = _rgb("#0F3D2A") if col == MISA_GREEN else _rgb("#A8793A")
        _ksq.line.fill.background()
        _add_text_box(slide, icon, kx + Inches(0.12), Inches(0.995), _dsq, _dsq,
                      font_size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _txt_kx = kx + _dsq + Inches(0.18)
        _txt_kw = kcard_w - _dsq - Inches(0.22)
        _add_text_box(slide, val, _txt_kx, Inches(1.00), _txt_kw, Inches(0.50),
                      font_size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _add_text_box(slide, lbl, _txt_kx, Inches(1.48), _txt_kw, Inches(0.34),
                      font_size=10, color=WHITE, align=PP_ALIGN.CENTER)

    # ── Layout constants ──────────────────────────────────────────────────────
    _SLD_FOOTER_H = Inches(0.42)
    _IA_STRIP_H   = Inches(1.20)   # IMPORTANT ACTIVATES strip
    _CONTENT_Y    = Inches(2.06)
    _CONTENT_BOT  = SLIDE_H - _SLD_FOOTER_H - _IA_STRIP_H
    _CONTENT_H    = _CONTENT_BOT - _CONTENT_Y

    _DONUT_X  = Inches(15.10)
    _DONUT_W  = SLIDE_W - _DONUT_X - Inches(0.20)
    _BARS_X   = Inches(0.30)
    _CO_NM_W  = Inches(2.60)
    _BAR_END_W = Inches(2.90)
    _BAR_TRK_W = _DONUT_X - _BARS_X - _CO_NM_W - _BAR_END_W - Inches(0.35)

    # ── Bars section title ────────────────────────────────────────────────────
    _add_text_box(slide, "Actions Progress by Company",
                  _BARS_X, _CONTENT_Y, Inches(8.0), Inches(0.26),
                  font_size=11, bold=True, color=DARK)

    # ── Single-row bars per company ───────────────────────────────────────────
    if not actions.empty and "Company Name" in actions.columns and "Status" in actions.columns:
        _co_stats = {}
        for _co_g, _grp in actions.groupby("Company Name"):
            _tot = len(_grp)
            _dn  = int(_grp["Status"].str.lower().str.contains("complet").sum())
            _pr  = int(_grp["Status"].isin(["In Progress", "Inprogress"]).sum())
            _ns  = max(_tot - _dn - _pr, 0)
            _co_stats[_co_g] = (_dn, _pr, _ns, _tot)

        _sorted_cos = sorted(_co_stats.items(), key=lambda x: x[1][3], reverse=True)

        _bar_h   = Inches(0.50)
        _bar_gap = Inches(0.04)
        _by = _CONTENT_Y + Inches(0.32)
        _max_by = _CONTENT_BOT - _bar_h

        _DONUT_PAL = ["#1B5C3F", "#2D7A54", "#3D9068", "#C9974A", "#E0B06A", "#888888"]

        for _co_nb, (_dn, _pr, _ns, _tot) in _sorted_cos:
            if _by > _max_by:
                break

            # Company name
            _add_text_box(slide, str(_co_nb)[:22], _BARS_X, _by,
                          _CO_NM_W, _bar_h,
                          font_size=13, bold=True, color=DARK)

            # 3-segment bar
            _bx  = _BARS_X + _CO_NM_W + Inches(0.08)
            _by2 = _by + Inches(0.10)
            _bsh = _bar_h - Inches(0.18)

            # Grey background (not started)
            _add_rect(slide, _bx, _by2, _BAR_TRK_W, _bsh,
                      fill_color=_rgb("#D0D0D0"), line_color=_rgb("#D0D0D0"))
            # Green (completed)
            if _dn > 0 and _tot > 0:
                _dw = _BAR_TRK_W * _dn / _tot
                _add_rect(slide, _bx, _by2, _dw, _bsh,
                          fill_color=GREEN, line_color=GREEN)
                if _dw > Inches(0.45):
                    _add_text_box(slide, f"{round(_dn/_tot*100)}%",
                                  _bx, _by2, _dw, _bsh,
                                  font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            # Gold (in progress)
            if _pr > 0 and _tot > 0:
                _px = _bx + _BAR_TRK_W * _dn / _tot
                _pw = _BAR_TRK_W * _pr / _tot
                _add_rect(slide, _px, _by2, _pw, _bsh,
                          fill_color=GOLD, line_color=GOLD)
                if _pw > Inches(0.45):
                    _add_text_box(slide, f"{round(_pr/_tot*100)}%",
                                  _px, _by2, _pw, _bsh,
                                  font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            # Grey label (not started)
            if _ns > 0 and _tot > 0:
                _nw = _BAR_TRK_W * _ns / _tot
                _nx = _bx + _BAR_TRK_W * (_dn + _pr) / _tot
                if _nw > Inches(0.45):
                    _add_text_box(slide, f"{round(_ns/_tot*100)}%",
                                  _nx, _by2, _nw, _bsh,
                                  font_size=9, bold=True, color=_rgb("#555555"), align=PP_ALIGN.CENTER)

            # Bar-end: Sector|Country / N opps / date range
            _ex = _bx + _BAR_TRK_W + Inches(0.10)

            _inv_r = investors[investors["Company Name"] == _co_nb] if not investors.empty and "Company Name" in investors.columns else pd.DataFrame()
            _sector  = str(_inv_r.iloc[0].get("Sector",  "") if len(_inv_r) > 0 else "").strip()
            _country = str(_inv_r.iloc[0].get("Country", "") if len(_inv_r) > 0 else "").strip()
            _sec_cty = f"{_sector} | {_country}" if _sector and _country else (_sector or _country)
            _be_third = _bar_h // 3
            _add_text_box(slide, _sec_cty, _ex, _by, _BAR_END_W, _be_third,
                          font_size=12, color=_rgb("#888888"))

            _n_opps = len(opportunities[opportunities["Company Name"] == _co_nb]) if not opportunities.empty and "Company Name" in opportunities.columns else 0
            _add_text_box(slide, f"{_n_opps} opp{'s' if _n_opps != 1 else ''}",
                          _ex, _by + _be_third, _BAR_END_W, _be_third,
                          font_size=9, bold=True, color=DARK)

            _co_acts_r = actions[actions["Company Name"] == _co_nb] if not actions.empty and "Company Name" in actions.columns else pd.DataFrame()
            _dates_txt = ""
            if not _co_acts_r.empty:
                try:
                    _sd = pd.to_datetime(_co_acts_r.get("Start Date", pd.Series(dtype=str)), errors="coerce").dropna()
                    _ed = pd.to_datetime(_co_acts_r.get("Due Date",   pd.Series(dtype=str)), errors="coerce").dropna()
                    if not _sd.empty and not _ed.empty:
                        _dates_txt = f"{_sd.min().strftime('%b %Y')} → {_ed.max().strftime('%b %Y')}"
                except Exception:
                    pass
            if _dates_txt:
                _add_text_box(slide, _dates_txt, _ex, _by + _be_third * 2, _BAR_END_W, _be_third,
                              font_size=9, color=_rgb("#888888"))

            _by += _bar_h + _bar_gap

    # ── Right column: By Sector donut ─────────────────────────────────────────
    _DONUT_PALETTE = ["#1B5C3F", "#2D7A54", "#3D9068", "#C9974A", "#E0B06A", "#888888"]
    _half_h = _CONTENT_H / 2

    _add_text_box(slide, "By Sector",
                  _DONUT_X, _CONTENT_Y, _DONUT_W, Inches(0.26),
                  font_size=10, bold=True, color=DARK)
    if not investors.empty and "Sector" in investors.columns:
        sec_counts = investors["Sector"].dropna().value_counts().head(6)
        _sec_buf = _donut_image_buf(dict(sec_counts), _DONUT_PALETTE)
        if _sec_buf:
            slide.shapes.add_picture(
                _sec_buf, _DONUT_X, _CONTENT_Y + Inches(0.28),
                Inches(3.10), Inches(3.10),
            )
        _leg_x = _DONUT_X + Inches(3.20)
        _leg_y = _CONTENT_Y + Inches(0.32)
        for _si, (_sname, _scnt) in enumerate(sec_counts.items()):
            _col = _DONUT_PALETTE[_si % len(_DONUT_PALETTE)]
            _add_rect(slide, _leg_x, _leg_y + Inches(0.04), Inches(0.09), Inches(0.09),
                      fill_color=_rgb(_col), line_color=_rgb(_col))
            _add_text_box(slide, f"{str(_sname)[:14]}: {_scnt}",
                          _leg_x + Inches(0.13), _leg_y,
                          _DONUT_W - Inches(3.32), Inches(0.22), font_size=9, color=DARK)
            _leg_y += Inches(0.24)

    # ── Right column: By Country donut ────────────────────────────────────────
    _CTY_PAL  = ["#C9974A", "#E0B06A", "#1B5C3F", "#2D7A54", "#888888", "#AAAAAA"]
    _d2_y     = _CONTENT_Y + _half_h + Inches(0.10)

    _add_text_box(slide, "By Country",
                  _DONUT_X, _d2_y, _DONUT_W, Inches(0.26),
                  font_size=10, bold=True, color=DARK)
    if not investors.empty and "Country" in investors.columns:
        cty_counts = investors["Country"].dropna().value_counts().head(6)
        _cty_buf = _donut_image_buf(dict(cty_counts), _CTY_PAL)
        if _cty_buf:
            slide.shapes.add_picture(
                _cty_buf, _DONUT_X, _d2_y + Inches(0.28),
                Inches(3.10), Inches(3.10),
            )
        _leg_x = _DONUT_X + Inches(3.20)
        _leg_y = _d2_y + Inches(0.32)
        for _ci, (_cname, _ccnt) in enumerate(cty_counts.items()):
            _col = _CTY_PAL[_ci % len(_CTY_PAL)]
            _add_rect(slide, _leg_x, _leg_y + Inches(0.04), Inches(0.09), Inches(0.09),
                      fill_color=_rgb(_col), line_color=_rgb(_col))
            _add_text_box(slide, f"{str(_cname)[:14]}: {_ccnt}",
                          _leg_x + Inches(0.13), _leg_y,
                          _DONUT_W - Inches(3.32), Inches(0.22), font_size=9, color=DARK)
            _leg_y += Inches(0.24)

    # ── IMPORTANT ACTIVATES (full-width bottom strip) ─────────────────────────
    _IA_Y     = _CONTENT_BOT
    _IA_X     = Inches(0)
    _IA_W     = SLIDE_W
    _IA_H     = _IA_STRIP_H
    _IA_TTL_H = Inches(0.28)
    _add_rect(slide, _IA_X, _IA_Y, _IA_W, _IA_TTL_H,
              fill_color=_rgb("#0B4A2F"), line_color=_rgb("#0B4A2F"))
    _add_rect(slide, _IA_X, _IA_Y, Inches(0.05), _IA_TTL_H,
              fill_color=_rgb("#C89B3C"), line_color=_rgb("#C89B3C"))
    _add_text_box(slide, "IMPORTANT ACTIVATES",
                  _IA_X + Inches(0.12), _IA_Y + Inches(0.04),
                  _IA_W - Inches(0.15), _IA_TTL_H - Inches(0.06),
                  font_size=9, bold=True, color=_rgb("#C89B3C"))
    _add_rect(slide, _IA_X, _IA_Y + _IA_TTL_H, _IA_W, _IA_H - _IA_TTL_H,
              fill_color=_rgb("#FAF7EE"), line_color=_rgb("#D8D2C0"))

    # Collect high-priority pending/in-progress actions
    _ia_items = []
    try:
        if not actions.empty and "Status" in actions.columns:
            _ia_all = actions[~actions["Status"].isin(["Completed", "Cancelled"])].copy()
            if "Priority" in _ia_all.columns:
                _ia_hi = _ia_all[_ia_all["Priority"].isin(["Very High", "High", "Blocked"])].copy()
            else:
                _ia_hi = _ia_all.copy()
            if "To Be In Dashboard" in _ia_hi.columns:
                _dash_flag = _ia_hi["To Be In Dashboard"].astype(str).str.strip().str.upper()
                _ia_yes = _ia_hi[_dash_flag.isin(["YES", "TOP"])]
                if not _ia_yes.empty:
                    _ia_hi = _ia_yes
            _pri_ord2 = {"Very High": 0, "High": 1, "Blocked": 2, "Medium": 3, "Low": 4}
            if "Priority" in _ia_hi.columns:
                _ia_hi["_po"] = _ia_hi["Priority"].map(_pri_ord2).fillna(5)
                _ia_hi = _ia_hi.sort_values("_po")
            for _, _ia_r in _ia_hi.head(8).iterrows():
                _co_n = str(_ia_r.get("Company Name", "") or "").strip()[:18]
                _ac_n = str(_ia_r.get("Action Description", "") or "").strip()
                if len(_ac_n) > 70:
                    _ac_n = _ac_n[:69].rsplit(" ", 1)[0] + "…"
                _pr_n = str(_ia_r.get("Priority", "") or "").strip()
                _ia_items.append((_co_n, _ac_n, _pr_n))
    except Exception:
        pass

    _ia_content_y = _IA_Y + _IA_TTL_H + Inches(0.06)

    if ministry_note and ministry_note.strip():
        _note_h = Inches(0.55)
        _add_rect(slide, _IA_X + Inches(0.08), _ia_content_y,
                  _IA_W - Inches(0.16), _note_h,
                  fill_color=_rgb("#FFF8E8"), line_color=_rgb("#C89B3C"))
        _add_text_box(slide, f"► {ministry_note.strip()[:220]}",
                      _IA_X + Inches(0.18), _ia_content_y + Inches(0.04),
                      _IA_W - Inches(0.36), _note_h - Inches(0.08),
                      font_size=11, bold=True, color=_rgb("#0B4A2F"))
        _ia_content_y += _note_h + Inches(0.05)

    _ia_item_y = _ia_content_y
    _ia_col_w  = (_IA_W - Inches(0.20)) / max(min(len(_ia_items), 4), 1)
    _ia_rows   = [_ia_items[i:i+4] for i in range(0, len(_ia_items), 4)]
    _ia_row_h  = (_IA_H - (_ia_content_y - _IA_Y) - Inches(0.06)) / max(len(_ia_rows), 1)
    for _ir_idx, _ia_row in enumerate(_ia_rows[:2]):
        _ia_col_x = _IA_X + Inches(0.10)
        _ia_ry = _ia_item_y + _ir_idx * _ia_row_h
        for _ic_idx, (_co_n, _ac_n, _pr_n) in enumerate(_ia_row):
            _pr_col = "#C00000" if _pr_n in ("Very High", "High", "Blocked") else "#888888"
            _iadot = slide.shapes.add_shape(9,
                         _ia_col_x + Inches(0.01),
                         _ia_ry + Inches(0.05),
                         Inches(0.09), Inches(0.09))
            _iadot.fill.solid(); _iadot.fill.fore_color.rgb = _rgb(_pr_col)
            _iadot.line.fill.background()
            _add_text_box(slide, f"{_co_n}  ·  {_ac_n}",
                          _ia_col_x + Inches(0.14), _ia_ry,
                          _ia_col_w - Inches(0.18), _ia_row_h,
                          font_size=11, color=_rgb("#1A1A1A"))
            _ia_col_x += _ia_col_w

    if not _ia_items and not (ministry_note and ministry_note.strip()):
        _add_text_box(slide, "No high-priority actions at this time.",
                      _IA_X + Inches(0.15), _IA_Y + _IA_TTL_H + Inches(0.10),
                      _IA_W - Inches(0.20), Inches(0.22),
                      font_size=8, color=_rgb("#888888"))

    # Dark green footer
    _add_rect(slide, Inches(0), SLIDE_H - _SLD_FOOTER_H, SLIDE_W, _SLD_FOOTER_H,
              fill_color=_rgb("#0B4A2F"), line_color=_rgb("#0B4A2F"))
    _add_text_box(slide, "CONFIDENTIAL | Ministry of Investment — وزارة الاستثمار",
                  Inches(0), SLIDE_H - _SLD_FOOTER_H, SLIDE_W, _SLD_FOOTER_H,
                  font_size=11, color=_rgb("#C89B3C"), align=PP_ALIGN.CENTER)

    # ── Slide 3: Minister's Company Action Summary ────────────────────────────
    _ms = _blank_slide(prs)

    # Header
    _add_rect(_ms, Inches(0), Inches(0), SLIDE_W, Inches(0.72),
              fill_color=_rgb("#0B4A2F"), line_color=_rgb("#0B4A2F"))
    _add_text_box(_ms, "Minister's Company Action Summary",
                  Inches(0.30), Inches(0.04), Inches(14.0), Inches(0.40),
                  font_size=18, bold=True, color=WHITE)
    _add_text_box(_ms, "Progress & Key Actions — All Companies at a Glance",
                  Inches(0.30), Inches(0.44), Inches(10.0), Inches(0.20),
                  font_size=9, color=_rgb("#C89B3C"))
    _add_text_box(_ms, date.today().strftime("%d %B %Y"),
                  Inches(16.0), Inches(0.18), Inches(3.6), Inches(0.26),
                  font_size=9, color=_rgb("#C89B3C"), align=PP_ALIGN.RIGHT)

    # Grid layout: 4 columns × up to 4 rows
    _MS_COLS    = 4
    _MS_COL_GAP = Inches(0.14)
    _MS_ROW_GAP = Inches(0.12)
    _MS_PAD_X   = Inches(0.28)
    _MS_PAD_Y   = Inches(0.80)
    _MS_FOOTER_H = Inches(0.36)
    _MS_CARD_W  = (SLIDE_W - 2 * _MS_PAD_X - (_MS_COLS - 1) * _MS_COL_GAP) / _MS_COLS
    _MS_AVAIL_H = SLIDE_H - _MS_PAD_Y - _MS_FOOTER_H

    # Collect per-company summary data
    _ms_companies = []
    if not investors.empty and "Company Name" in investors.columns:
        _pri_ms = {"high": 0, "medium": 1, "low": 2}
        _inv_ms = investors.copy()
        _inv_ms["_ps"] = (
            _inv_ms.get("Priority Classification", pd.Series(dtype=str))
            .fillna("").str.strip().str.lower().map(_pri_ms).fillna(3)
        )
        for _, _inv_ms_row in _inv_ms.sort_values("_ps").iterrows():
            _ms_co = str(_inv_ms_row.get("Company Name", "") or "").strip()
            if not _ms_co:
                continue
            _ms_acts = actions[actions["Company Name"] == _ms_co] if not actions.empty and "Company Name" in actions.columns else pd.DataFrame()
            _ms_tot  = len(_ms_acts)
            _ms_done = int(_ms_acts["Status"].str.lower().str.contains("complet").sum()) if not _ms_acts.empty and "Status" in _ms_acts.columns else 0
            _ms_prog = int(_ms_acts["Status"].isin(["In Progress", "Inprogress"]).sum()) if not _ms_acts.empty and "Status" in _ms_acts.columns else 0
            _ms_pend = max(_ms_tot - _ms_done - _ms_prog, 0)
            _ms_pct  = round((_ms_done + _ms_prog) / max(_ms_tot, 1) * 100)

            # Bullet points: prefer flagged actions, fall back to any active
            _ms_bullets = []
            if not _ms_acts.empty:
                _ms_flag = _ms_acts.copy()
                if "To Be In Dashboard" in _ms_flag.columns:
                    _fl = _ms_flag["To Be In Dashboard"].astype(str).str.strip().str.upper()
                    _flagged = _ms_flag[_fl.isin(["YES", "TOP"])]
                    if not _flagged.empty:
                        _ms_flag = _flagged
                for _, _mr in _ms_flag.head(3).iterrows():
                    _rm = str(_mr.get("Remarks", "") or "").strip()
                    _am = str(_mr.get("AM Input", "") or "").strip()
                    _rm = "" if _rm.lower() in ("nan", "none", "-", "n/a") else _rm
                    _am = "" if _am.lower() in ("nan", "none", "-", "n/a") else _am
                    _bt = _am if (_am and len(_am) >= len(_rm) * 0.6) else (_rm or _am)
                    if not _bt:
                        _bt = str(_mr.get("Action Description", "") or "").strip()
                    if not _bt:
                        continue
                    for _sep in (". ", ".\n", "\n"):
                        if _sep in _bt:
                            _bt = _bt.split(_sep)[0]; break
                    if len(_bt) > 90:
                        _bt = _bt[:88].rsplit(" ", 1)[0] + "…"
                    _ms_bullets.append(_bt.strip())
                    if len(_ms_bullets) >= 3:
                        break

            _ms_companies.append((_ms_co, _ms_pct, _ms_done, _ms_prog, _ms_pend, _ms_tot, _ms_bullets))

    # Compute row count and card height
    _ms_n    = len(_ms_companies)
    _ms_rows = max(1, -(-_ms_n // _MS_COLS))   # ceiling division
    _MS_CARD_H = (_MS_AVAIL_H - (_ms_rows - 1) * _MS_ROW_GAP) / max(_ms_rows, 1)

    for _mi, (_ms_co, _ms_pct, _ms_done, _ms_prog, _ms_pend, _ms_tot, _ms_bullets) in enumerate(_ms_companies):
        _mc  = _mi % _MS_COLS
        _mr  = _mi // _MS_COLS
        _cx  = _MS_PAD_X + _mc * (_MS_CARD_W + _MS_COL_GAP)
        _cy  = _MS_PAD_Y + _mr * (_MS_CARD_H + _MS_ROW_GAP)

        # Card background
        _add_rect(_ms, _cx, _cy, _MS_CARD_W, _MS_CARD_H,
                  fill_color=WHITE, line_color=_rgb("#D8D2C0"))

        # Company name header (dark green bar)
        _ms_hdr_h = Inches(0.28)
        _add_rect(_ms, _cx, _cy, _MS_CARD_W, _ms_hdr_h,
                  fill_color=_rgb("#0B4A2F"), line_color=_rgb("#0B4A2F"))
        _add_rect(_ms, _cx, _cy, Inches(0.04), _ms_hdr_h,
                  fill_color=_rgb("#C89B3C"), line_color=_rgb("#C89B3C"))
        _add_text_box(_ms, _ms_co,
                      _cx + Inches(0.10), _cy + Inches(0.04),
                      _MS_CARD_W - Inches(0.14), _ms_hdr_h - Inches(0.06),
                      font_size=13, bold=True, color=WHITE)

        # Progress bar
        _pb_y = _cy + _ms_hdr_h + Inches(0.07)
        _pb_h = Inches(0.13)
        _pb_w = _MS_CARD_W - Inches(0.20)
        _pb_x = _cx + Inches(0.10)
        _add_rect(_ms, _pb_x, _pb_y, _pb_w, _pb_h,
                  fill_color=_rgb("#E0E0DC"), line_color=_rgb("#E0E0DC"))
        if _ms_done > 0 and _ms_tot > 0:
            _add_rect(_ms, _pb_x, _pb_y, _pb_w * _ms_done / _ms_tot, _pb_h,
                      fill_color=GREEN, line_color=GREEN)
        if _ms_prog > 0 and _ms_tot > 0:
            _add_rect(_ms, _pb_x + _pb_w * _ms_done / _ms_tot, _pb_y,
                      _pb_w * _ms_prog / _ms_tot, _pb_h,
                      fill_color=GOLD, line_color=GOLD)
        _add_text_box(_ms, f"{_ms_pct}%",
                      _cx + _MS_CARD_W - Inches(0.48), _pb_y - Inches(0.02),
                      Inches(0.44), Inches(0.18),
                      font_size=8, bold=True, color=_rgb("#0B4A2F"), align=PP_ALIGN.RIGHT)

        # Status counts row
        _sc_y = _pb_y + _pb_h + Inches(0.04)
        _sc_items = [
            (f"✓ {_ms_done} Done",  "#1B5C3F"),
            (f"↺ {_ms_prog} Active","#C9974A"),
            (f"◷ {_ms_pend} Pend",  "#888888"),
        ]
        _sc_w = (_MS_CARD_W - Inches(0.10)) / 3
        for _sci, (_sc_lbl, _sc_col) in enumerate(_sc_items):
            _add_text_box(_ms, _sc_lbl,
                          _cx + Inches(0.05) + _sci * _sc_w, _sc_y,
                          _sc_w, Inches(0.16),
                          font_size=9, bold=True, color=_rgb(_sc_col))

        # Bullet points
        _bl_y = _sc_y + Inches(0.18)
        _bl_avail = _cy + _MS_CARD_H - _bl_y - Inches(0.05)
        _bl_h = _bl_avail / max(len(_ms_bullets), 1) if _ms_bullets else _bl_avail
        for _bi, _bt in enumerate(_ms_bullets):
            _bly = _bl_y + _bi * _bl_h
            if _bly + Inches(0.05) > _cy + _MS_CARD_H:
                break
            _bldot = _ms.shapes.add_shape(9,
                         _cx + Inches(0.09), _bly + Inches(0.06),
                         Inches(0.06), Inches(0.06))
            _bldot.fill.solid(); _bldot.fill.fore_color.rgb = _rgb("#C89B3C")
            _bldot.line.fill.background()
            _add_text_box(_ms, _bt,
                          _cx + Inches(0.19), _bly,
                          _MS_CARD_W - Inches(0.24), _bl_h,
                          font_size=10, color=_rgb("#2B2B2B"))

        if not _ms_bullets:
            _add_text_box(_ms, "No actions recorded.",
                          _cx + Inches(0.10), _bl_y,
                          _MS_CARD_W - Inches(0.14), Inches(0.20),
                          font_size=7, color=_rgb("#888888"))

    # Footer
    _add_rect(_ms, Inches(0), SLIDE_H - _MS_FOOTER_H, SLIDE_W, _MS_FOOTER_H,
              fill_color=_rgb("#0B4A2F"), line_color=_rgb("#0B4A2F"))
    _add_text_box(_ms, "CONFIDENTIAL | Ministry of Investment — وزارة الاستثمار",
                  Inches(0), SLIDE_H - _MS_FOOTER_H, SLIDE_W, _MS_FOOTER_H,
                  font_size=9, color=_rgb("#C89B3C"), align=PP_ALIGN.CENTER)

    # ── Slides 2+: one per company — sorted High → Medium → Low priority
    if not investors.empty and "Company Name" in investors.columns:
        _pri_sort_map = {"high": 0, "medium": 1, "low": 2}
        _inv_sorted = investors.copy()
        _inv_sorted["_ps"] = (
            _inv_sorted.get("Priority Classification", pd.Series(dtype=str))
            .fillna("").str.strip().str.lower()
            .map(_pri_sort_map).fillna(3)
        )
        _inv_sorted = _inv_sorted.sort_values("_ps").drop(columns=["_ps"])
        for _, inv in _inv_sorted.iterrows():
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

    # Also include companies present in action/meeting/opportunity data but absent from Investor Master
    _inv_cos = set(investors["Company Name"].dropna().str.strip().str.lower()) if not investors.empty else set()
    _seen_extra: set = set()
    for _src in [actions, opportunities, meetings]:
        if _src.empty or "Company Name" not in _src.columns:
            continue
        for _co_extra in _src["Company Name"].dropna().unique():
            _co_key = str(_co_extra).strip().lower()
            if not _co_key or _co_key in _inv_cos or _co_key in _seen_extra:
                continue
            _seen_extra.add(_co_key)
            _co_name = str(_co_extra).strip()
            _ea = actions[actions["Company Name"] == _co_extra] if not actions.empty and "Company Name" in actions.columns else pd.DataFrame()
            _eo = opportunities[opportunities["Company Name"] == _co_extra] if not opportunities.empty and "Company Name" in opportunities.columns else pd.DataFrame()
            _em = meetings[meetings["Company Name"] == _co_extra] if not meetings.empty and "Company Name" in meetings.columns else pd.DataFrame()
            try:
                _co_slide_cover_profile(prs, _co_name, pd.Series({"Company Name": _co_name}),
                                        _eo, _ea, _em, pd.DataFrame(), lang,
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
    """Company profile slide — rebuilt to match colleague template (20"×11.25")."""
    slide = _blank_slide(prs)
    _mv = lambda row, col, default="": (row.get(col, default) or default) if not (hasattr(row, "empty") and row.empty) else default

    # ── Slide background: light beige ────────────────────────────────────────
    _fill_background(slide, _rgb("#FAF7EE"))

    # ── Right vertical strip: dark green ─────────────────────────────────────
    _add_rect(slide, Inches(19.375), Inches(0), Inches(0.208), Inches(11.25),
              fill_color=_rgb("#0B4A2F"), line_color=_rgb("#0B4A2F"))

    # ── Header: dark green bar, h=1.354" ─────────────────────────────────────
    HDR_H = Inches(1.354)
    _add_rect(slide, Inches(0), Inches(0), Inches(19.375), HDR_H,
              fill_color=_rgb("#0B4A2F"), line_color=_rgb("#0B4A2F"))
    # Gold accent strip below header
    _add_rect(slide, Inches(0), HDR_H, Inches(19.375), Inches(0.042),
              fill_color=_rgb("#C89B3C"), line_color=_rgb("#C89B3C"))
    # Company name
    _add_text_box(slide, company,
                  Inches(0.417), Inches(0.208), Inches(14.0), Inches(0.90),
                  font_size=33, bold=True, color=WHITE)
    # Breadcrumb
    _add_text_box(slide,
                  "Ministry of Investment  ·  Minister Office  ·  Executive Outreach  ·  Man-marking Weekly Report",
                  Inches(0.417), Inches(1.00), Inches(14.0), Inches(0.26),
                  font_size=10, color=_rgb("#C89B3C"))
    # File owner + date — top right (from Excel header; fallback to default name)
    _file_owner = str(_mv(inv_row, "File Owner", "")).strip() or "Majed H. Al Saadi"
    _add_text_box(slide, f"File owner: {_file_owner}",
                  Inches(14.5), Inches(0.10), Inches(4.60), Inches(0.28),
                  font_size=12, color=WHITE, align=PP_ALIGN.RIGHT)
    _add_text_box(slide, date.today().strftime("%d %B %Y"),
                  Inches(14.5), Inches(0.42), Inches(4.6), Inches(0.35),
                  font_size=11, color=_rgb("#C89B3C"), align=PP_ALIGN.RIGHT)

    # ── Metadata bar: white band with 5 fields + oval icons ──────────────────
    META_Y = Inches(1.396)
    META_H = Inches(0.938)
    _add_rect(slide, Inches(0), META_Y, Inches(19.375), META_H,
              fill_color=WHITE, line_color=WHITE)
    meta_fields = [
        ("COUNTRY",              _mv(inv_row, "Country",              "—"), "⊕"),
        ("SECTOR",               _mv(inv_row, "Sector",               "—"), "≡"),
        ("JOURNEY STAGE",        _mv(inv_row, "Journey Stage",        "—"), "◷"),
        ("RM",                   _mv(inv_row, "Relationship Manager", "—"), "♟"),
        ("AM",                   _mv(inv_row, "Account Manager",      "TBD"), "♟"),
    ]
    _meta_field_w = Inches(19.375) / 5
    for _mi, (_mlbl, _mval, _micon) in enumerate(meta_fields):
        _mx = _mi * _meta_field_w
        # Oval icon
        _icon_r = Inches(0.208)
        _icx = _mx + Inches(0.417)
        _icy = META_Y + META_H / 2
        _icon_shape = slide.shapes.add_shape(9,
            _icx - _icon_r, _icy - _icon_r, _icon_r * 2, _icon_r * 2)
        _icon_shape.fill.solid()
        _icon_shape.fill.fore_color.rgb = _rgb("#0B4A2F")
        _icon_shape.line.fill.background()
        # Icon character centred in the oval
        _add_text_box(slide, _micon,
                      _icx - _icon_r, _icy - _icon_r,
                      _icon_r * 2, _icon_r * 2,
                      font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        # Label
        _add_text_box(slide, _mlbl,
                      _icx + _icon_r + Inches(0.083), META_Y + Inches(0.06),
                      _meta_field_w - _icon_r * 2 - Inches(0.25), Inches(0.32),
                      font_size=20, color=_rgb("#888888"))
        # Value
        _add_text_box(slide, str(_mval)[:30],
                      _icx + _icon_r + Inches(0.083), META_Y + Inches(0.40),
                      _meta_field_w - _icon_r * 2 - Inches(0.25), Inches(0.46),
                      font_size=24, bold=True, color=_rgb("#2B2B2B"))
    # Gold line at bottom of meta bar
    _add_rect(slide, Inches(0), META_Y + META_H, Inches(19.375), Inches(0.021),
              fill_color=_rgb("#C89B3C"), line_color=_rgb("#C89B3C"))

    # ── Compute action stats ──────────────────────────────────────────────────
    n_total  = len(acts)
    n_done_s = int(acts["Status"].str.lower().str.contains("complet").sum()) if not acts.empty and "Status" in acts.columns else 0
    n_prog_s = int(acts["Status"].isin(["In Progress", "Inprogress"]).sum()) if not acts.empty and "Status" in acts.columns else 0
    n_block_s = int(acts["Status"].str.lower().eq("blocked").sum()) if not acts.empty and "Status" in acts.columns else 0
    n_pend_s = max(n_total - n_done_s - n_prog_s - n_block_s, 0)
    n_opps_s = len(opps) if not opps.empty else 0

    if not acts.empty and "Progress" in acts.columns and n_total > 0:
        def _parse_prog(x):
            try:
                v = float(str(x).replace("%", "")) if x not in (None, "", "nan") else 0.0
                return v / 100.0 if v > 1.0 else v
            except Exception:
                return 0.0
        _prog_series = acts.apply(
            lambda r: 1.0 if str(r.get("Status", "")).lower() in ("completed", "cancelled")
            else _parse_prog(r.get("Progress", 0)), axis=1)
        pct_s = round(_prog_series.mean() * 100)
    else:
        pct_s = round((n_done_s + n_prog_s) / n_total * 100) if n_total > 0 else 0

    # Strategic brief
    goal_text, strat_text, action_text = _build_strategic_brief(acts, opps, meetings, inv_row)

    # ── 3 KPI cards — full slide width ───────────────────────────────────────
    KPI_Y    = Inches(2.374)
    KPI_H    = Inches(1.563)
    KPI_GAP  = Inches(0.208)
    _KPI_TOT = Inches(19.375) - Inches(0.417) - Inches(0.208)
    KPI_W    = (_KPI_TOT - 2 * KPI_GAP) / 3    # ~6.11" each

    def _kpi_sq_icon(sx, sy, size, hex_col, char):
        sq = slide.shapes.add_shape(5, sx, sy, size, size)
        sq.fill.solid(); sq.fill.fore_color.rgb = _rgb(hex_col)
        sq.line.fill.background()
        _add_text_box(slide, char, sx, sy, size, size,
                      font_size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    _SQ = Inches(0.65)

    # Card 1: Strategic Goal (dark green)
    C1_X = Inches(0.417)
    _add_rect(slide, C1_X, KPI_Y, KPI_W, KPI_H,
              fill_color=WHITE, line_color=_rgb("#E0E0DC"))
    _add_rect(slide, C1_X, KPI_Y, KPI_W, Inches(0.042),
              fill_color=_rgb("#0B4A2F"), line_color=_rgb("#0B4A2F"))
    _kpi_sq_icon(C1_X + Inches(0.22), KPI_Y + Inches(0.44), _SQ, "#0B4A2F", "◎")
    _add_text_box(slide, "STRATEGIC GOAL",
                  C1_X + Inches(1.05), KPI_Y + Inches(0.08),
                  KPI_W - Inches(1.20), Inches(0.40),
                  font_size=24, bold=True, color=_rgb("#0B4A2F"))
    _add_text_box(slide, goal_text[:220],
                  C1_X + Inches(1.05), KPI_Y + Inches(0.52),
                  KPI_W - Inches(1.20), Inches(0.95),
                  font_size=15, color=_rgb("#2B2B2B"))

    # Card 2: Overall Progress (gold background) — big %, subtitle, mini bar
    C2_X = C1_X + KPI_W + KPI_GAP
    _add_rect(slide, C2_X, KPI_Y, KPI_W, KPI_H,
              fill_color=_rgb("#C89B3C"), line_color=_rgb("#C89B3C"))
    _kpi_sq_icon(C2_X + Inches(0.22), KPI_Y + Inches(0.18), _SQ, "#A07828", "↗")
    _add_text_box(slide, "OVERALL PROGRESS",
                  C2_X + Inches(1.05), KPI_Y + Inches(0.08),
                  KPI_W - Inches(1.20), Inches(0.40),
                  font_size=24, bold=True, color=WHITE)
    _add_text_box(slide, f"{pct_s}%",
                  C2_X + Inches(1.00), KPI_Y + Inches(0.28),
                  Inches(2.60), Inches(0.90),
                  font_size=52, bold=True, color=WHITE)
    _add_text_box(slide, f"{n_done_s + n_prog_s} of {n_total} actions active",
                  C2_X + Inches(1.00), KPI_Y + Inches(0.98),
                  KPI_W - Inches(1.20), Inches(0.26),
                  font_size=16, color=WHITE)
    _pb_x = C2_X + Inches(1.00)
    _pb_y = KPI_Y + KPI_H - Inches(0.23)
    _pb_w = KPI_W - Inches(1.20)
    _add_rect(slide, _pb_x, _pb_y, _pb_w, Inches(0.07),
              fill_color=_rgb("#A07828"), line_color=_rgb("#A07828"))
    _add_rect(slide, _pb_x, _pb_y, max(_pb_w * pct_s / 100, Inches(0.02)), Inches(0.07),
              fill_color=WHITE, line_color=WHITE)

    # Card 3: Immediate Action (red) — ALL CAPS text + Priority badge top-right
    C3_X = C2_X + KPI_W + KPI_GAP
    _add_rect(slide, C3_X, KPI_Y, KPI_W, KPI_H,
              fill_color=WHITE, line_color=_rgb("#E0E0DC"))
    _add_rect(slide, C3_X, KPI_Y, KPI_W, Inches(0.042),
              fill_color=_rgb("#C0392B"), line_color=_rgb("#C0392B"))
    _kpi_sq_icon(C3_X + Inches(0.22), KPI_Y + Inches(0.44), _SQ, "#C0392B", "⚠")
    _add_text_box(slide, "IMMEDIATE ACTION",
                  C3_X + Inches(1.05), KPI_Y + Inches(0.08),
                  KPI_W - Inches(2.80), Inches(0.40),
                  font_size=24, bold=True, color=_rgb("#C0392B"))
    # Priority badge — top right (uses Priority Classification from investor row)
    _pc_val = str(_mv(inv_row, "Priority Classification", "")).strip()
    _pc_colors = {"high": "#C00000", "medium": "#2D6DA8", "low": "#888888"}
    _pc_key = _pc_val.lower()
    if _pc_key not in _pc_colors:
        _pc_key = "low"
        _pc_val = "Low"
    _pc_col = _rgb(_pc_colors[_pc_key])
    _add_rect(slide, C3_X + KPI_W - Inches(1.30), KPI_Y + Inches(0.07),
              Inches(1.15), Inches(0.32),
              fill_color=_pc_col, line_color=_pc_col)
    _add_text_box(slide, "Priority",
                  C3_X + KPI_W - Inches(1.30), KPI_Y + Inches(0.07),
                  Inches(1.15), Inches(0.15),
                  font_size=6, color=WHITE, align=PP_ALIGN.CENTER)
    _add_text_box(slide, _pc_val.upper(),
                  C3_X + KPI_W - Inches(1.30), KPI_Y + Inches(0.19),
                  Inches(1.15), Inches(0.18),
                  font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    _add_text_box(slide, (action_text or "No immediate action required.").upper()[:200],
                  C3_X + Inches(1.10), KPI_Y + Inches(0.48),
                  KPI_W - Inches(1.25), Inches(0.95),
                  font_size=15, bold=True, color=_rgb("#2B2B2B"))

    # ── Journey Timeline: y=4.145", w=8.854" — meetings as milestones ────────
    TL_X = Inches(0.417)
    TL_Y = Inches(4.145)
    TL_W = Inches(8.854)
    TL_H = Inches(1.771)
    _add_rect(slide, TL_X, TL_Y, TL_W, TL_H,
              fill_color=WHITE, line_color=_rgb("#E0E0DC"))
    _add_rect(slide, TL_X, TL_Y, TL_W, Inches(0.042),
              fill_color=_rgb("#C89B3C"), line_color=_rgb("#C89B3C"))
    _add_text_box(slide, "JOURNEY TIMELINE",
                  TL_X + Inches(0.15), TL_Y + Inches(0.06),
                  Inches(3.5), Inches(0.35),
                  font_size=22, bold=True, color=_rgb("#0B4A2F"))

    AX_Y = TL_Y + Inches(1.05)
    _add_rect(slide, TL_X + Inches(0.20), AX_Y, TL_W - Inches(0.40), Inches(0.015),
              fill_color=_rgb("#CCCCCC"), line_color=_rgb("#CCCCCC"))

    try:
        # Build unified date list: meetings first, then key action due dates
        _tl_events = []   # (date, label, sub_label, color)
        _MTG_COL = {"Completed": "#1B5C3F", "Scheduled": "#C89B3C",
                    "Cancelled": "#AAAAAA", "In Progress": "#C89B3C"}
        if not meetings.empty and "Meeting Date" in meetings.columns:
            _mtg_copy = meetings.copy()
            _mtg_copy["_dt"] = pd.to_datetime(_mtg_copy["Meeting Date"], errors="coerce")
            _mtg_copy = _mtg_copy[_mtg_copy["_dt"].notna()].sort_values("_dt")
            for _, _mr in _mtg_copy.iterrows():
                _mt = str(_mr.get("Meeting Type", "") or "Meeting").strip()
                _md = str(_mr.get("Meeting Description", _mr.get("Agenda", "")) or "").strip()
                _mc = _MTG_COL.get(_mt, "#1B5C3F")
                _tl_events.append((_mr["_dt"].date(), _mt, _md[:22], _mc))

        if not _tl_events and not acts.empty and "Due Date" in acts.columns:
            _act_copy = acts.copy()
            _act_copy["_dt"] = pd.to_datetime(_act_copy["Due Date"], errors="coerce")
            _act_copy = _act_copy[_act_copy["_dt"].notna()].sort_values("_dt")
            for _, _ar in _act_copy.iterrows():
                _tl_events.append((_ar["_dt"].date(),
                                   str(_ar.get("Status", "Action"))[:16],
                                   str(_ar.get("Action Description","") or "")[:22],
                                   "#C89B3C"))

        if _tl_events:
            _today = date.today()
            _all_dates = [e[0] for e in _tl_events]
            _vt_min = min(_all_dates).replace(day=1)
            _vt_max_raw = max(_all_dates)
            _ld = _calendar.monthrange(_vt_max_raw.year, _vt_max_raw.month)[1]
            _vt_max = _vt_max_raw.replace(day=_ld)
            if _today < _vt_min: _vt_min = _today.replace(day=1)
            if _today > _vt_max:
                _ld2 = _calendar.monthrange(_today.year, _today.month)[1]
                _vt_max = _today.replace(day=_ld2)
            _span = max((_vt_max - _vt_min).days, 1)
            _TL_AX_X = TL_X + Inches(0.20)
            _TL_AX_W = TL_W - Inches(0.40)

            def _tlx(d):
                frac = max(0.0, min(1.0, (d - _vt_min).days / _span))
                return _TL_AX_X + frac * _TL_AX_W

            # Month labels along axis
            _n_months = max(1, (_vt_max.year - _vt_min.year) * 12
                            + (_vt_max.month - _vt_min.month))
            _show_every = 1 if _n_months <= 8 else 2
            _mo = _vt_min.replace(day=1)
            _mo_count = 0
            while _mo <= _vt_max and _mo_count < 36:
                _mo_count += 1
                _mx = _tlx(_mo)
                _add_rect(slide, _mx, AX_Y, Inches(0.010), Inches(0.09),
                          fill_color=_rgb("#BBBBBB"), line_color=_rgb("#BBBBBB"))
                if _mo_count % _show_every == 0:
                    _lbl = _mo.strftime("%b '%y") if _n_months > 12 else _mo.strftime("%b")
                    _add_text_box(slide, _lbl,
                                  _mx - Inches(0.40), AX_Y + Inches(0.12),
                                  Inches(0.80), Inches(0.28),
                                  font_size=18, bold=True, color=_rgb("#333333"),
                                  align=PP_ALIGN.CENTER)
                _mo = (_mo.replace(year=_mo.year + 1, month=1)
                       if _mo.month == 12 else _mo.replace(month=_mo.month + 1))

            # NOW marker
            _nx = _tlx(_today)
            _add_rect(slide, _nx - Inches(0.018), AX_Y - Inches(0.28),
                      Inches(0.036), Inches(0.55),
                      fill_color=_rgb(MISA_GOLD), line_color=_rgb(MISA_GOLD))
            # NOW badge pill
            _add_rect(slide, _nx - Inches(0.28), AX_Y - Inches(0.52),
                      Inches(0.56), Inches(0.22),
                      fill_color=_rgb(MISA_GOLD), line_color=_rgb(MISA_GOLD))
            _add_text_box(slide, "NOW", _nx - Inches(0.28), AX_Y - Inches(0.52),
                          Inches(0.56), Inches(0.22),
                          font_size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

            # Milestone dots + labels
            for _ei, (_edate, _elbl, _esub, _ecol) in enumerate(_tl_events):
                _ex = _tlx(_edate)
                _dot_r = Inches(0.10)
                # Alternate above/below axis for crowded dots
                _above = (_ei % 2 == 0)
                _dot_y = AX_Y - _dot_r if _above else AX_Y + Inches(0.015)
                _dot = slide.shapes.add_shape(9, _ex - _dot_r, _dot_y, _dot_r * 2, _dot_r * 2)
                _dot.fill.solid(); _dot.fill.fore_color.rgb = _rgb(_ecol)
                _dot.line.fill.background()
                # Short description clipped tight to axis line
                if _esub:
                    _lbl_y = (AX_Y - Inches(0.25) if _above else AX_Y + Inches(0.22))
                    _add_text_box(slide, _esub,
                                  _ex - Inches(0.65), _lbl_y,
                                  Inches(1.30), Inches(0.15),
                                  font_size=6, color=_rgb(_ecol), align=PP_ALIGN.CENTER)
    except Exception:
        pass

    # ── Actions Summary panel: y=4.145", x=9.479" — 5 count circles ──────────
    AS_X    = Inches(9.479)
    AS_Y    = Inches(4.145)
    AS_W    = Inches(9.479)  # remaining width: 19.375 - 9.479 - 0.208 gap = 9.688
    AS_W    = Inches(9.688)
    AS_H    = Inches(1.771)
    _add_rect(slide, AS_X, AS_Y, AS_W, AS_H,
              fill_color=WHITE, line_color=_rgb("#E0E0DC"))
    _add_rect(slide, AS_X, AS_Y, AS_W, Inches(0.042),
              fill_color=_rgb("#0B4A2F"), line_color=_rgb("#0B4A2F"))
    _add_text_box(slide, "Actions Summary",
                  AS_X + Inches(0.15), AS_Y + Inches(0.01),
                  Inches(3.5), Inches(0.35),
                  font_size=18, bold=True, color=_rgb("#0B4A2F"))

    # 4 count circles (Blocked removed)
    _circ_items = [
        (str(n_done_s),  "Completed",  "#1B5C3F", "✓"),
        (str(n_prog_s),  "In Progress","#C9974A", "↺"),
        (str(n_pend_s),  "Pending",    "#888888", "◷"),
        (str(n_opps_s),  "Opps",       "#0B4A2F", "★"),
    ]
    _circ_r   = Inches(0.60)
    _circ_gap = (AS_W - Inches(0.30)) / 4
    for _ci_idx, (_cval, _clbl, _ccol, _cicon) in enumerate(_circ_items):
        _cx = AS_X + Inches(0.15) + _circ_gap * _ci_idx + _circ_gap / 2
        _cy = AS_Y + Inches(0.90)
        _cs = slide.shapes.add_shape(9, _cx - _circ_r, _cy - _circ_r, _circ_r*2, _circ_r*2)
        _cs.fill.solid(); _cs.fill.fore_color.rgb = _rgb(_ccol); _cs.line.fill.background()
        # Icon in upper half of circle
        _add_text_box(slide, _cicon,
                      _cx - _circ_r, _cy - _circ_r,
                      _circ_r*2, _circ_r,
                      font_size=30, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        # Count in lower half
        _add_text_box(slide, _cval,
                      _cx - _circ_r, _cy - Inches(0.05),
                      _circ_r*2, _circ_r,
                      font_size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _add_text_box(slide, _clbl,
                      _cx - _circ_r - Inches(0.10), _cy + _circ_r + Inches(0.02),
                      _circ_r*2 + Inches(0.20), Inches(0.30),
                      font_size=16, color=_rgb("#555555"), align=PP_ALIGN.CENTER)

    # ── Bottom section: y=6.437" — Opportunities (left) + Actions table (right)
    BOT_Y = Inches(6.437)
    BOT_H = Inches(10.833) - BOT_Y - Inches(0.417)  # up to footer

    # ── Opportunities panel: left=0.417", w=3.750" ───────────────────────────
    OP_X = Inches(0.417)
    OP_W = Inches(3.750)
    _add_rect(slide, OP_X, BOT_Y, OP_W, BOT_H,
              fill_color=_rgb("#F9F5EA"), line_color=_rgb("#E0D8CC"))
    _add_rect(slide, OP_X, BOT_Y, OP_W, Inches(0.042),
              fill_color=_rgb("#C89B3C"), line_color=_rgb("#C89B3C"))
    # Count active opps for the header
    _n_active_opps = (
        int((opps["Opportunity Status"] == "Active").sum())
        if not opps.empty and "Opportunity Status" in opps.columns
        else n_opps_s
    )
    _add_text_box(slide, f"OPPORTUNITIES ({_n_active_opps} ACTIVE)",
                  OP_X + Inches(0.12), BOT_Y + Inches(0.05),
                  OP_W - Inches(0.20), Inches(0.26),
                  font_size=14, bold=True, color=_rgb("#0B4A2F"))

    def _get_sector_icon(sec):
        _si = {
            "water": "≋", "utilities": "⚡", "energy": "⚡",
            "aviation": "✈", "transport": "✈", "logistics": "✈",
            "tourism": "⌂", "hospitality": "⌂", "hotel": "⌂",
            "finance": "$", "banking": "$", "investment": "$",
            "tech": "⚙", "digital": "⚙", "technology": "⚙",
            "mining": "◆", "resources": "◆", "metal": "◆",
            "construction": "⊞", "real estate": "⊞",
            "health": "✚", "pharma": "✚", "medical": "✚",
            "food": "❖", "agriculture": "❖",
        }
        sl = str(sec).lower()
        for k, v in _si.items():
            if k in sl:
                return v
        return "◎"

    def _stage_badge_col(st):
        sl = str(st).lower()
        if any(x in sl for x in ("contract", "signed", "complet")):
            return "#1B5C3F"
        if any(x in sl for x in ("negotiat", "advanced")):
            return "#C9974A"
        if any(x in sl for x in ("due dilig", "review")):
            return "#2D6DA8"
        return "#888888"

    _inv_sector_val = str(inv_row.get("Sector", "") or "")
    _opp_item_y = BOT_Y + Inches(0.40)
    _opp_card_h = Inches(0.68)
    if not opps.empty and "Opportunity Name" in opps.columns:
        for _, _orow in opps.iterrows():
            if _opp_item_y + _opp_card_h > BOT_Y + BOT_H - Inches(0.05):
                break
            _oname  = str(_orow.get("Opportunity Name", "") or "").strip()
            _ostage = str(_orow.get("Opportunity Stage", "") or "").strip()
            _osub   = str(_orow.get("Description", "") or _orow.get("Partners", "") or "").strip()
            if _oname in ("nan", ""):
                continue
            # Card background
            _add_rect(slide, OP_X + Inches(0.08), _opp_item_y,
                      OP_W - Inches(0.16), _opp_card_h,
                      fill_color=WHITE, line_color=_rgb("#E0D8CC"))
            # Sector icon circle
            _sicon = _get_sector_icon(_inv_sector_val)
            _sic_sz = Inches(0.46)
            _sic_x  = OP_X + Inches(0.12)
            _sic_y  = _opp_item_y + (_opp_card_h - _sic_sz) / 2
            _sic = slide.shapes.add_shape(9, _sic_x, _sic_y, _sic_sz, _sic_sz)
            _sic.fill.solid(); _sic.fill.fore_color.rgb = _rgb("#EEE8D5"); _sic.line.fill.background()
            _add_text_box(slide, _sicon, _sic_x, _sic_y, _sic_sz, _sic_sz,
                          font_size=14, bold=True, color=_rgb("#0B4A2F"), align=PP_ALIGN.CENTER)
            # Opportunity name
            _txt_x = OP_X + Inches(0.65)
            _txt_w = OP_W - Inches(0.78)
            _add_text_box(slide, _oname[:36],
                          _txt_x, _opp_item_y + Inches(0.04),
                          _txt_w, Inches(0.32),
                          font_size=14, bold=True, color=_rgb("#2B2B2B"))
            # Subtitle (description / partners)
            if _osub and _osub not in ("nan", ""):
                _add_text_box(slide, _osub[:42],
                              _txt_x, _opp_item_y + Inches(0.36),
                              _txt_w, Inches(0.26),
                              font_size=16, color=_rgb("#888888"))
            _opp_item_y += _opp_card_h + Inches(0.06)
    else:
        _add_text_box(slide, "No opportunities recorded.",
                      OP_X + Inches(0.12), _opp_item_y,
                      OP_W - Inches(0.20), Inches(0.26),
                      font_size=8.5, color=_rgb("#888888"))
        _opp_item_y += Inches(0.32)

    # ── Major Highlights — below last opportunity card ─────────────────────────
    _mh_y = _opp_item_y
    _mh_avail = BOT_Y + BOT_H - _mh_y - Inches(0.05)
    if _mh_avail > Inches(0.65):
        _MH_TTL_H = Inches(0.28)
        _add_rect(slide, OP_X, _mh_y, OP_W, _MH_TTL_H,
                  fill_color=_rgb("#0B4A2F"), line_color=_rgb("#0B4A2F"))
        _add_rect(slide, OP_X, _mh_y, Inches(0.042), _MH_TTL_H,
                  fill_color=_rgb("#C89B3C"), line_color=_rgb("#C89B3C"))
        _add_text_box(slide, "MAJOR HIGHLIGHTS",
                      OP_X + Inches(0.10), _mh_y + Inches(0.05),
                      OP_W - Inches(0.16), _MH_TTL_H - Inches(0.08),
                      font_size=12, bold=True, color=_rgb("#C89B3C"))
        _mh_body_y = _mh_y + _MH_TTL_H + Inches(0.06)
        # Pull top 2 highlights from Remarks/AM Input (dashboard-flagged actions first)
        _mh_bullets = []
        _mh_src = acts.copy() if not acts.empty else pd.DataFrame()
        if not _mh_src.empty and "To Be In Dashboard" in _mh_src.columns:
            _mh_flagged = _mh_src[
                _mh_src["To Be In Dashboard"].astype(str).str.strip().str.upper().isin(["YES", "TOP"])
            ]
            if not _mh_flagged.empty:
                _mh_src = _mh_flagged
        if not _mh_src.empty:
            for _, _mhr in _mh_src.head(5).iterrows():
                _r = str(_mhr.get("Remarks", "") or "").strip()
                _a = str(_mhr.get("AM Input", "") or "").strip()
                _r = "" if _r.lower() in ("nan", "none", "-", "n/a") else _r
                _a = "" if _a.lower() in ("nan", "none", "-", "n/a") else _a
                _mh_txt = _a if (_a and len(_a) >= len(_r) * 0.6) else (_r or _a)
                if not _mh_txt:
                    continue
                for _sep in (". ", ".\n", "\n"):
                    if _sep in _mh_txt:
                        _mh_txt = _mh_txt.split(_sep)[0]
                        break
                if len(_mh_txt) > 100:
                    _mh_txt = _mh_txt[:99].rsplit(" ", 1)[0] + "…"
                _mh_bullets.append(_mh_txt.strip())
                if len(_mh_bullets) >= 2:
                    break
        if not _mh_bullets:
            _notes_val = str(inv_row.get("Notes", "") or "").strip()
            if _notes_val and _notes_val.lower() not in ("nan", "none", ""):
                _mh_bullets = [_notes_val[:100]]
        for _mh_b in _mh_bullets:
            if _mh_body_y + Inches(0.30) > BOT_Y + BOT_H - Inches(0.05):
                break
            _mhdot = slide.shapes.add_shape(
                9, OP_X + Inches(0.10), _mh_body_y + Inches(0.09),
                Inches(0.07), Inches(0.07))
            _mhdot.fill.solid()
            _mhdot.fill.fore_color.rgb = _rgb("#C89B3C")
            _mhdot.line.fill.background()
            _add_text_box(slide, _mh_b,
                          OP_X + Inches(0.24), _mh_body_y,
                          OP_W - Inches(0.32), Inches(0.52),
                          font_size=11, color=_rgb("#2B2B2B"))
            _mh_body_y += Inches(0.56)

    # ── Actions table: left=4.167", y=BOT_Y ──────────────────────────────────
    TBL_X = Inches(4.167)
    TBL_W = Inches(19.375) - TBL_X - Inches(0.10)
    # Column widths: # | Action | Owner | Start | End | Update | Progress
    # Priority removed (1.55"); its width added to Update → 4.95"
    _cw_raw = [0.45, 4.20, 1.65, 1.10, 1.10, 4.95, 1.60]
    _cw_sum = sum(_cw_raw)
    _scale = TBL_W / Inches(_cw_sum)
    CW_TBL = [Inches(w) * _scale for w in _cw_raw]
    COL_HDR = ["#", "Action", "Owner", "Start Date", "End Date", "Update", "Progress"]
    HDR_H2  = Inches(0.333)
    ROW_H2  = Inches(0.375)

    def _co_update_text(remarks, am_input, max_chars=55):
        r = str(remarks or "").strip()
        a = str(am_input or "").strip()
        r = "" if r.lower() in ("nan", "none", "-", "n/a") else r
        a = "" if a.lower() in ("nan", "none", "-", "n/a") else a
        text = a if (a and a.lower() != r.lower() and len(a) >= len(r) * 0.6) else (r or a)
        if not text:
            return "—"
        for sep in (". ", ".\n", "\n"):
            if sep in text:
                text = text.split(sep)[0]
                break
        if len(text) > max_chars:
            text = text[:max_chars - 1].rsplit(" ", 1)[0] + "…"
        return text.strip()

    # Section title "PENDING / IN PROGRESS ACTIONS"
    _TBL_TTL_H = Inches(0.35)
    _add_rect(slide, TBL_X, BOT_Y, TBL_W, _TBL_TTL_H,
              fill_color=_rgb("#EEE8D5"), line_color=_rgb("#D8D2C0"))
    _add_text_box(slide, "PENDING / IN PROGRESS ACTIONS",
                  TBL_X + Inches(0.12), BOT_Y + Inches(0.03),
                  TBL_W - Inches(0.20), _TBL_TTL_H - Inches(0.04),
                  font_size=16, bold=True, color=_rgb("#0B4A2F"))
    _TBL_START_Y = BOT_Y + _TBL_TTL_H

    # Table header
    _cx_cur = TBL_X
    for _hdr, _cw in zip(COL_HDR, CW_TBL):
        _add_rect(slide, _cx_cur, _TBL_START_Y, _cw, HDR_H2,
                  fill_color=_rgb("#0B4A2F"), line_color=_rgb("#0B4A2F"))
        _add_text_box(slide, _hdr, _cx_cur + Inches(0.04), _TBL_START_Y + Inches(0.04),
                      _cw - Inches(0.08), HDR_H2 - Inches(0.04),
                      font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _cx_cur += _cw

    # Sort actions: pending first by due date, then done
    def _fmt_dt2(v):
        try:
            d = pd.to_datetime(v, errors="coerce")
            return d.strftime("%d %b") if pd.notna(d) else ""
        except Exception:
            return ""

    if not acts.empty and "Status" in acts.columns:
        _pend2 = acts[~acts["Status"].isin(["Completed","Cancelled"])].copy()
        _done2 = acts[acts["Status"].isin(["Completed","Cancelled"])].copy()
        if "Due Date" in _pend2.columns:
            _pend2["_due"] = pd.to_datetime(_pend2["Due Date"], errors="coerce")
            _pend2 = _pend2.sort_values("_due")
        _acts_ordered = pd.concat([_pend2, _done2], ignore_index=True)
    else:
        _acts_ordered = acts.copy() if not acts.empty else pd.DataFrame()

    # Filter to rows flagged "YES" or "TOP" for the dashboard if the column exists
    if not _acts_ordered.empty and "To Be In Dashboard" in _acts_ordered.columns:
        _dash_mask = _acts_ordered["To Be In Dashboard"].astype(str).str.strip().str.upper().isin(["YES", "TOP"])
        if _dash_mask.any():
            _acts_ordered = _acts_ordered[_dash_mask].reset_index(drop=True)

    _pri_colors = {"High": "#C00000", "Very High": "#C00000",
                   "Medium": "#FFC000", "Low": "#0B4A2F"}

    _row_y = _TBL_START_Y + HDR_H2
    _row_i = 0
    for _, _rrow in _acts_ordered.iterrows():
        if _row_y + ROW_H2 > Inches(10.833) - Inches(0.417):
            break
        _alt = _rgb("#FAF7EE") if _row_i % 2 == 1 else WHITE
        _row_status = str(_rrow.get("Status", "") or "")
        _is_done2 = _row_status.lower() in ("completed", "cancelled")

        try:
            _prog_raw = _rrow.get("Progress", 0)
            _prog_v = float(str(_prog_raw).replace("%","")) if _prog_raw not in (None,"","nan") else 0.0
            if _prog_v > 1.0: _prog_v /= 100.0
        except Exception:
            _prog_v = 0.0
        if _is_done2: _prog_v = 1.0

        _cx_cur = TBL_X
        # Col 0: #
        _add_rect(slide, _cx_cur, _row_y, CW_TBL[0], ROW_H2, fill_color=_alt, line_color=_rgb("#E0E0DC"))
        _add_text_box(slide, str(_row_i+1), _cx_cur + Inches(0.02), _row_y + Inches(0.07),
                      CW_TBL[0] - Inches(0.04), Inches(0.26),
                      font_size=13, color=_rgb("#2B2B2B"), align=PP_ALIGN.CENTER)
        _cx_cur += CW_TBL[0]
        # Col 1: Action
        _add_rect(slide, _cx_cur, _row_y, CW_TBL[1], ROW_H2, fill_color=_alt, line_color=_rgb("#E0E0DC"))
        _desc2 = str(_rrow.get("Action Description","") or "")
        _add_text_box(slide, _desc2[:80], _cx_cur + Inches(0.05), _row_y + Inches(0.04),
                      CW_TBL[1] - Inches(0.08), ROW_H2 - Inches(0.06),
                      font_size=10, color=_rgb("#2B2B2B"))
        _cx_cur += CW_TBL[1]
        # Col 2: Owner
        _add_rect(slide, _cx_cur, _row_y, CW_TBL[2], ROW_H2, fill_color=_alt, line_color=_rgb("#E0E0DC"))
        _owner2 = str(_rrow.get("Assigned To","") or "")[:22]
        _add_text_box(slide, _owner2, _cx_cur + Inches(0.04), _row_y + Inches(0.07),
                      CW_TBL[2] - Inches(0.06), Inches(0.26),
                      font_size=10, color=_rgb("#2B2B2B"))
        _cx_cur += CW_TBL[2]
        # Col 3: Start Date
        _add_rect(slide, _cx_cur, _row_y, CW_TBL[3], ROW_H2, fill_color=_alt, line_color=_rgb("#E0E0DC"))
        _sdate = _fmt_dt2(_rrow.get("Start Date", _rrow.get("Created Date", None)))
        _add_text_box(slide, _sdate, _cx_cur + Inches(0.04), _row_y + Inches(0.07),
                      CW_TBL[3] - Inches(0.06), Inches(0.26),
                      font_size=13, color=_rgb("#2B2B2B"), align=PP_ALIGN.CENTER)
        _cx_cur += CW_TBL[3]
        # Col 4: End Date
        _add_rect(slide, _cx_cur, _row_y, CW_TBL[4], ROW_H2, fill_color=_alt, line_color=_rgb("#E0E0DC"))
        _edate = _fmt_dt2(_rrow.get("Due Date", None))
        _add_text_box(slide, _edate, _cx_cur + Inches(0.04), _row_y + Inches(0.07),
                      CW_TBL[4] - Inches(0.06), Inches(0.26),
                      font_size=13, color=_rgb("#2B2B2B"), align=PP_ALIGN.CENTER)
        _cx_cur += CW_TBL[4]
        # Col 5: Update (intelligent summary of Remarks + AM Input)
        _add_rect(slide, _cx_cur, _row_y, CW_TBL[5], ROW_H2, fill_color=_alt, line_color=_rgb("#E0E0DC"))
        _upd2 = _co_update_text(_rrow.get("Remarks", ""), _rrow.get("AM Input", ""))
        _add_text_box(slide, _upd2,
                      _cx_cur + Inches(0.04), _row_y + Inches(0.04),
                      CW_TBL[5] - Inches(0.06), ROW_H2 - Inches(0.05),
                      font_size=13, color=_rgb("#444444"))
        _cx_cur += CW_TBL[5]
        # Col 6: Progress %
        _add_rect(slide, _cx_cur, _row_y, CW_TBL[6], ROW_H2, fill_color=_alt, line_color=_rgb("#E0E0DC"))
        _bx2 = _cx_cur + Inches(0.05)
        _bw2 = CW_TBL[6] - Inches(0.10)
        _bh2 = Inches(0.10)
        _by2 = _row_y + ROW_H2/2 - _bh2/2 - Inches(0.05)
        _add_rect(slide, _bx2, _by2, _bw2, _bh2, fill_color=_rgb("#E0E0DC"), line_color=_rgb("#E0E0DC"))
        if _prog_v > 0:
            _fc3 = (_rgb("#1B5C3F") if _prog_v >= 1.0
                    else (_rgb("#C9974A") if _prog_v >= 0.5 else _rgb("#888888")))
            _add_rect(slide, _bx2, _by2, max(_bw2 * _prog_v, Inches(0.02)), _bh2,
                      fill_color=_fc3, line_color=_fc3)
        _add_text_box(slide, f"{int(_prog_v*100)}%",
                      _cx_cur, _by2 + _bh2 + Inches(0.01),
                      CW_TBL[6], Inches(0.22),
                      font_size=13, color=_rgb("#2B2B2B"), align=PP_ALIGN.CENTER)

        _row_y += ROW_H2
        _row_i += 1

    if not _acts_ordered.empty and _row_i == 0:
        _add_text_box(slide, "No action items recorded.",
                      TBL_X + Inches(0.10), _TBL_START_Y + HDR_H2 + Inches(0.10),
                      TBL_W, Inches(0.26),
                      font_size=11, color=_rgb("#888888"))

    # ── Footer ────────────────────────────────────────────────────────────────
    _add_rect(slide, Inches(0), Inches(10.833), Inches(19.375), Inches(0.417),
              fill_color=_rgb("#0B4A2F"), line_color=_rgb("#0B4A2F"))
    _add_text_box(slide, "CONFIDENTIAL | Ministry of Investment — وزارة الاستثمار",
                  Inches(0), Inches(10.833), Inches(19.375), Inches(0.417),
                  font_size=11, color=_rgb("#C89B3C"), align=PP_ALIGN.CENTER)

def _co_slide_opps_deals(prs, company, inv_row, opps, deals, acts, lang):
    """Slide 2 — Minister-grade opportunity cards, full-width 2-column grid."""
    slide = _blank_slide(prs)

    # ── Green header band
    _add_rect(slide, Inches(0), Inches(0), Inches(20.0), Inches(0.9),
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

    # ── Footer
    _add_rect(slide, Inches(0), Inches(10.833), Inches(20.0), Inches(0.417),
              fill_color=_rgb("#0B4A2F"), line_color=_rgb("#0B4A2F"))
    _add_text_box(slide, "CONFIDENTIAL | Ministry of Investment — وزارة الاستثمار",
                  Inches(0), Inches(10.833), Inches(20.0), Inches(0.417),
                  font_size=11, color=_rgb("#C89B3C"), align=PP_ALIGN.CENTER)


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

    # If any action is flagged TOP, use it directly as the IMMEDIATE ACTION text
    _top_action_text = ""
    if not acts.empty and "To Be In Dashboard" in acts.columns:
        _top_acts = acts[acts["To Be In Dashboard"].astype(str).str.strip().str.upper() == "TOP"]
        if not _top_acts.empty:
            _top_desc = str(_top_acts.iloc[0].get("Action Description", "")).strip()
            if _top_desc and _top_desc not in ("nan", ""):
                _top_action_text = _top_desc

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
    # "TOP" flagged action overrides everything else
    if _top_action_text:
        m_act = _top_action_text
    elif minister_act and minister_act not in ("None Required", "—", ""):
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

