
# PDF executive summary generator — clean single-document format for the Minister.

import io
from datetime import date
import pandas as pd

from config.settings import MISA_GREEN, MISA_GOLD, STATUS_COLORS, TIER_COLORS
from config.translations import t


def generate_pdf(dfs: dict, lang: str = "en") -> bytes:
    """
    Generate a PDF executive summary using reportlab.
    Falls back to a simple text-based PDF if charts fail.
    """
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import cm
        from reportlab.lib.colors import HexColor, white, black
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
            HRFlowable, KeepTogether,
        )
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

        investors     = dfs.get("Investor Master",    pd.DataFrame())
        meetings      = dfs.get("Meeting Log",        pd.DataFrame())
        opportunities = dfs.get("Opportunity Pipeline", pd.DataFrame())
        actions       = dfs.get("Action Items",       pd.DataFrame())

        buf = io.BytesIO()
        doc = SimpleDocTemplate(
            buf, pagesize=A4,
            rightMargin=2*cm, leftMargin=2*cm,
            topMargin=2*cm, bottomMargin=2*cm,
        )

        GREEN  = HexColor(MISA_GREEN)
        GOLD   = HexColor(MISA_GOLD)
        LGRAY  = HexColor("#F7F7F2")
        DGRAY  = HexColor("#4A4A4A")
        RED    = HexColor("#C0392B")

        styles = getSampleStyleSheet()
        style_title    = ParagraphStyle("title",    fontSize=22, textColor=white,    spaceAfter=4, alignment=TA_CENTER, fontName="Helvetica-Bold")
        style_h1       = ParagraphStyle("h1",       fontSize=14, textColor=GREEN,    spaceAfter=6, spaceBefore=12, fontName="Helvetica-Bold")
        style_h2       = ParagraphStyle("h2",       fontSize=11, textColor=DGRAY,    spaceAfter=4, spaceBefore=8,  fontName="Helvetica-Bold")
        style_body     = ParagraphStyle("body",     fontSize=9,  textColor=black,    spaceAfter=3, fontName="Helvetica")
        style_small    = ParagraphStyle("small",    fontSize=8,  textColor=DGRAY,    spaceAfter=2, fontName="Helvetica")
        style_kpi_val  = ParagraphStyle("kpi_val",  fontSize=18, textColor=GREEN,    spaceAfter=2, alignment=TA_CENTER, fontName="Helvetica-Bold")
        style_kpi_lbl  = ParagraphStyle("kpi_lbl",  fontSize=8,  textColor=DGRAY,    spaceAfter=0, alignment=TA_CENTER, fontName="Helvetica")

        story = []

        # ── Cover banner ─────────────────────────────────────────────────────
        story.append(Table(
            [[Paragraph(t("app_title", lang), style_title)],
             [Paragraph(f"Executive Summary | {date.today().strftime('%d %B %Y')}", style_kpi_lbl)]],
            colWidths=[17*cm],
            style=TableStyle([
                ("BACKGROUND", (0,0), (-1,-1), GREEN),
                ("TEXTCOLOR",  (0,0), (-1,-1), white),
                ("TOPPADDING",  (0,0), (-1,-1), 12),
                ("BOTTOMPADDING",(0,0),(-1,-1), 12),
                ("ROUNDEDCORNERS", (0,0), (-1,-1), [4,4,4,4]),
            ])
        ))
        story.append(Spacer(1, 0.4*cm))
        story.append(Paragraph("CONFIDENTIAL — Ministry of Investment | وزارة الاستثمار", style_small))
        story.append(HRFlowable(width="100%", thickness=2, color=GOLD, spaceAfter=8))

        # ── KPI section ───────────────────────────────────────────────────────
        today = date.today()
        total_inv     = len(investors)
        tier1         = _count_where(investors, "Investor Tier", "Tier 1 — Strategic")
        pipeline_val  = _sum_col(investors, "Est. Investment Value (SAR)")
        commitment    = _sum_col(investors, "Actual Commitment (SAR)")
        active_opps   = _count_where(opportunities, "Opportunity Status", "Active")
        overdue       = _overdue_count(actions, today)

        kpi_data = [[
            Paragraph(f"{total_inv}", style_kpi_val),
            Paragraph(f"{tier1}", style_kpi_val),
            Paragraph(_fmt_sar(pipeline_val), style_kpi_val),
            Paragraph(_fmt_sar(commitment), style_kpi_val),
            Paragraph(f"{active_opps}", style_kpi_val),
            Paragraph(f"{overdue}", ParagraphStyle("kpi_red", fontSize=18, textColor=RED, alignment=TA_CENTER, fontName="Helvetica-Bold")),
        ],[
            Paragraph(t("total_investors", lang), style_kpi_lbl),
            Paragraph(t("tier1_investors", lang), style_kpi_lbl),
            Paragraph(t("total_pipeline_value", lang), style_kpi_lbl),
            Paragraph(t("commitment_value", lang), style_kpi_lbl),
            Paragraph(t("active_opportunities", lang), style_kpi_lbl),
            Paragraph(t("overdue_actions", lang), style_kpi_lbl),
        ]]
        kpi_table = Table(kpi_data, colWidths=[2.83*cm]*6)
        kpi_table.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), LGRAY),
            ("TOPPADDING",  (0,0), (-1,-1), 8),
            ("BOTTOMPADDING",(0,0),(-1,-1), 4),
            ("GRID",        (0,0), (-1,-1), 0.5, HexColor("#E0E0E0")),
            ("ROUNDEDCORNERS",(0,0),(-1,-1),[4,4,4,4]),
        ]))
        story.append(kpi_table)
        story.append(Spacer(1, 0.4*cm))

        # ── Investor portfolio ────────────────────────────────────────────────
        story.append(Paragraph(t("nav_investors", lang), style_h1))
        if not investors.empty:
            inv_cols = ["Company Name", "Country", "Sector", "Investor Tier",
                        "Journey Stage", "Relationship Status",
                        "Est. Investment Value (SAR)"]
            inv_cols = [c for c in inv_cols if c in investors.columns]
            inv_data = [[Paragraph(str(c), style_h2) for c in inv_cols]]
            for _, row in investors.iterrows():
                r = []
                for c in inv_cols:
                    val = row.get(c, "—")
                    if c == "Est. Investment Value (SAR)":
                        val = _fmt_sar(val) if pd.notna(val) else "—"
                    r.append(Paragraph(str(val)[:30] if pd.notna(val) else "—", style_body))
                inv_data.append(r)
            inv_table = Table(inv_data, colWidths=[17*cm/len(inv_cols)]*len(inv_cols))
            inv_table.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), GREEN),
                ("TEXTCOLOR",  (0,0), (-1,0), white),
                ("ROWBACKGROUNDS", (0,1), (-1,-1), [white, LGRAY]),
                ("GRID", (0,0), (-1,-1), 0.25, HexColor("#D0D0D0")),
                ("TOPPADDING",  (0,0), (-1,-1), 4),
                ("BOTTOMPADDING",(0,0),(-1,-1), 3),
                ("FONTSIZE", (0,1), (-1,-1), 8),
            ]))
            story.append(inv_table)
        story.append(Spacer(1, 0.3*cm))

        # ── Active opportunities ───────────────────────────────────────────────
        story.append(Paragraph(t("nav_opportunities", lang), style_h1))
        if not opportunities.empty:
            opp_active = opportunities[opportunities.get("Opportunity Status", pd.Series()) == "Active"] if "Opportunity Status" in opportunities.columns else opportunities
            if not opp_active.empty:
                opp_cols = ["Company Name", "Opportunity Name", "Sector",
                            "Opportunity Stage", "Confidence Level", "Est. Value (SAR)"]
                opp_cols = [c for c in opp_cols if c in opp_active.columns]
                opp_data = [[Paragraph(str(c), style_h2) for c in opp_cols]]
                for _, row in opp_active.head(15).iterrows():
                    r = []
                    for c in opp_cols:
                        val = row.get(c, "—")
                        if c == "Est. Value (SAR)":
                            val = _fmt_sar(val) if pd.notna(val) else "—"
                        r.append(Paragraph(str(val)[:35] if pd.notna(val) else "—", style_body))
                    opp_data.append(r)
                opp_table = Table(opp_data, colWidths=[17*cm/len(opp_cols)]*len(opp_cols))
                opp_table.setStyle(TableStyle([
                    ("BACKGROUND", (0,0), (-1,0), HexColor(MISA_GOLD)),
                    ("TEXTCOLOR",  (0,0), (-1,0), white),
                    ("ROWBACKGROUNDS", (0,1), (-1,-1), [white, LGRAY]),
                    ("GRID", (0,0), (-1,-1), 0.25, HexColor("#D0D0D0")),
                    ("TOPPADDING",  (0,0), (-1,-1), 4),
                    ("BOTTOMPADDING",(0,0),(-1,-1), 3),
                    ("FONTSIZE", (0,1), (-1,-1), 8),
                ]))
                story.append(opp_table)

        # ── Overdue actions ───────────────────────────────────────────────────
        story.append(Spacer(1, 0.3*cm))
        story.append(Paragraph(t("alerts_overdue", lang), style_h1))
        if not actions.empty and "Due Date" in actions.columns and "Status" in actions.columns:
            due = pd.to_datetime(actions["Due Date"], errors="coerce")
            overdue_df = actions[(due.dt.date < today) &
                                  (~actions["Status"].isin(["Completed","Cancelled"]))]
            if not overdue_df.empty:
                ov_cols = ["Company Name", "Action Description", "Assigned To",
                           "Due Date", "Priority", "Status"]
                ov_cols = [c for c in ov_cols if c in overdue_df.columns]
                ov_data = [[Paragraph(str(c), style_h2) for c in ov_cols]]
                for _, row in overdue_df.iterrows():
                    r = [Paragraph(str(row.get(c,"—"))[:40] if pd.notna(row.get(c)) else "—", style_body)
                         for c in ov_cols]
                    ov_data.append(r)
                ov_table = Table(ov_data, colWidths=[17*cm/len(ov_cols)]*len(ov_cols))
                ov_table.setStyle(TableStyle([
                    ("BACKGROUND", (0,0), (-1,0), RED),
                    ("TEXTCOLOR",  (0,0), (-1,0), white),
                    ("ROWBACKGROUNDS", (0,1), (-1,-1), [HexColor("#FFF5F5"), white]),
                    ("GRID", (0,0), (-1,-1), 0.25, HexColor("#D0D0D0")),
                    ("TOPPADDING",  (0,0), (-1,-1), 4),
                    ("BOTTOMPADDING",(0,0),(-1,-1), 3),
                    ("FONTSIZE", (0,1), (-1,-1), 8),
                ]))
                story.append(ov_table)
            else:
                story.append(Paragraph("✓ No overdue actions.", style_body))

        # ── Footer ────────────────────────────────────────────────────────────
        story.append(Spacer(1, 0.5*cm))
        story.append(HRFlowable(width="100%", thickness=1, color=GOLD))
        story.append(Paragraph(
            f"Generated: {date.today().strftime('%d %B %Y')} | CONFIDENTIAL | Ministry of Investment",
            ParagraphStyle("footer", fontSize=7, textColor=DGRAY, alignment=TA_CENTER)
        ))

        doc.build(story)
        return buf.getvalue()

    except Exception as exc:
        return _fallback_pdf(str(exc))


def _fallback_pdf(error_msg: str) -> bytes:
    """Return a minimal error PDF if reportlab build fails."""
    content = f"""%PDF-1.4
1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj
2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj
3 0 obj<</Type/Page/MediaBox[0 0 612 792]/Parent 2 0 R/Contents 4 0 R/Resources<</Font<</F1 5 0 R>>>>>>endobj
4 0 obj<</Length 80>>stream
BT /F1 12 Tf 72 720 Td (PDF generation error: {error_msg[:60]}) Tj ET
endstream endobj
5 0 obj<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>endobj
xref 0 6
trailer<</Size 6/Root 1 0 R>>
startxref 0
%%EOF"""
    return content.encode()


# ── Helpers ───────────────────────────────────────────────────────────────────

def _sum_col(df: pd.DataFrame, col: str) -> float:
    if df.empty or col not in df.columns:
        return 0.0
    return pd.to_numeric(df[col], errors="coerce").fillna(0).sum()


def _count_where(df: pd.DataFrame, col: str, val: str) -> int:
    if df.empty or col not in df.columns:
        return 0
    return int((df[col] == val).sum())


def _overdue_count(actions: pd.DataFrame, today: date) -> int:
    if actions.empty or "Due Date" not in actions.columns or "Status" not in actions.columns:
        return 0
    due = pd.to_datetime(actions["Due Date"], errors="coerce")
    return int(((due.dt.date < today) & (~actions["Status"].isin(["Completed","Cancelled"]))).sum())


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
