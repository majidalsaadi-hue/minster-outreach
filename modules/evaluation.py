
# Evaluation & Briefing — Ministerial Briefing Note generator
# Upload visitor bio / email / company profile → Claude extracts → formatted .docx

import io
import os
import base64
import json
import re
from datetime import date, datetime

import streamlit as st
import docx
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ─── Constants ─────────────────────────────────────────────────────────────────

_GREEN  = "1B5C3F"   # MISA green (hex, no #)
_LGREEN = "EAF4EE"   # Light green
_MGREEN = "C5DDD0"   # Mid green
_GOLD   = "C9974A"
_GREY   = "999999"
_DARK   = "1A1A1A"
_MED    = "444444"

_MODEL  = "claude-sonnet-4-6"

# ─── Session state ─────────────────────────────────────────────────────────────

def _init():
    defaults = {
        "ev_api_key":   "",
        "ev_brief":     None,
        "ev_docx":      None,
        "ev_context":   "",
        "ev_file_key":  "",
        "ev_photo_key": "",
        "ev_photo_bytes": None,
        "ev_logo_key":  "",
        "ev_logo_bytes": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ─── CSS ───────────────────────────────────────────────────────────────────────

def _css():
    st.markdown("""
    <style>
    .ev-badge {
        background:#1B5C3F; color:white; font-size:11px; font-weight:700;
        padding:3px 10px; border-radius:12px; letter-spacing:.04em;
    }
    .ev-section {
        font-size:13px; font-weight:700; color:#1B5C3F;
        text-transform:uppercase; letter-spacing:.06em; margin:0;
    }
    .ev-brief-wrap {
        border:1.5px solid #C5DDD0; border-radius:10px; overflow:hidden;
        font-family:'Segoe UI',Arial,sans-serif;
    }
    .ev-brief-hdr {
        background:#1B5C3F; color:white; padding:14px 20px;
        display:flex; align-items:center; justify-content:space-between;
    }
    .ev-brief-hdr h2 { font-size:15px; font-weight:700; margin:0; }
    .ev-brief-hdr p  { font-size:11px; color:#C8E6D4; margin:3px 0 0; }
    .ev-conf  { color:#FFD700; font-weight:700; font-size:12px; }
    .ev-brief-body { padding:18px 20px; }
    .ev-shead {
        font-size:11px; font-weight:700; letter-spacing:.08em; color:#1B5C3F;
        text-transform:uppercase; border-bottom:2px solid #1B5C3F;
        padding-bottom:4px; margin:18px 0 10px;
    }
    .ev-subj { font-size:13px; margin-bottom:14px; padding-bottom:10px;
               border-bottom:1px solid #E8F0EC; }
    .ev-ptable { width:100%; border-collapse:collapse; font-size:12px; }
    .ev-ptable td { padding:5px 8px; }
    .ev-ptable tr:nth-child(odd) td { background:#EAF4EE; }
    .ev-ptd-lbl { font-weight:700; color:#1B5C3F; width:28%; }
    .ev-body p { font-size:12px; line-height:1.6; color:#333; margin:0 0 8px; }
    .ev-cols { display:grid; grid-template-columns:1fr 1fr; gap:0 20px; }
    .ev-bul { font-size:12px; margin:4px 0; padding-left:14px; position:relative; }
    .ev-bul::before { content:'•'; position:absolute; left:0; color:#1B5C3F; font-weight:700; }
    .ev-sub { font-size:11px; margin:2px 0 2px 22px; color:#555; position:relative; padding-left:12px; }
    .ev-sub::before { content:'–'; position:absolute; left:0; }
    .ev-recbox {
        background:#EAF4EE; border:1.5px solid #1B5C3F;
        border-radius:6px; padding:14px 16px; margin:14px 0;
    }
    .ev-recbox h4 { font-size:12px; font-weight:700; color:#1B5C3F; margin:0 0 8px; }
    .ev-footer {
        border-top:1px solid #C5DDD0; padding:10px 20px;
        display:flex; justify-content:space-between; font-size:10px; color:#999;
    }
    </style>
    """, unsafe_allow_html=True)


# ─── File helpers ───────────────────────────────────────────────────────────────

def _media_type(name: str, ftype: str) -> str:
    if ftype:
        return ftype
    ext = name.rsplit(".", 1)[-1].lower()
    return {
        "pdf": "application/pdf",
        "png": "image/png",
        "jpg": "image/jpeg", "jpeg": "image/jpeg",
        "txt": "text/plain",
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }.get(ext, "application/octet-stream")


def _extract_text_from_docx(raw: bytes) -> str:
    try:
        d = Document(io.BytesIO(raw))
        parts = [p.text for p in d.paragraphs if p.text.strip()]
        for tbl in d.tables:
            for row in tbl.rows:
                row_txt = " | ".join(c.text.strip() for c in row.cells if c.text.strip())
                if row_txt:
                    parts.append(row_txt)
        return "\n".join(parts)
    except Exception:
        return raw.decode("utf-8", errors="ignore")


def _fetch_logo(domain: str) -> bytes | None:
    """Try to fetch company logo from Clearbit. Returns bytes or None."""
    if not domain:
        return None
    try:
        import urllib.request
        url = f"https://logo.clearbit.com/{domain.strip().lower()}"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=5) as resp:
            if resp.status == 200:
                return resp.read()
    except Exception:
        pass
    return None


def _build_content_blocks(files: list, context: str) -> list:
    """Build Anthropic message content blocks from uploaded files."""
    content = []
    for f in files:
        mt = f["media_type"]
        b64 = base64.b64encode(f["bytes"]).decode()
        if mt == "application/pdf":
            content.append({"type": "document",
                             "source": {"type": "base64", "media_type": mt, "data": b64}})
        elif mt.startswith("image/"):
            content.append({"type": "image",
                             "source": {"type": "base64", "media_type": mt, "data": b64}})
        elif "docx" in mt or "word" in mt:
            txt = _extract_text_from_docx(f["bytes"])
            content.append({"type": "text", "text": f"--- File: {f['name']} ---\n{txt}"})
        else:
            try:
                txt = f["bytes"].decode("utf-8", errors="ignore")
                content.append({"type": "text", "text": f"--- File: {f['name']} ---\n{txt}"})
            except Exception:
                pass

    today = date.today().strftime("%d %B %Y").lstrip("0")
    ctx_block = ("Additional context from the user:\n" + context + "\n") if context else ""
    prompt = f"""Today's date: {today}

{ctx_block}

Extract all available information from the uploaded files and produce a MISA Ministerial Briefing Note as a structured JSON object.

Respond ONLY with a JSON object — no preamble, no markdown fences.

Required JSON structure:
{{
  "visitorName": "string",
  "visitorTitle": "string",
  "company": "string",
  "companyShort": "string (2-3 word abbreviation)",
  "companyDomain": "string (primary website domain e.g. capitaland.com, blackrock.com — no https://)",
  "visitDates": "string",
  "accompaniedBy": "string",
  "organisation": "string (include country and ecosystem)",
  "revenue": "string (with currency and year)",
  "employees": "string",
  "aum": "string (if applicable, else empty string)",
  "subject": "string (one-line meeting subject for header)",
  "refNumber": "string (format MISA/BRIEF/YYYY-MM)",
  "strategicContext": "string (2-3 sentences, Vision 2030 alignment, ecosystem link)",
  "sectors": [
    {{
      "title": "string (sector name)",
      "subbullet": "string (one line elaboration)"
    }}
  ],
  "recommendation": {{
    "delegateTo": "string (name and title of recommended delegate)",
    "rationale": ["string", "string", "string"]
  }},
  "discussionPoints": ["string", "string", "string", "string"]
}}

Rules:
- sectors: extract from any images/screenshots showing sector bullets; otherwise derive from company portfolio; include 4-6 sectors
- recommendation.delegateTo: default "Assistant Minister H.E. Ibrahim" unless visitor is CEO of Fortune 100 or top sovereign fund, then suggest Minister directly
- always map sectors to Vision 2030 pillars in strategicContext
- revenue/employees/aum: if not in documents, estimate from company's known profile and mark as "approx."
- discussionPoints: exactly 4 points
- aum: empty string if not applicable
"""
    content.append({"type": "text", "text": prompt})
    return content


# ─── Claude API call ────────────────────────────────────────────────────────────

def _call_claude(files: list, context: str, api_key: str) -> dict:
    import anthropic
    client = anthropic.Anthropic(api_key=api_key)
    content = _build_content_blocks(files, context)
    response = client.messages.create(
        model=_MODEL,
        max_tokens=2000,
        messages=[{"role": "user", "content": content}],
    )
    raw = "".join(b.text for b in response.content if hasattr(b, "text"))
    # Strip accidental markdown fences
    raw = re.sub(r"^```[a-z]*\n?", "", raw, flags=re.MULTILINE).rstrip("`").strip()
    return json.loads(raw)


# ─── .docx generator ────────────────────────────────────────────────────────────

def _hex_to_rgb(h: str):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


def _cell_shading(cell, hex_color: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color.lstrip("#"))
    tcPr.append(shd)


def _cell_borders(cell, **sides):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side, style in sides.items():
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),   style.get("val", "none"))
        el.set(qn("w:sz"),    str(style.get("sz", 0)))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), style.get("color", "FFFFFF"))
        tcBorders.append(el)
    tcPr.append(tcBorders)


def _no_borders(cell):
    _cell_borders(cell,
        top={"val":"none"}, bottom={"val":"none"},
        left={"val":"none"}, right={"val":"none"})


def _green_border(cell):
    s = {"val": "single", "sz": 8, "color": _GREEN}
    _cell_borders(cell, top=s, bottom=s, left=s, right=s)


def _add_run(para, text, bold=False, italic=False, size=11,
             color=_DARK, font="Arial"):
    run = para.add_run(text)
    run.bold = italic
    run.bold = bold
    run.italic = italic
    run.font.name = font
    run.font.size = Pt(size)
    r, g, b = _hex_to_rgb(color)
    run.font.color.rgb = RGBColor(r, g, b)
    return run


def _section_head(doc, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(3)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom")
    bot.set(qn("w:val"),   "single")
    bot.set(qn("w:sz"),    "8")
    bot.set(qn("w:space"), "4")
    bot.set(qn("w:color"), _GREEN)
    pBdr.append(bot)
    pPr.append(pBdr)
    run = p.add_run(text.upper())
    run.bold      = True
    run.font.name = "Arial"
    run.font.size = Pt(16)
    r, g, b = _hex_to_rgb(_GREEN)
    run.font.color.rgb = RGBColor(r, g, b)
    return p


def _bullet_para(doc_or_cell, text: str, sub=False, bold=False, font_size=9):
    """Add a bullet paragraph to doc or a table cell."""
    target = doc_or_cell
    p = target.add_paragraph(style="List Bullet" if not sub else "List Bullet 2")
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    run = p.add_run(text)
    run.bold      = bold
    run.font.name = "Arial"
    run.font.size = Pt(font_size)
    r, g, b = _hex_to_rgb(_MED if sub else _DARK)
    run.font.color.rgb = RGBColor(r, g, b)
    return p


def _build_docx(d: dict, photo_bytes: bytes = None, logo_bytes: bytes = None) -> bytes:
    doc = Document()

    # ── Page setup (A4, tighter margins to fit one page) ─────────────────────
    section = doc.sections[0]
    section.page_height   = Cm(29.7)
    section.page_width    = Cm(21.0)
    section.top_margin    = Cm(1.0)
    section.bottom_margin = Cm(1.0)
    section.left_margin   = Cm(1.5)
    section.right_margin  = Cm(1.5)

    today = date.today().strftime("%d %B %Y").lstrip("0")

    # ── Remove default paragraph spacing ─────────────────────────────────────
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(10)

    # ── HEADER TABLE (green, 3 cols) ─────────────────────────────────────────
    hdr_tbl = doc.add_table(rows=1, cols=3)
    hdr_tbl.style = "Table Grid"
    hdr_tbl.autofit = False
    from docx.shared import Inches
    col_widths = [Cm(9.5), Cm(4.5), Cm(4.0)]
    for i, w in enumerate(col_widths):
        hdr_tbl.columns[i].width = w

    c0, c1, c2 = hdr_tbl.rows[0].cells

    for cell in (c0, c1, c2):
        _cell_shading(cell, _GREEN)
        _no_borders(cell)

    # Left cell
    c0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = c0.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(2)
    _add_run(p, "MINISTERIAL BRIEFING NOTE", bold=True, size=13, color="FFFFFF")
    p2 = c0.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    _add_run(p2, "Ministry of Investment of Saudi Arabia (MISA)", size=9, color="C8E6D4")
    c0.paragraphs[0].paragraph_format.space_before = Pt(4)

    # Middle cell — company logo or company name
    c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if logo_bytes:
        try:
            pl_logo = c1.add_paragraph()
            pl_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pl_logo.paragraph_format.space_before = Pt(4)
            pl_logo.paragraph_format.space_after  = Pt(4)
            pl_logo.add_run().add_picture(io.BytesIO(logo_bytes), width=Cm(3.5))
        except Exception:
            pm = c1.add_paragraph()
            pm.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _add_run(pm, d.get("company", ""), bold=True, size=12, color="FFFFFF")
    else:
        pm = c1.add_paragraph()
        pm.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_run(pm, d.get("company", ""), bold=True, size=12, color="FFFFFF")
        pm2 = c1.add_paragraph()
        pm2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_run(pm2, "Investment", size=9, color="C8E6D4")

    # Right cell — confidential / date / ref
    c2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    pr = c2.add_paragraph()
    pr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _add_run(pr, "CONFIDENTIAL", bold=True, size=9, color="FFD700")
    pr2 = c2.add_paragraph()
    pr2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _add_run(pr2, today, size=9, color="C8E6D4")
    pr3 = c2.add_paragraph()
    pr3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _add_run(pr3, d.get("refNumber", ""), size=8, color="C8E6D4")

    # ── Subject line ─────────────────────────────────────────────────────────
    subj = doc.add_paragraph()
    subj.paragraph_format.space_before = Pt(4)
    subj.paragraph_format.space_after  = Pt(4)
    _add_run(subj, "Subject:  ", bold=True, size=11)
    _add_run(subj, f"{d.get('subject','')}   |   Visit: {d.get('visitDates','')}", size=11)

    # divider
    div = doc.add_paragraph()
    div.paragraph_format.space_before = Pt(2)
    div.paragraph_format.space_after  = Pt(6)
    pPr = div._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom")
    bot.set(qn("w:val"),   "single"); bot.set(qn("w:sz"), "4")
    bot.set(qn("w:space"), "4");      bot.set(qn("w:color"), _MGREEN)
    pBdr.append(bot); pPr.append(pBdr)

    # ── Visitor Profile ───────────────────────────────────────────────────────
    _section_head(doc, "Visitor Profile")

    profile_rows = [
        ("Name",           d.get("visitorName",    "")),
        ("Title",          d.get("visitorTitle",   "")),
        ("Accompanied by", d.get("accompaniedBy",  "")),
        ("Organisation",   d.get("organisation",   "")),
        ("Revenue",        d.get("revenue",        "")),
        ("Employees",      d.get("employees",      "")),
    ]
    if d.get("aum"):
        profile_rows.append(("AUM", d["aum"]))

    # Profile table (label | value | name placeholder), 3 cols
    n_rows = len(profile_rows)
    pt = doc.add_table(rows=n_rows, cols=3)
    pt.style = "Table Grid"
    pt.autofit = False
    pt.columns[0].width = Cm(4.0)
    pt.columns[1].width = Cm(9.8)
    pt.columns[2].width = Cm(3.7)
    for i, (lbl, val) in enumerate(profile_rows):
        lc, vc, nc = pt.rows[i].cells
        fill = _LGREEN if i % 2 == 0 else "FFFFFF"
        _cell_shading(lc, fill); _cell_shading(vc, fill)
        _no_borders(lc); _no_borders(vc); _no_borders(nc)
        lc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        vc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        pl = lc.add_paragraph()
        pl.paragraph_format.space_before = Pt(3)
        pl.paragraph_format.space_after  = Pt(3)
        _add_run(pl, lbl, bold=True, size=10, color=_GREEN)
        pv = vc.add_paragraph()
        pv.paragraph_format.space_before = Pt(3)
        pv.paragraph_format.space_after  = Pt(3)
        _add_run(pv, val or "—", size=10, color=_DARK)
        # Right column: photo (all rows merged visually) or name/title placeholder
        _cell_shading(nc, "F0F7F3")
        if i == 0:
            nc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if photo_bytes:
                try:
                    pp_img = nc.add_paragraph()
                    pp_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    pp_img.paragraph_format.space_before = Pt(4)
                    pp_img.paragraph_format.space_after  = Pt(2)
                    pp_img.add_run().add_picture(io.BytesIO(photo_bytes), width=Cm(3.0))
                except Exception:
                    pn = nc.add_paragraph()
                    pn.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    _add_run(pn, d.get("visitorName", ""), bold=True, size=9, color=_GREEN)
            else:
                pn = nc.add_paragraph()
                pn.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pn.paragraph_format.space_before = Pt(4)
                _add_run(pn, d.get("visitorName", ""), bold=True, size=9, color=_GREEN)
        elif i == 1 and not photo_bytes:
            nc.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            pt2 = nc.add_paragraph()
            pt2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title_short = (d.get("visitorTitle") or "").split(",")[0]
            _add_run(pt2, title_short, size=8, color="666666")

    # ── Strategic Context ─────────────────────────────────────────────────────
    _section_head(doc, "Strategic Context")
    ctx = doc.add_paragraph()
    ctx.paragraph_format.space_before = Pt(2)
    ctx.paragraph_format.space_after  = Pt(3)
    _add_run(ctx, d.get("strategicContext", ""), size=11)

    # ── Areas / Sectors (2-column) ────────────────────────────────────────────
    company_short = d.get("companyShort", d.get("company", ""))
    _section_head(doc, f"Areas Where {company_short} Can Benefit the Kingdom")

    sectors = d.get("sectors", [])
    mid = (len(sectors) + 1) // 2
    left_s, right_s = sectors[:mid], sectors[mid:]

    st_tbl = doc.add_table(rows=1, cols=2)
    st_tbl.style = "Table Grid"
    st_tbl.autofit = False
    st_tbl.columns[0].width = Cm(9.0)
    st_tbl.columns[1].width = Cm(9.0)
    lc_s, rc_s = st_tbl.rows[0].cells
    _no_borders(lc_s); _no_borders(rc_s)
    lc_s.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    rc_s.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    for cell, arr in ((lc_s, left_s), (rc_s, right_s)):
        for sec in arr:
            pb = cell.add_paragraph()
            pb.paragraph_format.space_before = Pt(3)
            pb.paragraph_format.space_after  = Pt(1)
            pb.paragraph_format.left_indent  = Cm(0.4)
            _add_run(pb, f"• {sec.get('title','')}", bold=True, size=11)
            ps = cell.add_paragraph()
            ps.paragraph_format.space_before = Pt(0)
            ps.paragraph_format.space_after  = Pt(4)
            ps.paragraph_format.left_indent  = Cm(0.8)
            _add_run(ps, f"– {sec.get('subbullet','')}", size=11, color=_MED)

    # ── Recommendation box ────────────────────────────────────────────────────
    rec_tbl = doc.add_table(rows=1, cols=1)
    rec_tbl.style = "Table Grid"
    rec_tbl.autofit = False
    rec_tbl.columns[0].width = Cm(18.0)
    rc = rec_tbl.rows[0].cells[0]
    _cell_shading(rc, _LGREEN)
    _green_border(rc)
    rc.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    rh = rc.add_paragraph()
    rh.paragraph_format.space_before = Pt(2)
    rh.paragraph_format.space_after  = Pt(2)
    _add_run(rh, "RECOMMENDATION", bold=True, size=13, color=_GREEN)

    rec = d.get("recommendation", {})
    rd = rc.add_paragraph()
    rd.paragraph_format.space_after = Pt(2)
    _add_run(rd, "Delegate this meeting to ", size=11)
    _add_run(rd, rec.get("delegateTo", ""), bold=True, size=11)
    _add_run(rd, ", given:", size=11)

    for r_item in rec.get("rationale", []):
        rp = rc.add_paragraph()
        rp.paragraph_format.space_before = Pt(2)
        rp.paragraph_format.space_after  = Pt(2)
        rp.paragraph_format.left_indent  = Cm(0.4)
        _add_run(rp, f"• {r_item}", size=11)

    # ── Discussion Points (2-column) ──────────────────────────────────────────
    _section_head(doc, "Suggested Discussion Points")

    dps = d.get("discussionPoints", [])
    mid2 = (len(dps) + 1) // 2
    left_d, right_d = dps[:mid2], dps[mid2:]

    dp_tbl = doc.add_table(rows=1, cols=2)
    dp_tbl.style = "Table Grid"
    dp_tbl.autofit = False
    dp_tbl.columns[0].width = Cm(9.0)
    dp_tbl.columns[1].width = Cm(9.0)
    lc_d, rc_d = dp_tbl.rows[0].cells
    _no_borders(lc_d); _no_borders(rc_d)
    lc_d.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    rc_d.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    for cell, arr in ((lc_d, left_d), (rc_d, right_d)):
        for dp in arr:
            pp = cell.add_paragraph()
            pp.paragraph_format.space_before = Pt(3)
            pp.paragraph_format.space_after  = Pt(3)
            pp.paragraph_format.left_indent  = Cm(0.4)
            _add_run(pp, f"• {dp}", size=11)

    # ── Footer ────────────────────────────────────────────────────────────────
    div2 = doc.add_paragraph()
    div2.paragraph_format.space_before = Pt(4)
    div2.paragraph_format.space_after  = Pt(2)
    pPr2 = div2._p.get_or_add_pPr()
    pBdr2 = OxmlElement("w:pBdr")
    top2  = OxmlElement("w:top")
    top2.set(qn("w:val"),   "single"); top2.set(qn("w:sz"), "4")
    top2.set(qn("w:space"), "4");      top2.set(qn("w:color"), _MGREEN)
    pBdr2.append(top2); pPr2.append(pBdr2)

    ft = doc.add_table(rows=1, cols=2)
    ft.style = "Table Grid"
    ft.autofit = False
    ft.columns[0].width = Cm(9.0)
    ft.columns[1].width = Cm(9.0)
    flc, frc = ft.rows[0].cells
    _no_borders(flc); _no_borders(frc)
    pfl = flc.add_paragraph()
    _add_run(pfl, "Prepared by: Minister Outreach Office, MISA", size=8, color=_GREY)
    pfr = frc.add_paragraph()
    pfr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _add_run(pfr, "For internal use only – Ministry of Investment of Saudi Arabia", size=8, color=_GREY)

    # ── Save to bytes ─────────────────────────────────────────────────────────
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ─── HTML preview ───────────────────────────────────────────────────────────────

def _render_preview(d: dict):
    today = date.today().strftime("%d %B %Y").lstrip("0")

    profile_rows = [
        ("Name",           d.get("visitorName",    "—")),
        ("Title",          d.get("visitorTitle",   "—")),
        ("Accompanied by", d.get("accompaniedBy",  "—")),
        ("Organisation",   d.get("organisation",   "—")),
        ("Revenue",        d.get("revenue",        "—")),
        ("Employees",      d.get("employees",      "—")),
    ]
    if d.get("aum"):
        profile_rows.append(("AUM", d["aum"]))

    profile_html = "".join(
        f'<tr><td class="ev-ptd-lbl">{l}</td><td>{v}</td></tr>'
        for l, v in profile_rows
    )

    sectors = d.get("sectors", [])
    mid = (len(sectors) + 1) // 2
    left_s, right_s = sectors[:mid], sectors[mid:]

    def sec_html(arr):
        return "".join(
            f'<div class="ev-bul"><strong>{s["title"]}</strong></div>'
            f'<div class="ev-sub">{s["subbullet"]}</div>'
            for s in arr
        )

    rec = d.get("recommendation", {})
    rat_html = "".join(f'<div class="ev-bul">{r}</div>' for r in rec.get("rationale", []))

    dps = d.get("discussionPoints", [])
    mid2 = (len(dps) + 1) // 2
    left_d  = "".join(f'<div class="ev-bul">{p}</div>' for p in dps[:mid2])
    right_d = "".join(f'<div class="ev-bul">{p}</div>' for p in dps[mid2:])

    html = f"""
    <div class="ev-brief-wrap">
      <div class="ev-brief-hdr">
        <div>
          <h2>MINISTERIAL BRIEFING NOTE</h2>
          <p>Ministry of Investment of Saudi Arabia (MISA)</p>
        </div>
        <div style="text-align:right">
          <div class="ev-conf">CONFIDENTIAL</div>
          <div style="color:#C8E6D4;font-size:11px;margin-top:4px">{today} &nbsp; {d.get("refNumber","")}</div>
        </div>
      </div>
      <div class="ev-brief-body">
        <div class="ev-subj">
          <strong>Subject:</strong> {d.get("subject","")}&nbsp;&nbsp;|&nbsp;&nbsp;
          <strong>Visit:</strong> {d.get("visitDates","")}
        </div>

        <div class="ev-shead">Visitor Profile</div>
        <table class="ev-ptable">{profile_html}</table>

        <div class="ev-shead">Strategic Context</div>
        <div class="ev-body"><p>{d.get("strategicContext","")}</p></div>

        <div class="ev-shead">Areas Where {d.get("companyShort", d.get("company",""))} Can Benefit the Kingdom</div>
        <div class="ev-cols">
          <div>{sec_html(left_s)}</div>
          <div>{sec_html(right_s)}</div>
        </div>

        <div class="ev-recbox">
          <h4>RECOMMENDATION</h4>
          <p style="font-size:12px;margin:0 0 8px">Delegate this meeting to
            <strong>{rec.get("delegateTo","")}</strong>, given:</p>
          {rat_html}
        </div>

        <div class="ev-shead">Suggested Discussion Points</div>
        <div class="ev-cols">
          <div>{left_d}</div>
          <div>{right_d}</div>
        </div>
      </div>
      <div class="ev-footer">
        <span>Prepared by: Minister Outreach Office, MISA</span>
        <span>For internal use only – Ministry of Investment of Saudi Arabia</span>
      </div>
    </div>
    """
    full = f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
*{{box-sizing:border-box;margin:0;padding:0;font-family:'Segoe UI',Arial,sans-serif}}
body{{background:#fff;padding:0}}
.wrap{{border:1.5px solid #C5DDD0;border-radius:10px;overflow:hidden}}
.hdr{{background:#1B5C3F;color:white;padding:14px 20px;display:flex;align-items:center;justify-content:space-between}}
.hdr h2{{font-size:15px;font-weight:700;margin:0}}
.hdr p{{font-size:11px;color:#C8E6D4;margin:3px 0 0}}
.conf{{color:#FFD700;font-weight:700;font-size:12px}}
.body{{padding:18px 20px}}
.shead{{font-size:11px;font-weight:700;letter-spacing:.08em;color:#1B5C3F;text-transform:uppercase;border-bottom:2px solid #1B5C3F;padding-bottom:4px;margin:18px 0 10px}}
.subj{{font-size:13px;margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid #E8F0EC}}
.ptable{{width:100%;border-collapse:collapse;font-size:12px}}
.ptable td{{padding:5px 8px}}
.ptable tr:nth-child(odd) td{{background:#EAF4EE}}
.lbl{{font-weight:700;color:#1B5C3F;width:28%}}
.bodyp{{font-size:12px;line-height:1.6;color:#333;margin:0 0 8px}}
.cols{{display:grid;grid-template-columns:1fr 1fr;gap:0 20px}}
.bul{{font-size:12px;margin:4px 0;padding-left:14px;position:relative;color:#1a1a1a}}
.bul::before{{content:'•';position:absolute;left:0;color:#1B5C3F;font-weight:700}}
.sub{{font-size:11px;margin:2px 0 2px 22px;color:#555;position:relative;padding-left:12px}}
.sub::before{{content:'–';position:absolute;left:0}}
.recbox{{background:#EAF4EE;border:1.5px solid #1B5C3F;border-radius:6px;padding:14px 16px;margin:14px 0}}
.recbox h4{{font-size:12px;font-weight:700;color:#1B5C3F;margin:0 0 8px}}
.recbox p{{font-size:12px;margin:0 0 8px;color:#1a1a1a}}
.footer{{border-top:1px solid #C5DDD0;padding:10px 20px;display:flex;justify-content:space-between;font-size:10px;color:#999}}
</style></head><body>
<div class="wrap">
  <div class="hdr">
    <div><h2>MINISTERIAL BRIEFING NOTE</h2><p>Ministry of Investment of Saudi Arabia (MISA)</p></div>
    <div style="text-align:right">
      <div class="conf">CONFIDENTIAL</div>
      <div style="color:#C8E6D4;font-size:11px;margin-top:4px">{today} &nbsp; {d.get("refNumber","")}</div>
    </div>
  </div>
  <div class="body">
    <div class="subj"><strong>Subject:</strong> {d.get("subject","")}&nbsp;&nbsp;|&nbsp;&nbsp;<strong>Visit:</strong> {d.get("visitDates","")}</div>
    <div class="shead">Visitor Profile</div>
    <table class="ptable">{profile_html}</table>
    <div class="shead">Strategic Context</div>
    <p class="bodyp">{d.get("strategicContext","")}</p>
    <div class="shead">Areas Where {d.get("companyShort", d.get("company",""))} Can Benefit the Kingdom</div>
    <div class="cols"><div>{sec_html(left_s)}</div><div>{sec_html(right_s)}</div></div>
    <div class="recbox">
      <h4>RECOMMENDATION</h4>
      <p>Delegate this meeting to <strong>{rec.get("delegateTo","")}</strong>, given:</p>
      {rat_html}
    </div>
    <div class="shead">Suggested Discussion Points</div>
    <div class="cols"><div>{left_d}</div><div>{right_d}</div></div>
  </div>
  <div class="footer">
    <span>Prepared by: Minister Outreach Office, MISA</span>
    <span>For internal use only – Ministry of Investment of Saudi Arabia</span>
  </div>
</div>
</body></html>"""
    import streamlit.components.v1 as components
    components.html(full, height=920, scrolling=True)


# ─── Main render ───────────────────────────────────────────────────────────────

def render():
    _init()
    _css()

    hc, bc = st.columns([5, 1])
    hc.markdown("**Evaluation & Briefing — Ministerial Briefing Note**")
    bc.markdown('<div style="text-align:right"><span class="ev-badge">AI-powered</span></div>',
                unsafe_allow_html=True)
    st.markdown('<hr style="margin:.25rem 0 1.25rem;border:none;border-top:0.5px solid #e5e7eb">',
                unsafe_allow_html=True)

    s = st.session_state

    # ── Step 1: Upload ────────────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown('<p class="ev-section">Step 1 — Upload visitor / company files</p>',
                    unsafe_allow_html=True)
        st.caption("Visitor bio, company profile, email screenshot — PDF, image, Word doc, or text")

        uploaded = st.file_uploader(
            "files",
            type=["pdf", "png", "jpg", "jpeg", "txt", "docx"],
            accept_multiple_files=True,
            key="ev_file_up",
            label_visibility="collapsed",
        )

        if uploaded:
            file_key = "_".join(f"{f.name}_{f.size}" for f in uploaded)
            if file_key != s.get("ev_file_key", ""):
                s["ev_file_key"] = file_key
                s["ev_loaded_files"] = [
                    {
                        "name":       f.name,
                        "bytes":      f.read(),
                        "media_type": _media_type(f.name, f.type),
                    }
                    for f in uploaded
                ]
            cols = st.columns(min(len(uploaded), 4))
            for i, f in enumerate(uploaded):
                cols[i % 4].success(f"📄 {f.name}")

    # ── Step 2: Context + images + API key ───────────────────────────────────
    with st.container(border=True):
        st.markdown('<p class="ev-section">Step 2 — Context & settings</p>',
                    unsafe_allow_html=True)
        col_a, col_b, col_c, col_d = st.columns([3, 1, 1, 2])
        with col_a:
            s["ev_context"] = st.text_area(
                "Additional context (optional)",
                value=s["ev_context"],
                height=100,
                key="ev_ctx",
                placeholder="e.g. Visitor arriving 20–22 June. Focus on logistics and data centres. Recommend delegating to HE Ibrahim…",
            )
        with col_b:
            st.markdown("**Visitor photo**")
            st.caption("Optional — PNG/JPG")
            photo_up = st.file_uploader("photo", type=["png","jpg","jpeg"],
                                        key="ev_photo_up", label_visibility="collapsed")
            if photo_up is not None:
                _pk = f"{photo_up.name}_{photo_up.size}"
                if _pk != s.get("ev_photo_key", ""):
                    s["ev_photo_key"]   = _pk
                    s["ev_photo_bytes"] = photo_up.read()
                st.image(s["ev_photo_bytes"], width=80)
        with col_c:
            st.markdown("**Company logo**")
            st.caption("Optional — PNG/JPG")
            logo_up = st.file_uploader("logo", type=["png","jpg","jpeg"],
                                       key="ev_logo_up", label_visibility="collapsed")
            if logo_up is not None:
                _lk = f"{logo_up.name}_{logo_up.size}"
                if _lk != s.get("ev_logo_key", ""):
                    s["ev_logo_key"]   = _lk
                    s["ev_logo_bytes"] = logo_up.read()
                st.image(s["ev_logo_bytes"], width=80)
        with col_d:
            # Resolve key: session state → Report Builder key → env var
            _resolved_key = (
                s.get("ev_api_key")
                or s.get("rb_ar_api_key")
                or os.environ.get("ANTHROPIC_API_KEY", "")
            )
            s["ev_api_key"] = st.text_input(
                "Anthropic API Key",
                value=_resolved_key,
                type="password",
                key="ev_key",
                placeholder="sk-ant-…  (or set ANTHROPIC_API_KEY env var)",
                help="Get your key at console.anthropic.com",
            )
            if s["ev_api_key"]:
                st.caption("✅ API key ready")
            else:
                st.caption("⚠️ Enter your Anthropic API key to generate briefings")

    # ── Step 3: Generate ─────────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown('<p class="ev-section">Step 3 — Generate briefing note</p>',
                    unsafe_allow_html=True)

        files = s.get("ev_loaded_files", [])
        api_key = (s.get("ev_api_key") or "").strip() or os.environ.get("ANTHROPIC_API_KEY", "")
        can_run = bool(files) and bool(api_key)

        g1, g2 = st.columns([3, 1])
        with g1:
            if not files:
                st.info("Upload at least one file in Step 1 to enable generation.")
            elif not api_key:
                st.info("Enter your Anthropic API key in Step 2 to enable generation.")

        run_clicked = st.button(
            "▶  Generate Ministerial Briefing Note",
            type="primary",
            use_container_width=True,
            key="ev_run",
            disabled=not can_run,
        )

        if run_clicked and can_run:
            prog = st.progress(0)
            status = st.empty()
            try:
                status.markdown("→ Reading uploaded files…")
                prog.progress(15)

                status.markdown("→ Extracting visitor details & company information…")
                prog.progress(35)

                brief = _call_claude(files, s.get("ev_context", ""), api_key)
                prog.progress(60)

                # Auto-fetch logo from Clearbit if not manually uploaded
                logo_bytes = s.get("ev_logo_bytes")
                if not logo_bytes:
                    domain = brief.get("companyDomain", "")
                    if domain:
                        status.markdown(f"→ Fetching {brief.get('company','')} logo…")
                        logo_bytes = _fetch_logo(domain)
                        if logo_bytes:
                            s["ev_logo_bytes"] = logo_bytes

                prog.progress(75)

                status.markdown("→ Generating briefing document…")
                docx_bytes = _build_docx(
                    brief,
                    photo_bytes=s.get("ev_photo_bytes"),
                    logo_bytes=logo_bytes,
                )
                prog.progress(100)

                s["ev_brief"] = brief
                s["ev_docx"]  = docx_bytes

                status.success(f"✅ Briefing note ready — {brief.get('visitorName','')} / {brief.get('company','')}")
                prog.empty()

            except Exception as e:
                prog.empty()
                status.error(f"Error: {e}")
                if "api_key" in str(e).lower() or "401" in str(e):
                    st.error("Invalid API key — check your Anthropic key at console.anthropic.com")

    # ── Results ───────────────────────────────────────────────────────────────
    brief = s.get("ev_brief")
    docx_bytes = s.get("ev_docx")

    if brief and docx_bytes:
        with st.container(border=True):
            st.markdown('<p class="ev-section">📄 Ministerial Briefing Note</p>',
                        unsafe_allow_html=True)

            company_short = (brief.get("companyShort") or brief.get("company", "Brief")).replace(" ", "_")
            month = date.today().strftime("%Y-%m")
            fname = f"MISA_Briefing_{company_short}_{month}.docx"

            dl1, dl2 = st.columns([2, 1])
            with dl1:
                st.download_button(
                    "⬇  Download Briefing Note (.docx)",
                    data=docx_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="ev_dl",
                )
            with dl2:
                if st.button("+ New Brief", use_container_width=True, key="ev_new"):
                    for k in ["ev_brief", "ev_docx", "ev_context", "ev_file_key",
                              "ev_loaded_files", "ev_api_key"]:
                        s.pop(k, None)
                    st.rerun()

            st.markdown("---")
            _render_preview(brief)
