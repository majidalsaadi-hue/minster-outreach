
# Evaluation & Briefing — Ministerial Briefing Note generator
# Upload visitor bio / email / company profile → Claude extracts → formatted .docx

import io
import os
import base64
import json
import re
import zipfile
import urllib.parse
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
        "ev_api_key":    "",
        "ev_brief":      None,
        "ev_docx":       None,
        "ev_context":    "",
        "ev_file_key":   "",
        "ev_photo_key":  "",
        "ev_photo_bytes": None,
        "ev_logo_key":   "",
        "ev_logo_bytes": None,
        "ev_attendees":  "",
        "ev_news":       [],
        "ev_auto_photo": False,
        "ev_website":    "",
        "ev_email":      "",
        "ev_phone":      "",
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


def _extract_photo_from_file(file_bytes: bytes, media_type: str) -> bytes | None:
    """Try to pull a person photo out of a bio PDF or DOCX without extra libraries."""
    try:
        if "pdf" in media_type:
            # Scan raw bytes for an embedded JPEG (most bio PDFs embed one)
            start = file_bytes.find(b'\xff\xd8\xff')
            end   = file_bytes.rfind(b'\xff\xd9')
            if start != -1 and end > start and (end - start) > 8000:
                return file_bytes[start:end + 2]
        elif "docx" in media_type or "word" in media_type:
            with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
                imgs = sorted(
                    [n for n in z.namelist()
                     if n.startswith("word/media/") and
                     n.lower().rsplit(".", 1)[-1] in ("jpg", "jpeg", "png")],
                    key=lambda n: z.getinfo(n).file_size, reverse=True,
                )
                if imgs:
                    return z.read(imgs[0])
    except Exception:
        pass
    return None


def _fetch_news_for_company(company_name: str, domain: str = "") -> list:
    """Fetch latest company news via Google News RSS. Returns list of {title, date}."""
    import urllib.request
    import xml.etree.ElementTree as ET
    query = urllib.parse.quote(f'"{company_name}"' if company_name else domain.split(".")[0])
    url = f"https://news.google.com/rss/search?q={query}&hl=en&gl=US&ceid=US:en"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            tree = ET.parse(resp)
        items = []
        for item in tree.getroot().iter("item"):
            title   = (item.findtext("title") or "").split(" - ")[0].strip()
            pubdate = (item.findtext("pubDate") or "")[:16].strip()
            source  = ""
            src_el  = item.find("{https://news.google.com/rss}source")
            if src_el is not None:
                source = src_el.text or ""
            if title and len(title) > 10:
                items.append({"title": title, "date": pubdate, "source": source})
            if len(items) >= 4:
                break
        return items
    except Exception:
        return []


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
  "discussionPoints": ["string", "string", "string", "string"],
  "investmentRegions": [
    {{
      "region": "string (e.g. Southeast Asia, Europe, GCC)",
      "focus": "string (key sectors or themes in that region, one line)"
    }}
  ],
  "globalSubsidiaries": ["string (subsidiary name — country, e.g. Acme Capital — UK)"],
  "saudiPresence": {{
    "investments": "string (current/planned investments in Saudi Arabia; 'No known current investments' if none)",
    "jvPartners": ["string (Saudi local JV partner or government entity name)"],
    "majorProjects": ["string (project name — brief description)"]
  }}
}}

Rules:
- sectors: extract from any images/screenshots showing sector bullets; otherwise derive from company portfolio; include 4-6 sectors
- recommendation.delegateTo: default "Assistant Minister H.E. Ibrahim Al-Rashed" unless visitor is CEO of Fortune 100 or top sovereign fund, then suggest "H.E. Minister Fahad Al-Saif" directly
- always map sectors to Vision 2030 pillars in strategicContext
- revenue/employees/aum: if not in documents, estimate from company's known profile and mark as "est."
- discussionPoints: exactly 4 points
- aum: empty string if not applicable
- investmentRegions: 3-6 regions where the company actively deploys capital or operates; omit if company is purely domestic
- globalSubsidiaries: up to 8 major subsidiaries, JV vehicles, or related entities with their country; empty array if not applicable
- saudiPresence.investments: describe any Saudi investment, commitment, or MOU; if none write "No known current investments in Saudi Arabia"
- saudiPresence.jvPartners: Saudi counterparties (ARAMCO, PIF entities, local developers, etc.); empty array if none
- saudiPresence.majorProjects: named projects, NEOM involvement, data centres, manufacturing plants, etc.; empty array if none
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
        max_tokens=4000,
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
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(2)
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


def _build_docx(d: dict, photo_bytes: bytes = None, logo_bytes: bytes = None,
                attendees: str = "", news: list = None,
                contact_email: str = "", contact_phone: str = "",
                meeting_mode: str = "recommendation", direction_host: str = "") -> bytes:
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
    if contact_email:
        profile_rows.append(("Email", contact_email))
    if contact_phone:
        profile_rows.append(("Phone", contact_phone))

    # Profile table (label | value | name placeholder), 3 cols
    n_rows = len(profile_rows)
    pt = doc.add_table(rows=n_rows, cols=3)
    pt.style = "Table Grid"
    pt.autofit = False
    pt.columns[0].width = Cm(3.2)
    pt.columns[1].width = Cm(10.5)
    pt.columns[2].width = Cm(4.3)
    for i, (lbl, val) in enumerate(profile_rows):
        lc, vc, nc = pt.rows[i].cells
        fill = _LGREEN if i % 2 == 0 else "FFFFFF"
        _cell_shading(lc, fill); _cell_shading(vc, fill)
        _no_borders(lc); _no_borders(vc); _no_borders(nc)
        lc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        vc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        pl = lc.add_paragraph()
        pl.paragraph_format.space_before = Pt(2)
        pl.paragraph_format.space_after  = Pt(2)
        _add_run(pl, lbl, bold=True, size=10, color=_GREEN)
        pv = vc.add_paragraph()
        pv.paragraph_format.space_before = Pt(2)
        pv.paragraph_format.space_after  = Pt(2)
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
                    pp_img.add_run().add_picture(io.BytesIO(photo_bytes), width=Cm(3.8))
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
            ps.paragraph_format.space_after  = Pt(2)
            ps.paragraph_format.left_indent  = Cm(0.8)
            _add_run(ps, f"– {sec.get('subbullet','')}", size=11, color=_MED)

    # ── Decision box (Direction or Recommendation) ────────────────────────────
    rec_tbl = doc.add_table(rows=1, cols=1)
    rec_tbl.style = "Table Grid"
    rec_tbl.autofit = False
    rec_tbl.columns[0].width = Cm(18.0)
    rc = rec_tbl.rows[0].cells[0]
    rc.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    if meeting_mode == "direction":
        _cell_shading(rc, "EAF4EE")
        _green_border(rc)
        rh = rc.add_paragraph()
        rh.paragraph_format.space_before = Pt(2)
        rh.paragraph_format.space_after  = Pt(2)
        _add_run(rh, "DIRECTION", bold=True, size=13, color=_GREEN)
        host = direction_host or "H.E. Fahad Al-Saif, Minister of Investment"
        rd = rc.add_paragraph()
        rd.paragraph_format.space_after = Pt(4)
        _add_run(rd, "This meeting has been approved. ", size=11)
        _add_run(rd, host, bold=True, size=11)
        _add_run(rd, " will host this meeting directly.", size=11)
    else:
        _cell_shading(rc, _LGREEN)
        _green_border(rc)
        rh = rc.add_paragraph()
        rh.paragraph_format.space_before = Pt(2)
        rh.paragraph_format.space_after  = Pt(2)
        _add_run(rh, "RECOMMENDATION", bold=True, size=13, color=_GREEN)
        rec = d.get("recommendation", {})
        _sel_delegate = _get_selected_delegate_name()
        _delegate_display = _sel_delegate if _sel_delegate else rec.get("delegateTo", "")
        rd = rc.add_paragraph()
        rd.paragraph_format.space_after = Pt(2)
        if _delegate_display:
            _add_run(rd, "Delegate this meeting to ", size=11)
            _add_run(rd, _delegate_display, bold=True, size=11)
            _add_run(rd, ", given:", size=11)
        else:
            _add_run(rd, "Select meeting nature in Decision & Direction to confirm the recommended delegate.", size=11, color=_MED)
        for r_item in rec.get("rationale", []):
            rp = rc.add_paragraph()
            rp.paragraph_format.space_before = Pt(2)
            rp.paragraph_format.space_after  = Pt(2)
            rp.paragraph_format.left_indent  = Cm(0.4)
            _add_run(rp, f"• {r_item}", size=11)

    # ── Ministry Recommended Attendees ────────────────────────────────────────
    if attendees and attendees.strip():
        att_tbl = doc.add_table(rows=1, cols=1)
        att_tbl.style = "Table Grid"
        att_tbl.autofit = False
        att_tbl.columns[0].width = Cm(18.0)
        ac = att_tbl.rows[0].cells[0]
        _cell_shading(ac, "FFF8EC")
        s_g = {"val": "single", "sz": 8, "color": _GOLD}
        _cell_borders(ac, top=s_g, bottom=s_g, left=s_g, right=s_g)
        ac.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        ah = ac.add_paragraph()
        ah.paragraph_format.space_before = Pt(2)
        ah.paragraph_format.space_after  = Pt(4)
        _add_run(ah, "RECOMMENDED MINISTRY ATTENDEES", bold=True, size=12, color=_GOLD)
        for name_line in attendees.strip().splitlines():
            name_line = name_line.strip().lstrip("•-").strip()
            if name_line:
                ap = ac.add_paragraph()
                ap.paragraph_format.space_before = Pt(2)
                ap.paragraph_format.space_after  = Pt(2)
                ap.paragraph_format.left_indent  = Cm(0.4)
                _add_run(ap, f"• {name_line}", size=11)

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
            pp.paragraph_format.space_before = Pt(2)
            pp.paragraph_format.space_after  = Pt(2)
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

    # ══════════════════════════════════════════════════════════════════════════
    # PAGE 2 — Investment Intelligence Brief
    # ══════════════════════════════════════════════════════════════════════════
    from docx.enum.text import WD_BREAK

    pb = doc.add_paragraph()
    pb.paragraph_format.space_before = Pt(0)
    pb.paragraph_format.space_after  = Pt(0)
    pb.add_run().add_break(WD_BREAK.PAGE)

    # ── Page 2 header ─────────────────────────────────────────────────────────
    p2_hdr = doc.add_table(rows=1, cols=2)
    p2_hdr.style = "Table Grid"
    p2_hdr.autofit = False
    p2_hdr.columns[0].width = Cm(10.7)
    p2_hdr.columns[1].width = Cm(7.3)
    p2h_l, p2h_r = p2_hdr.rows[0].cells
    _cell_shading(p2h_l, _GREEN); _cell_shading(p2h_r, _GREEN)
    _no_borders(p2h_l); _no_borders(p2h_r)
    p2h_l.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    ph2_title = p2h_l.add_paragraph()
    ph2_title.paragraph_format.space_before = Pt(4)
    ph2_title.paragraph_format.space_after  = Pt(2)
    _add_run(ph2_title, f"{d.get('company','')} — Investment Intelligence Brief",
             bold=True, size=12, color="FFFFFF")
    ph2_sub = p2h_l.add_paragraph()
    ph2_sub.paragraph_format.space_before = Pt(0)
    ph2_sub.paragraph_format.space_after  = Pt(4)
    _add_run(ph2_sub, "Global Presence · Saudi Contribution · Market Intelligence",
             size=9, color="C8E6D4")
    p2h_r.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    ph2_r = p2h_r.add_paragraph()
    ph2_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    ph2_r.paragraph_format.space_before = Pt(4)
    _add_run(ph2_r, "CONFIDENTIAL", bold=True, size=9, color="FFD700")
    ph2_r2 = p2h_r.add_paragraph()
    ph2_r2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    ph2_r2.paragraph_format.space_after = Pt(4)
    _add_run(ph2_r2, date.today().strftime("%d %B %Y").lstrip("0"), size=8, color="C8E6D4")

    # ── Global Investment Regions + Subsidiaries ──────────────────────────────
    _section_head(doc, "Global Investment Footprint")

    inv_regions = d.get("investmentRegions", [])
    subsidiaries = d.get("globalSubsidiaries", [])

    gi_tbl = doc.add_table(rows=1, cols=2)
    gi_tbl.style = "Table Grid"
    gi_tbl.autofit = False
    gi_tbl.columns[0].width = Cm(9.0)
    gi_tbl.columns[1].width = Cm(9.0)
    gi_l, gi_r = gi_tbl.rows[0].cells
    _no_borders(gi_l); _no_borders(gi_r)
    gi_l.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    gi_r.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    # Left col: Regions
    gi_lh = gi_l.add_paragraph()
    gi_lh.paragraph_format.space_before = Pt(4)
    gi_lh.paragraph_format.space_after  = Pt(4)
    _add_run(gi_lh, "Where They Invest", bold=True, size=10, color=_GREEN)
    if inv_regions:
        for reg in inv_regions:
            rp2 = gi_l.add_paragraph()
            rp2.paragraph_format.space_before = Pt(2)
            rp2.paragraph_format.space_after  = Pt(1)
            rp2.paragraph_format.left_indent  = Cm(0.3)
            _add_run(rp2, f"• {reg.get('region','')}", bold=True, size=10)
            rs = gi_l.add_paragraph()
            rs.paragraph_format.space_before = Pt(0)
            rs.paragraph_format.space_after  = Pt(3)
            rs.paragraph_format.left_indent  = Cm(0.7)
            _add_run(rs, f"– {reg.get('focus','')}", size=9, color=_MED)
    else:
        _add_run(gi_l.add_paragraph(), "Data not available", size=9, color=_GREY)

    # Right col: Subsidiaries
    gi_rh = gi_r.add_paragraph()
    gi_rh.paragraph_format.space_before = Pt(4)
    gi_rh.paragraph_format.space_after  = Pt(4)
    _add_run(gi_rh, "Global Entities & Subsidiaries", bold=True, size=10, color=_GREEN)
    if subsidiaries:
        for sub in subsidiaries:
            sp = gi_r.add_paragraph()
            sp.paragraph_format.space_before = Pt(2)
            sp.paragraph_format.space_after  = Pt(2)
            sp.paragraph_format.left_indent  = Cm(0.3)
            _add_run(sp, f"• {sub}", size=10)
    else:
        _add_run(gi_r.add_paragraph(), "Data not available", size=9, color=_GREY)

    # ── Saudi Arabia Presence ─────────────────────────────────────────────────
    _section_head(doc, "Saudi Arabia Presence & Contribution")

    saudi = d.get("saudiPresence", {})
    sa_tbl = doc.add_table(rows=1, cols=1)
    sa_tbl.style = "Table Grid"
    sa_tbl.autofit = False
    sa_tbl.columns[0].width = Cm(18.0)
    sa_c = sa_tbl.rows[0].cells[0]
    _cell_shading(sa_c, _LGREEN)
    _green_border(sa_c)
    sa_c.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    # Investments row
    sa_inv_h = sa_c.add_paragraph()
    sa_inv_h.paragraph_format.space_before = Pt(4)
    sa_inv_h.paragraph_format.space_after  = Pt(2)
    _add_run(sa_inv_h, "Investments & Commitments in Saudi Arabia", bold=True, size=11, color=_GREEN)
    sa_inv_p = sa_c.add_paragraph()
    sa_inv_p.paragraph_format.space_before = Pt(0)
    sa_inv_p.paragraph_format.space_after  = Pt(6)
    sa_inv_p.paragraph_format.left_indent  = Cm(0.3)
    _add_run(sa_inv_p, saudi.get("investments", "No known current investments in Saudi Arabia"),
             size=10)

    # JV Partners
    jv_partners = saudi.get("jvPartners", [])
    if jv_partners:
        sa_jvh = sa_c.add_paragraph()
        sa_jvh.paragraph_format.space_before = Pt(2)
        sa_jvh.paragraph_format.space_after  = Pt(2)
        _add_run(sa_jvh, "Joint Venture Partners (Saudi)", bold=True, size=11, color=_GREEN)
        for jv in jv_partners:
            jvp = sa_c.add_paragraph()
            jvp.paragraph_format.space_before = Pt(1)
            jvp.paragraph_format.space_after  = Pt(1)
            jvp.paragraph_format.left_indent  = Cm(0.4)
            _add_run(jvp, f"• {jv}", size=10)

    # Major Projects
    projects = saudi.get("majorProjects", [])
    if projects:
        sa_mph = sa_c.add_paragraph()
        sa_mph.paragraph_format.space_before = Pt(6)
        sa_mph.paragraph_format.space_after  = Pt(2)
        _add_run(sa_mph, "Major Projects in Saudi Arabia", bold=True, size=11, color=_GREEN)
        for proj in projects:
            pp2 = sa_c.add_paragraph()
            pp2.paragraph_format.space_before = Pt(1)
            pp2.paragraph_format.space_after  = Pt(1)
            pp2.paragraph_format.left_indent  = Cm(0.4)
            _add_run(pp2, f"• {proj}", size=10)

    sa_c.add_paragraph().paragraph_format.space_after = Pt(4)

    # ── Latest News ───────────────────────────────────────────────────────────
    news_items = news or []
    if news_items:
        _section_head(doc, "Latest Market & Company News")
        news_tbl = doc.add_table(rows=len(news_items), cols=2)
        news_tbl.style = "Table Grid"
        news_tbl.autofit = False
        news_tbl.columns[0].width = Cm(2.1)
        news_tbl.columns[1].width = Cm(15.9)
        for ni, nitem in enumerate(news_items):
            nd_c, nt_c = news_tbl.rows[ni].cells
            fill = _LGREEN if ni % 2 == 0 else "FFFFFF"
            _cell_shading(nd_c, fill); _cell_shading(nt_c, fill)
            _no_borders(nd_c); _no_borders(nt_c)
            nd_c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            nt_c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            date_p = nd_c.add_paragraph()
            date_p.paragraph_format.space_before = Pt(3)
            date_p.paragraph_format.space_after  = Pt(3)
            _add_run(date_p, nitem.get("date", "")[:11], size=8, color=_GREY)
            title_p = nt_c.add_paragraph()
            title_p.paragraph_format.space_before = Pt(3)
            title_p.paragraph_format.space_after  = Pt(3)
            _add_run(title_p, nitem.get("title", ""), size=9)

    # ── Page 2 Footer ─────────────────────────────────────────────────────────
    div_p2 = doc.add_paragraph()
    div_p2.paragraph_format.space_before = Pt(8)
    div_p2.paragraph_format.space_after  = Pt(2)
    pPr_p2 = div_p2._p.get_or_add_pPr()
    pBdr_p2 = OxmlElement("w:pBdr")
    top_p2  = OxmlElement("w:top")
    top_p2.set(qn("w:val"),   "single"); top_p2.set(qn("w:sz"), "4")
    top_p2.set(qn("w:space"), "4");      top_p2.set(qn("w:color"), _MGREEN)
    pBdr_p2.append(top_p2); pPr_p2.append(pBdr_p2)

    ft2 = doc.add_table(rows=1, cols=2)
    ft2.style = "Table Grid"
    ft2.autofit = False
    ft2.columns[0].width = Cm(9.0)
    ft2.columns[1].width = Cm(9.0)
    f2l, f2r = ft2.rows[0].cells
    _no_borders(f2l); _no_borders(f2r)
    pf2l = f2l.add_paragraph()
    _add_run(pf2l, "Prepared by: Minister Outreach Office, MISA", size=8, color=_GREY)
    pf2r = f2r.add_paragraph()
    pf2r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _add_run(pf2r, "For internal use only – Ministry of Investment of Saudi Arabia", size=8, color=_GREY)

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

    inv_regions  = d.get("investmentRegions", [])
    subsidiaries = d.get("globalSubsidiaries", [])
    saudi        = d.get("saudiPresence", {})

    reg_html = "".join(
        f'<div class="ev-bul"><strong>{r["region"]}</strong></div>'
        f'<div class="ev-sub">{r["focus"]}</div>'
        for r in inv_regions
    ) or "<div style='color:#999;font-size:11px'>Data not available</div>"

    sub_html = "".join(
        f'<div class="ev-bul">{s}</div>' for s in subsidiaries
    ) or "<div style='color:#999;font-size:11px'>Data not available</div>"

    jv_html  = "".join(f'<div class="ev-bul">{jv}</div>' for jv in saudi.get("jvPartners", []))
    prj_html = "".join(f'<div class="ev-bul">{p}</div>' for p in saudi.get("majorProjects", []))

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
          <p style="font-size:12px;margin:0 0 8px">{
            f'Delegate this meeting to <strong>{_get_selected_delegate_name() or rec.get("delegateTo","")}</strong>, given:'
            if (_get_selected_delegate_name() or rec.get("delegateTo",""))
            else 'Select meeting nature in Decision &amp; Direction to confirm the recommended delegate.'
          }</p>
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

    <div class="ev-brief-wrap" style="margin-top:16px">
      <div class="ev-brief-hdr">
        <div>
          <h2>{d.get("company","")} — Investment Intelligence Brief</h2>
          <p>Global Presence · Saudi Contribution · Market Intelligence</p>
        </div>
        <div style="text-align:right">
          <div class="ev-conf">CONFIDENTIAL</div>
          <div style="color:#C8E6D4;font-size:11px;margin-top:4px">{today}</div>
        </div>
      </div>
      <div class="ev-brief-body">
        <div class="ev-shead">Global Investment Footprint</div>
        <div class="ev-cols">
          <div>
            <div style="font-weight:700;font-size:11px;color:#1B5C3F;margin-bottom:6px">Where They Invest</div>
            {reg_html}
          </div>
          <div>
            <div style="font-weight:700;font-size:11px;color:#1B5C3F;margin-bottom:6px">Global Entities &amp; Subsidiaries</div>
            {sub_html}
          </div>
        </div>

        <div class="ev-shead">Saudi Arabia Presence &amp; Contribution</div>
        <div class="ev-recbox">
          <h4>Investments &amp; Commitments</h4>
          <p style="font-size:11px;margin:4px 0 8px">{saudi.get("investments","No known current investments in Saudi Arabia")}</p>
          {"<h4 style='margin-top:10px'>Joint Venture Partners (Saudi)</h4>" + jv_html if jv_html else ""}
          {"<h4 style='margin-top:10px'>Major Projects in Saudi Arabia</h4>" + prj_html if prj_html else ""}
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
      <p>{"Delegate this meeting to <strong>" + (_get_selected_delegate_name() or rec.get("delegateTo","")) + "</strong>, given:"
         if (_get_selected_delegate_name() or rec.get("delegateTo",""))
         else "Select meeting nature in Decision &amp; Direction to confirm the recommended delegate."
      }</p>
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


_DELEGATES = {
    "challenges_deals": {
        "name":  "Assistant Minister Abdullah A. Aldubaikhi",
        "role":  "Challenges & Deals",
        "desc":  "Recommended for meetings focused on active investment deals, regulatory challenges, or situations requiring ministerial-level problem resolution.",
        "color": "#1B5C3F",
    },
    "exploration_events": {
        "name":  "CEO of Saudi Investment Promotion Authority — Khaled S. Alkhattaf",
        "role":  "Exploration & Events",
        "desc":  "Recommended for companies exploring the Saudi market for the first time, event participation, or early-stage investment roadshows.",
        "color": "#C9974A",
    },
    "sector_services": {
        "name":  "Assistant Deputy — Services Industry",
        "role":  "Sector Services",
        "desc":  "Recommended based on the company's sector, size, and strategic relevance to Saudi service industry priorities.",
        "color": "#2D7A54",
    },
    "senior_org": {
        "name":  "H.E. Assistant Minister Ibrahim Al-Mubarak",
        "role":  "Very Senior / Organisation Meeting",
        "desc":  "Recommended for meetings with very senior executives or high-profile organisations where ministerial-level representation is required.",
        "color": "#1D4ED8",
    },
    "minister_direct": {
        "name":  "H.E. The Minister of Investment",
        "role":  "Direct Ministerial Meeting",
        "desc":  "Reserved for engagements that meet the highest strategic threshold — sovereign-level counterparts, national-impact commitments, or state-to-state partnerships.",
        "color": "#7C1A1A",
        "minister_criteria": [
            "Counterpart is a Head of State, Minister, or CEO of a sovereign / Fortune 100 entity",
            "Engagement involves a national-scale commitment (SAR 1B+, MOU, or state partnership)",
            "The meeting outcome directly shapes Vision 2030 investment targets or MISA's mandate",
            "No assistant minister can represent MISA at the required protocol level",
        ],
    },
}

_TALKING_POINTS_BY_SECTOR = {
    "Technology": [
        "Saudi Arabia's National Digital Transformation Programme and MISA's ICT investment pipeline.",
        "Data centre incentives and cloud adoption targets under Vision 2030.",
        "NEOM, Diriyah, and mega-projects as test-beds for technology deployment.",
    ],
    "Healthcare": [
        "Saudi Arabia's Vision 2030 healthcare localisation targets (Saudisation and local manufacturing).",
        "National Health Transformation Programme and opportunities in digital health.",
        "Partnership models with MOH, Saudi Health Council, and existing healthcare city developments.",
    ],
    "Finance": [
        "SAMA and CMA regulatory environment — new licences and fintech sandbox opportunities.",
        "Riyadh Financial District as the regional financial hub and gateway for Gulf capital deployment.",
        "Saudi Arabia's sovereign wealth ecosystem (PIF, Sanabil) and co-investment opportunities.",
    ],
    "Infrastructure": [
        "SAR 1 trillion infrastructure pipeline across transport, logistics, and utilities.",
        "GIGA projects (NEOM, Red Sea, Qiddiya) requiring international infrastructure expertise.",
        "PPP framework and MISA's one-stop-shop for regulatory approvals.",
    ],
    "Real Estate": [
        "Premium residency programme and its link to property investment thresholds.",
        "Affordable housing demand — Vision 2030 target of 70% homeownership.",
        "Development opportunities across GIGA and new city projects.",
    ],
    "default": [
        "MISA's mandate to attract, retain, and grow foreign investment in Saudi Arabia.",
        "Vision 2030 pillars aligned to the company's sector and areas of expertise.",
        "Licensing, regulatory, and operational support available through MISA's one-stop-shop.",
        "Saudi Arabia's economic transformation: open markets, privatisation, and PPP opportunities.",
    ],
}

_EV_NATURE_KEY_MAP = {
    "Challenges or Active Deals":            "challenges_deals",
    "Exploration / Events / New Companies":  "exploration_events",
    "Based on Sector & Company Level":       "sector_services",
    "Very Senior / Organisation Meeting":    "senior_org",
    "Match Minister Criteria":               "minister_direct",
}


def _get_selected_delegate_name() -> str:
    """Return the delegate name chosen in the selectbox, or empty string if not yet selected."""
    sel = st.session_state.get("ev_nature", "")
    if sel and sel in _EV_NATURE_KEY_MAP:
        return _DELEGATES[_EV_NATURE_KEY_MAP[sel]]["name"]
    return ""


def _render_recommendation_mode(brief: dict):
    """Step 4 Option 1 — Recommendation Mode: suggest leadership level to delegate to."""
    st.session_state.pop("_last_direction_key", None)
    rec = brief.get("recommendation", {})

    st.markdown("**Select meeting nature to determine the appropriate leadership level:**")
    nature = st.selectbox(
        "Meeting nature",
        ["", "Challenges or Active Deals", "Exploration / Events / New Companies",
         "Based on Sector & Company Level", "Very Senior / Organisation Meeting",
         "Match Minister Criteria"],
        format_func=lambda x: "— Select meeting type to see recommendation —" if x == "" else x,
        key="ev_nature",
        label_visibility="collapsed",
    )
    if not nature:
        st.caption("Select a meeting type above to see the recommended leadership level.")
        return

    selected_key = _EV_NATURE_KEY_MAP[nature]
    dg = _DELEGATES[selected_key]

    col_dg, col_rat = st.columns([2, 3])
    with col_dg:
        st.markdown(f"""
        <div style="background:{dg['color']};color:#fff;border-radius:8px;padding:12px 14px;margin-top:6px;">
          <div style="font-size:9px;letter-spacing:.08em;opacity:.8;text-transform:uppercase;margin-bottom:4px;">Recommended Delegate</div>
          <div style="font-size:13px;font-weight:700;line-height:1.3;">{dg['name']}</div>
          <div style="font-size:10px;opacity:.85;margin-top:4px;">{dg['role']}</div>
        </div>
        """, unsafe_allow_html=True)
        st.caption(dg["desc"])

        # Rebuild the docx with the selected delegate so the download button is up to date
        _s = st.session_state
        if _s.get("ev_brief") and _s.get("_last_delegate") != dg["name"]:
            _s["_last_delegate"] = dg["name"]
            _brief_upd = dict(_s["ev_brief"])
            _brief_upd["recommendation"] = dict(_s["ev_brief"].get("recommendation", {}))
            _brief_upd["recommendation"]["delegateTo"] = dg["name"]
            _s["ev_docx"] = _build_docx(
                _brief_upd,
                photo_bytes=_s.get("ev_photo_bytes"),
                logo_bytes=_s.get("ev_logo_bytes"),
                attendees=_s.get("ev_attendees", ""),
                news=_s.get("ev_news", []),
                contact_email=_s.get("ev_email", ""),
                contact_phone=_s.get("ev_phone", ""),
            )
            st.rerun()  # refresh so the download button above picks up the new docx

    with col_rat:
        if dg.get("minister_criteria"):
            criteria_rows = "".join(
                f'<div style="display:flex;align-items:flex-start;gap:6px;margin-bottom:5px;">'
                f'<span style="color:#1D4ED8;font-weight:700;flex-shrink:0;">✓</span>'
                f'<span style="font-size:11px;color:#1E3A5F;">{c}</span>'
                f'</div>'
                for c in dg["minister_criteria"]
            )
            st.markdown(
                f'<div style="background:#EFF6FF;border:1px solid #BFDBFE;border-radius:8px;padding:12px 14px;margin-top:6px;">'
                f'<div style="font-size:9px;font-weight:700;letter-spacing:.06em;color:#1D4ED8;text-transform:uppercase;margin-bottom:8px;">Minister Criteria — Meeting Qualifies If:</div>'
                f'{criteria_rows}'
                f'</div>',
                unsafe_allow_html=True,
            )
        elif rec.get("rationale"):
            st.markdown("**Rationale from AI briefing:**")
            for pt in rec.get("rationale", []):
                st.markdown(f"- {pt}")


def _render_direction_mode(brief: dict):
    """Step 4 Option 2 — Direction Mode: minister has approved, assign stakeholder + talking points."""
    st.session_state.pop("_last_delegate", None)
    sector   = brief.get("sectors", [{}])[0].get("title", "") if brief.get("sectors") else ""
    company  = brief.get("company", "")
    subject  = brief.get("subject", "")
    dps      = brief.get("discussionPoints", [])

    st.markdown("**H.E. Minister Fahad Al-Saif has approved this meeting and will attend directly.**")

    a1, a2 = st.columns(2)
    with a1:
        assigned = st.text_input(
            "Meeting Host",
            value="H.E. Fahad Al-Saif, Minister of Investment",
            key="ev_dir_stakeholder",
        )
    with a2:
        meeting_date = st.text_input(
            "Scheduled Date",
            value=brief.get("visitDates", ""),
            key="ev_dir_date",
            placeholder="e.g. 25 June 2026",
        )

    # Talking points
    sector_key = next((k for k in _TALKING_POINTS_BY_SECTOR if k.lower() in sector.lower()), "default")
    tp_list = _TALKING_POINTS_BY_SECTOR[sector_key]

    st.markdown("**Predefined Talking Points:**")
    col_ai, col_pre = st.columns(2)
    with col_ai:
        st.markdown("*From AI Briefing:*")
        for pt in (dps or tp_list)[:4]:
            st.markdown(f"- {pt}")
    with col_pre:
        st.markdown(f"*MISA Standard — {sector or 'General'}:*")
        for pt in tp_list[:4]:
            st.markdown(f"- {pt}")

    if assigned:
        st.success(f"✅ Direction confirmed: **{assigned}** will host this meeting"
                   + (f" on **{meeting_date}**" if meeting_date else "") + ".")

        # Rebuild docx with Direction Mode content whenever host or date changes
        _s = st.session_state
        _dir_key = f"{assigned}|{meeting_date}"
        if _s.get("ev_brief") and _s.get("_last_direction_key") != _dir_key:
            _s["_last_direction_key"] = _dir_key
            _s["ev_docx"] = _build_docx(
                _s["ev_brief"],
                photo_bytes=_s.get("ev_photo_bytes"),
                logo_bytes=_s.get("ev_logo_bytes"),
                attendees=_s.get("ev_attendees", ""),
                news=_s.get("ev_news", []),
                contact_email=_s.get("ev_email", ""),
                contact_phone=_s.get("ev_phone", ""),
                meeting_mode="direction",
                direction_host=assigned,
            )
            st.rerun()  # refresh so the download button above picks up the new docx


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
                loaded = []
                for f in uploaded:
                    fb = f.read()
                    mt = _media_type(f.name, f.type)
                    loaded.append({"name": f.name, "bytes": fb, "media_type": mt})
                    # Auto-extract photo from bio if none uploaded yet
                    if not s.get("ev_photo_bytes") and not s.get("ev_auto_photo"):
                        extracted = _extract_photo_from_file(fb, mt)
                        if extracted:
                            s["ev_photo_bytes"] = extracted
                            s["ev_auto_photo"]  = True
                s["ev_loaded_files"] = loaded
            cols = st.columns(min(len(uploaded), 4))
            for i, f in enumerate(uploaded):
                cols[i % 4].success(f"📄 {f.name}")
            if s.get("ev_auto_photo"):
                st.caption("📸 Person photo auto-extracted from uploaded bio")

    # ── Step 2: Company details, attendees, media & API key ──────────────────
    with st.container(border=True):
        st.markdown('<p class="ev-section">Step 2 — Company details & settings</p>',
                    unsafe_allow_html=True)

        # ── Row A: Company contact details ────────────────────────────────────
        st.markdown("**Company Contact Details** — used to auto-fetch logo and enrich the document")
        cw, ce, cp = st.columns(3)
        s["ev_website"] = cw.text_input(
            "Company Website",
            value=s.get("ev_website", ""),
            key="ev_web",
            placeholder="e.g. blackrock.com  (no https://)",
        )
        s["ev_email"] = ce.text_input(
            "Contact Email",
            value=s.get("ev_email", ""),
            key="ev_email_inp",
            placeholder="e.g. name@company.com",
        )
        s["ev_phone"] = cp.text_input(
            "Contact Phone",
            value=s.get("ev_phone", ""),
            key="ev_phone_inp",
            placeholder="e.g. +966 11 000 0000",
        )

        st.markdown("---")

        # ── Row B: Context + Ministry Attendees ───────────────────────────────
        ctx_col, att_col = st.columns(2)
        with ctx_col:
            s["ev_context"] = st.text_area(
                "Additional context (optional)",
                value=s["ev_context"],
                height=110,
                key="ev_ctx",
                placeholder="e.g. Visitor arriving 20–22 June. Focus on logistics and data centres. Escalate if CEO of Fortune 100…",
            )
        with att_col:
            st.markdown("**Recommended Ministry Attendees** — will appear as a named box in the document")
            st.text_area(
                "One name per line",
                height=90,
                key="ev_attendees",
                label_visibility="collapsed",
                placeholder="H.E. Fahad Al-Saif, Minister of Investment\nH.E. Ibrahim Al-Rashed, Asst. Minister\nDr. Khalid Al-Falih, Adviser\n…",
            )

        st.markdown("---")

        # ── Row C: Visitor photo, company logo, API key ───────────────────────
        col_b, col_c, col_d = st.columns([1, 1, 2])
        with col_b:
            st.markdown("**Visitor photo** *(optional)*")
            st.caption("Upload manually, or auto-extracted from bio")
            photo_up = st.file_uploader("photo", type=["png","jpg","jpeg"],
                                        key="ev_photo_up", label_visibility="collapsed")
            if photo_up is not None:
                _pk = f"{photo_up.name}_{photo_up.size}"
                if _pk != s.get("ev_photo_key", ""):
                    s["ev_photo_key"]   = _pk
                    s["ev_photo_bytes"] = photo_up.read()
                    s["ev_auto_photo"]  = False
                st.image(s["ev_photo_bytes"], width=80)
            elif s.get("ev_photo_bytes") and s.get("ev_auto_photo"):
                st.image(s["ev_photo_bytes"], width=80)
                st.caption("📸 Auto-extracted from bio")
        with col_c:
            st.markdown("**Company logo** *(optional)*")
            st.caption("Upload manually, or auto-fetched from website")
            logo_up = st.file_uploader("logo", type=["png","jpg","jpeg"],
                                       key="ev_logo_up", label_visibility="collapsed")
            if logo_up is not None:
                _lk = f"{logo_up.name}_{logo_up.size}"
                if _lk != s.get("ev_logo_key", ""):
                    s["ev_logo_key"]   = _lk
                    s["ev_logo_bytes"] = logo_up.read()
                st.image(s["ev_logo_bytes"], width=80)
            elif s.get("ev_website") and not s.get("ev_logo_bytes"):
                st.caption("Logo will be auto-fetched on Generate")
        with col_d:
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

                # Domain: prefer manually entered website, fallback to Claude's extraction
                manual_domain = (s.get("ev_website") or "").strip().lower()
                manual_domain = manual_domain.replace("https://", "").replace("http://", "").split("/")[0]
                domain = manual_domain or brief.get("companyDomain", "")

                # Auto-fetch logo from Clearbit if not manually uploaded
                logo_bytes = s.get("ev_logo_bytes")
                if not logo_bytes and domain:
                    status.markdown(f"→ Fetching {brief.get('company','')} logo from {domain}…")
                    logo_bytes = _fetch_logo(domain)
                    if logo_bytes:
                        s["ev_logo_bytes"] = logo_bytes

                prog.progress(65)

                # Fetch latest news
                status.markdown("→ Fetching latest news…")
                news_items = _fetch_news_for_company(brief.get("company", ""), domain)
                s["ev_news"] = news_items

                prog.progress(80)

                status.markdown("→ Building 2-page briefing document…")
                docx_bytes = _build_docx(
                    brief,
                    photo_bytes=s.get("ev_photo_bytes"),
                    logo_bytes=logo_bytes,
                    attendees=s.get("ev_attendees", ""),
                    news=news_items,
                    contact_email=s.get("ev_email", ""),
                    contact_phone=s.get("ev_phone", ""),
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

        # ── Decision / Mode selector ──────────────────────────────────────────
        with st.container(border=True):
            st.markdown('<p class="ev-section">Decision & Direction</p>',
                        unsafe_allow_html=True)
            st.caption("Select how to proceed with this meeting.")

            mode = st.radio(
                "Mode",
                ["Recommendation Mode", "Direction Mode"],
                horizontal=True,
                key="ev_mode",
                label_visibility="collapsed",
            )

            if mode == "Recommendation Mode":
                _render_recommendation_mode(brief)
            else:
                _render_direction_mode(brief)

        st.markdown("---")
        _render_preview(brief)
