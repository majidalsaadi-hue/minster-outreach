
# Minister Meeting Brief — generate filled PPTX from template
# Sources: CRM data, uploaded documents, Claude API, internet

import io
import re
import json
import base64
import urllib.parse
import urllib.request
import http.client
from datetime import date
from pathlib import Path

import streamlit as st
import pandas as pd

from config.settings import MISA_GREEN, MISA_GOLD

_GREEN = "#1B5C3F"
_GOLD  = "#C9974A"
_RED   = "#DC2626"
_MODEL = "claude-sonnet-4-6"
_TEMPLATE = Path(__file__).parent.parent / "templates" / "minister_brief_template.pptx"

# ── Field registry ────────────────────────────────────────────────────────────
# (key, label, section, required)
_FIELDS = [
    # Company Profile
    ("company_name",      "Company Name",       "company",    True),
    ("website",           "Website",            "company",    False),
    ("year_founded",      "Year Founded",       "company",    False),
    ("description",       "Description",        "company",    False),
    ("sectors",           "Sectors Covered",    "company",    False),
    ("hq",                "Company HQ",         "company",    False),
    ("global_branches",   "Global Branches",    "company",    False),
    ("employee_count",    "Employee #",         "company",    False),
    ("revenue",           "Revenue",            "company",    False),
    ("company_size",      "Company Size",       "company",    False),
    ("ksa_presence",      "KSA Presence",       "company",    False),
    ("engaged_with_he",   "Engaged with HE",    "company",    False),
    ("engaged_with_misa", "Engaged with MISA",  "company",    False),
    # Leadership Profile
    ("full_name",         "Full Name",          "leadership", True),
    ("position",          "Position",           "leadership", True),
    ("biography",         "Biography",          "leadership", False),
    # Meeting Overview
    ("meeting_reason",    "Meeting Reason",     "meeting",    True),
    ("investor_presence", "Investor Presence",  "meeting",    False),
    ("in_saudi",          "In Saudi",           "meeting",    False),
    ("attendees",         "Attendees",          "meeting",    False),
    ("brief_owner",       "Brief Owner",        "meeting",    False),
    ("misa_rm",           "MISA RM",            "meeting",    False),
    ("incentives",        "Incentives",         "meeting",    False),
    ("special_needs",     "Special Arrangement","meeting",    False),
]
_FIELD_KEYS   = [f[0] for f in _FIELDS]
_FIELD_LABELS = {f[0]: f[1] for f in _FIELDS}
_REQUIRED     = {f[0] for f in _FIELDS if f[3]}


def _empty() -> dict:
    return {k: "" for k in _FIELD_KEYS}


# ── Session state ─────────────────────────────────────────────────────────────
def _init():
    defaults = {
        "mb_api_key":        "",
        "mb_data":           _empty(),
        "mb_found":          set(),
        "mb_photo":          None,   # bytes
        "mb_logo":           None,   # bytes
        "mb_active_company": "",
        "mb_ev_seeded_key":  "",     # tracks which ev_brief has been mapped → mb_data
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ── CSS ───────────────────────────────────────────────────────────────────────
def _css():
    st.markdown("""
    <style>
    .mb-section {background:#f9fafb;border:1px solid #e5e7eb;border-radius:10px;
                 padding:16px 18px;margin-bottom:14px;}
    .mb-section-title {font-size:12px;font-weight:700;letter-spacing:.6px;
                        text-transform:uppercase;color:#1B5C3F;margin-bottom:12px;}
    .mb-missing {background:#fef2f2;border:1px solid #fca5a5;border-radius:6px;
                  padding:6px 10px;font-size:12px;color:#991B1B;margin-bottom:6px;}
    .mb-found   {background:#f0fdf4;border:1px solid #86efac;border-radius:6px;
                  padding:6px 10px;font-size:12px;color:#065F46;margin-bottom:6px;}
    </style>
    """, unsafe_allow_html=True)


# ── CRM pre-fill ──────────────────────────────────────────────────────────────
def _prefill_from_crm(dfs: dict, company: str) -> tuple[dict, set]:
    data  = _empty()
    data["company_name"] = company   # always keep the typed name
    found = {"company_name"}
    if not company:
        return data, found

    inv = dfs.get("Investor Master", pd.DataFrame())
    if not inv.empty and "Company Name" in inv.columns:
        row = inv[inv["Company Name"].str.lower() == company.lower()]
        if not row.empty:
            r = row.iloc[0]
            _map = {
                "company_name":  "Company Name",
                "website":       "Website",
                "sectors":       "Sector",
                "hq":            "Country",
                "employee_count":"Company Size (Global)",
                "engaged_with_misa": "Relationship Status",
                "misa_rm":       "Relationship Manager",
                "full_name":     "Key Contact Name",
                "position":      "Key Contact Title",
            }
            for fk, col in _map.items():
                v = str(r.get(col, "") or "").strip()
                if v and v.lower() not in ("nan", "none", ""):
                    data[fk] = v
                    found.add(fk)

    # Company rep from investor
    rep  = str(inv[inv["Company Name"].str.lower() == company.lower()].iloc[0].get("Company Rep", "") if not inv.empty and "Company Name" in inv.columns and company.lower() in inv["Company Name"].str.lower().values else "").strip()
    rep_pos = str(inv[inv["Company Name"].str.lower() == company.lower()].iloc[0].get("Rep Position", "") if not inv.empty and "Company Name" in inv.columns and company.lower() in inv["Company Name"].str.lower().values else "").strip()
    if rep and not data.get("full_name"):
        data["full_name"] = rep; found.add("full_name")
    if rep_pos and not data.get("position"):
        data["position"] = rep_pos; found.add("position")

    # Meetings for attendees / meeting reason
    mtgs = dfs.get("Meeting Log", pd.DataFrame())
    if not mtgs.empty and "Company Name" in mtgs.columns:
        co_mtgs = mtgs[mtgs["Company Name"].str.lower() == company.lower()]
        if not co_mtgs.empty:
            latest = co_mtgs.sort_values("Meeting Date", ascending=False).iloc[0] if "Meeting Date" in co_mtgs.columns else co_mtgs.iloc[0]
            obj = str(latest.get("Meeting Objective", "") or "").strip()
            att = str(latest.get("MISA Attendees", "") or "").strip()
            if obj and not data.get("meeting_reason"):
                data["meeting_reason"] = obj; found.add("meeting_reason")
            if att and not data.get("attendees"):
                data["attendees"] = att; found.add("attendees")

    return data, found


# ── File text extraction ───────────────────────────────────────────────────────
def _extract_text_from_file(file_bytes: bytes, filename: str) -> tuple[str, bytes | None]:
    """Return (text_content, photo_bytes_or_None)."""
    ext = Path(filename).suffix.lower()
    photo = None
    text  = ""

    if ext in (".pptx", ".ppt"):
        try:
            from pptx import Presentation as _Prs
            prs = _Prs(io.BytesIO(file_bytes))
            parts = []
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        parts.append(shape.text_frame.text.strip())
            text = "\n".join(p for p in parts if p)
        except Exception as e:
            text = f"[PPT extraction failed: {e}]"

    elif ext in (".docx", ".doc"):
        try:
            import docx as _docx
            doc = _docx.Document(io.BytesIO(file_bytes))
            text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        except Exception as e:
            text = f"[DOCX extraction failed: {e}]"

    elif ext == ".pdf":
        # Return raw bytes — Claude can read PDFs directly
        text = "__PDF__"

    elif ext in (".jpg", ".jpeg", ".png", ".gif", ".webp"):
        # The file itself is a photo
        photo = file_bytes
        text  = "__IMAGE__"

    else:
        try:
            text = file_bytes.decode("utf-8", errors="replace")[:8000]
        except Exception:
            text = ""

    return text, photo


# ── Claude extraction ─────────────────────────────────────────────────────────
def _call_claude(api_key: str, messages: list, system: str = "") -> str:
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        kwargs = dict(model=_MODEL, max_tokens=2048, messages=messages)
        if system:
            kwargs["system"] = system
        resp = client.messages.create(**kwargs)
        return resp.content[0].text
    except Exception as e:
        return f"[error: {e}]"


def _extract_with_claude(api_key: str, files: list, current_data: dict) -> tuple[dict, bytes | None]:
    """Use Claude to extract structured data + photo from one or more uploaded files."""
    schema = {k: _FIELD_LABELS[k] for k in _FIELD_KEYS}
    system = (
        "You are an expert at extracting structured company and leadership information from documents. "
        "Return ONLY a valid JSON object with the exact keys provided. "
        "Use empty string for any value you cannot find. Be concise — no markdown."
    )
    prompt = (
        f"Extract information about this company/person and fill the following JSON schema:\n"
        f"{json.dumps(schema, indent=2)}\n\n"
        f"For 'company_size' use one of: Micro, Small, Medium, Large.\n"
        f"For 'ksa_presence', 'engaged_with_he', 'engaged_with_misa', 'global_branches' use Yes or No.\n"
        f"Already known values (do not override unless you find better data):\n"
        f"{json.dumps({k: v for k, v in current_data.items() if v}, indent=2)}\n\n"
        f"Document content:\n"
    )

    content_blocks = []
    photo = None

    for file_bytes, filename in files:
        text, file_photo = _extract_text_from_file(file_bytes, filename)
        ext = Path(filename).suffix.lower()
        if not photo and file_photo:
            photo = file_photo
        if text == "__PDF__":
            b64 = base64.standard_b64encode(file_bytes).decode()
            content_blocks.append({
                "type": "document",
                "source": {"type": "base64", "media_type": "application/pdf", "data": b64},
            })
        elif text == "__IMAGE__":
            mt = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png",
                  "gif": "image/gif", "webp": "image/webp"}.get(ext.lstrip("."), "image/png")
            b64 = base64.standard_b64encode(file_bytes).decode()
            content_blocks.append({
                "type": "image",
                "source": {"type": "base64", "media_type": mt, "data": b64},
            })
        else:
            content_blocks.append({
                "type": "text",
                "text": f"--- File: {filename} ---\n{text[:6000]}",
            })

    content_blocks.append({"type": "text", "text": prompt})
    messages = [{"role": "user", "content": content_blocks}]

    raw = _call_claude(api_key, messages, system)
    new_data = dict(current_data)
    try:
        m = re.search(r"\{[\s\S]*\}", raw)
        if m:
            extracted = json.loads(m.group())
            for k, v in extracted.items():
                if k in _FIELD_KEYS and str(v).strip() and str(v).strip().lower() not in ("", "null", "none"):
                    new_data[k] = str(v).strip()
    except Exception:
        pass

    return new_data, photo


# ── Internet search ────────────────────────────────────────────────────────────
def _fetch_logo(domain: str) -> bytes | None:
    if not domain:
        return None
    domain = re.sub(r"https?://", "", domain).split("/")[0].strip()
    try:
        url = f"https://logo.clearbit.com/{domain}"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=5) as r:
            return r.read() if r.status == 200 else None
    except Exception:
        return None


def _search_online(api_key: str, company: str, person: str, website: str, current_data: dict) -> dict:
    """Use DuckDuckGo + Claude to fill missing fields from the internet."""
    try:
        domain = re.sub(r"https?://", "", website or "").split("/")[0].strip()
        query  = urllib.parse.quote(f"{company} {person} investment company profile")
        url    = f"https://api.duckduckgo.com/?q={query}&format=json&no_redirect=1&no_html=1&skip_disambig=1"
        req    = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8) as r:
            ddg = json.loads(r.read().decode())
        abstract = ddg.get("AbstractText", "") or ddg.get("Answer", "")
        related  = " ".join(r.get("Text", "") for r in ddg.get("RelatedTopics", [])[:5])
        snippet  = (abstract + " " + related)[:3000]
    except Exception:
        snippet = ""

    if not snippet.strip():
        return current_data

    schema  = {k: _FIELD_LABELS[k] for k in _FIELD_KEYS if not current_data.get(k)}
    if not schema:
        return current_data

    prompt = (
        f"Using the web snippet below, fill these missing fields for company '{company}' / person '{person}'.\n"
        f"Schema: {json.dumps(schema)}\n\n"
        f"Return ONLY a JSON object with the keys above. Empty string if not found.\n\n"
        f"Web content:\n{snippet}"
    )
    raw = _call_claude(api_key, [{"role": "user", "content": prompt}])
    new_data = dict(current_data)
    try:
        m = re.search(r"\{[\s\S]*\}", raw)
        if m:
            for k, v in json.loads(m.group()).items():
                if k in _FIELD_KEYS and str(v).strip() and str(v).strip().lower() not in ("", "null", "none"):
                    if not new_data.get(k):
                        new_data[k] = str(v).strip()
    except Exception:
        pass
    return new_data


# ── PPTX generation ───────────────────────────────────────────────────────────
_SLIDE1_MAP = [
    # (first_line_prefix, top_inch, left_inch, field_key)
    ("Description",       1.73, 0.50, "description"),
    ("Sectors Covered",   3.14, 0.50, "sectors"),
    ("Company HQ",        4.44, 0.50, "hq"),
    ("Employee #",        5.98, 0.50, "employee_count"),
    ("Revenue",           5.98, 2.45, "revenue"),
    ("Global branches",   5.41, 0.50, "global_branches"),
    ("Engaged with HE",   6.56, 0.50, "engaged_with_he"),
    ("Engaged with MISA", 6.56, 2.45, "engaged_with_misa"),
    ("Full Name:",        1.73, 4.78, "full_name"),
    ("Position:",         3.00, 4.78, "position"),
    ("Biography:",        4.27, 4.78, "biography"),
    ("Meeting Reason",    1.73, 9.06, "meeting_reason"),
    ("Investor Presence", 2.35, 9.06, "investor_presence"),
    ("KSA Presence",      2.70, 9.21, "ksa_presence"),
    ("Incentives",        3.17, 9.21, "incentives"),
    ("In Saudi",          3.68, 9.06, "in_saudi"),
    ("Attendees",         5.49, 9.06, "attendees"),
    ("Brief Owner",       6.29, 9.06, "brief_owner"),
    ("MISA RM",           6.30, 10.84, "misa_rm"),
    ("Need special",      7.17, 0.54, "special_needs"),
]

_EMU = 914400  # 1 inch in EMU


def _near(shape_inch: float, target: float, tol: float = 0.18) -> bool:
    return abs(shape_inch - target) < tol


def _set_content(shape, value: str):
    """Keep first paragraph (label), replace subsequent paragraphs with value."""
    from pptx.util import Pt
    from pptx.oxml.ns import qn
    from lxml import etree

    tf = shape.text_frame
    tf.word_wrap = True

    # Remove all but first paragraph
    paras = tf.paragraphs
    while len(tf.paragraphs) > 1:
        p_elem = tf.paragraphs[-1]._p
        p_elem.getparent().remove(p_elem)

    # Add value as new paragraph(s)
    if not value:
        return
    for line in value.split("\n"):
        p = tf.add_paragraph()
        run = p.add_run()
        run.text = line
        try:
            run.font.size = Pt(9)
        except Exception:
            pass


def _replace_with_image(slide, rect_name: str, img_bytes: bytes, w_inch: float, h_inch: float):
    """Find a named rectangle placeholder and overlay it with a picture."""
    from pptx.util import Inches
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    target = None
    for shape in slide.shapes:
        if shape.name == rect_name:
            target = shape
            break
    if target is None:
        return

    left  = target.left
    top   = target.top
    width  = int(w_inch * _EMU)
    height = int(h_inch * _EMU)

    try:
        slide.shapes.add_picture(io.BytesIO(img_bytes), left, top, width, height)
    except Exception:
        pass


def _generate_pptx(data: dict, photo_bytes: bytes | None, logo_bytes: bytes | None) -> bytes:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor

    if not _TEMPLATE.exists():
        raise FileNotFoundError(f"Template not found: {_TEMPLATE}")

    prs  = Presentation(str(_TEMPLATE))
    sl1  = prs.slides[0]

    # ── Title (company name)
    for shape in sl1.shapes:
        if shape.name == "Title 1" and shape.has_text_frame:
            tf = shape.text_frame
            for para in tf.paragraphs:
                for run in para.runs:
                    run.text = ""
            tf.paragraphs[0].runs[0].text = data.get("company_name", "")
            break

    # ── Year Founded text box
    for shape in sl1.shapes:
        if shape.name == "TextBox 18" and shape.has_text_frame:
            tf = shape.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = f"Year Founded: {data.get('year_founded', '—')}"
            try:
                run.font.size = Pt(9)
            except Exception:
                pass
            break

    # ── Shape fill by position map
    for label_pfx, top_in, left_in, field in _SLIDE1_MAP:
        value = data.get(field, "")
        for shape in sl1.shapes:
            if not shape.has_text_frame:
                continue
            s_top  = shape.top  / _EMU
            s_left = shape.left / _EMU
            texts  = [p.text.strip() for p in shape.text_frame.paragraphs if p.text.strip()]
            if not texts:
                continue
            if texts[0].lower().startswith(label_pfx.lower()) and _near(s_top, top_in) and _near(s_left, left_in):
                _set_content(shape, value)
                break

    # ── Company size checkboxes (highlight selected)
    _SIZE_RECTS = {"Micro": "Rectangle 113", "Small": "Rectangle 114",
                   "Medium": "Rectangle 115", "Large": "Rectangle 116"}
    co_size = data.get("company_size", "")
    for label, rect_name in _SIZE_RECTS.items():
        for shape in sl1.shapes:
            if shape.name == rect_name:
                try:
                    from pptx.dml.color import RGBColor
                    fill = shape.fill
                    if label.lower() == co_size.lower():
                        fill.solid()
                        fill.fore_color.rgb = RGBColor(0x1B, 0x5C, 0x3F)
                    else:
                        fill.solid()
                        fill.fore_color.rgb = RGBColor(0xF3, 0xF4, 0xF6)
                except Exception:
                    pass

    # ── Photo
    if photo_bytes:
        _replace_with_image(sl1, "Rectangle 2", photo_bytes, 0.90, 0.90)

    # ── Logo
    logo = logo_bytes
    if not logo and data.get("website"):
        logo = _fetch_logo(data["website"])
    if not logo and data.get("company_name"):
        domain = data["company_name"].lower().replace(" ", "") + ".com"
        logo = _fetch_logo(domain)
    if logo:
        _replace_with_image(sl1, "Rectangle 3", logo, 0.90, 0.42)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ── Form rendering ─────────────────────────────────────────────────────────────
def _render_form(data: dict, found: set) -> dict:
    updated = dict(data)

    def _field(key: str, multiline: bool = False, cols_in=None):
        label  = _FIELD_LABELS[key]
        val    = updated.get(key, "")
        is_ok  = bool(val.strip()) and key in found
        is_req = key in _REQUIRED
        suffix = " *" if is_req else ""
        color  = _GREEN if is_ok else (_RED if (is_req and not val) else "#92400E")
        hint   = "✓ auto-filled" if is_ok else ("⚠ required — enter manually" if is_req and not val else ("⚠ not found — enter manually" if not val else ""))

        target = cols_in if cols_in else st
        target.markdown(
            f'<div style="font-size:11px;font-weight:600;color:{color};margin-bottom:2px;">'
            f'{label}{suffix} <span style="font-size:10px;font-weight:400;color:{color};">{hint}</span></div>',
            unsafe_allow_html=True,
        )
        if multiline:
            new_val = target.text_area(label, value=val, label_visibility="collapsed",
                                       height=100, key=f"mb_{key}")
        else:
            new_val = target.text_input(label, value=val, label_visibility="collapsed",
                                        key=f"mb_{key}")
        updated[key] = new_val
        return new_val

    c1, c2, c3 = st.columns(3)

    # ── Company Profile
    with c1:
        st.markdown('<div class="mb-section-title">Company Profile</div>', unsafe_allow_html=True)
        for key in ["company_name", "website", "year_founded", "sectors", "hq",
                    "global_branches", "employee_count", "revenue",
                    "company_size", "ksa_presence", "engaged_with_he", "engaged_with_misa"]:
            _field(key, multiline=(key in ("description", "sectors")))
        _field("description", multiline=True)

    # ── Leadership
    with c2:
        st.markdown('<div class="mb-section-title">Leadership Profile</div>', unsafe_allow_html=True)
        for key in ["full_name", "position"]:
            _field(key)
        _field("biography", multiline=True)

    # ── Meeting Overview
    with c3:
        st.markdown('<div class="mb-section-title">Meeting Overview</div>', unsafe_allow_html=True)
        for key in ["meeting_reason", "investor_presence", "ksa_presence",
                    "incentives", "in_saudi", "attendees",
                    "brief_owner", "misa_rm", "special_needs"]:
            if key not in ["ksa_presence"]:  # ksa_presence shown in company col already
                _field(key, multiline=(key in ("meeting_reason", "in_saudi", "attendees")))

    return updated


# ── Main render ───────────────────────────────────────────────────────────────
def render(dfs: dict, lang: str):
    _init()
    _css()

    st.markdown(
        f"<h2 style='color:{_GREEN};margin-bottom:2px;'>Minister Meeting Brief</h2>"
        f"<p style='color:#6b7280;font-size:13px;margin-top:0;'>"
        f"Generate a filled company brief PPTX for minister meetings — "
        f"auto-filled from CRM, uploaded documents, and internet search.</p>",
        unsafe_allow_html=True,
    )

    # ── Step 0: API key ───────────────────────────────────────────────────────
    with st.expander("⚙️ API Key (required for AI extraction & internet fill)", expanded=not st.session_state["mb_api_key"]):
        api_key = st.text_input("Anthropic API Key", value=st.session_state["mb_api_key"],
                                type="password", key="mb_api_key_input")
        if api_key:
            st.session_state["mb_api_key"] = api_key

    api_key = st.session_state["mb_api_key"]

    # ── Step 1: Company selection & CRM pre-fill ──────────────────────────────
    st.markdown("### 1 — Company Name")
    inv = dfs.get("Investor Master", pd.DataFrame())
    companies = sorted(inv["Company Name"].dropna().unique().tolist()) if not inv.empty and "Company Name" in inv.columns else []

    co_col, btn_col = st.columns([4, 1])
    with co_col:
        company_input = st.text_input(
            "Company name",
            placeholder="e.g. Barclays, CDJ Capital, Morgan Stanley…",
            key="mb_co_input",
            label_visibility="collapsed",
        )
    with btn_col:
        load_clicked = st.button("Load →", key="mb_load_btn", type="primary", use_container_width=True)

    # Show CRM matches as quick-pick chips (purely informational — clicking loads that company)
    if company_input and companies:
        matches = [c for c in companies if company_input.lower() in c.lower()][:5]
        if matches:
            st.markdown(
                '<div style="font-size:11px;color:#6b7280;margin-bottom:4px;">CRM matches — click to load:</div>',
                unsafe_allow_html=True,
            )
            chip_cols = st.columns(min(len(matches), 5))
            for i, m in enumerate(matches):
                if chip_cols[i].button(m, key=f"mb_chip_{i}"):
                    st.session_state["mb_active_company"] = m
                    data, found = _prefill_from_crm(dfs, m)
                    st.session_state["mb_data"]  = data
                    st.session_state["mb_found"] = found

    # Explicit Load button — no rerun needed, just update session state
    if load_clicked and company_input.strip():
        active = company_input.strip()
        st.session_state["mb_active_company"] = active
        data, found = _prefill_from_crm(dfs, active)
        st.session_state["mb_data"]  = data
        st.session_state["mb_found"] = found

    active_company = st.session_state.get("mb_active_company", "")
    data  = st.session_state["mb_data"]
    found = st.session_state["mb_found"]

    if not active_company:
        st.info("Type a company name above and click **Load →** to begin. The company does not need to be in the CRM.")
        return

    n_found   = sum(1 for k in _FIELD_KEYS if data.get(k))
    n_missing = sum(1 for k in _FIELD_KEYS if not data.get(k))
    st.markdown(
        f'<div style="background:#f0fdf4;border:1px solid #86efac;border-radius:8px;'
        f'padding:8px 14px;margin-bottom:12px;font-size:13px;">'
        f'<strong>{active_company}</strong> — '
        f'CRM pre-fill: <strong style="color:{_GREEN}">{len(found)} fields found</strong> · '
        f'<strong style="color:{_RED}">{n_missing} fields missing</strong>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ── Step 2: File upload ───────────────────────────────────────────────────
    st.markdown("### 2 — Upload Document (optional)")
    st.markdown(
        '<p style="font-size:12px;color:#6b7280;">Upload a company brief, profile, PPT or PDF — '
        'the AI will extract all available information automatically.</p>',
        unsafe_allow_html=True,
    )

    doc_col, photo_col = st.columns(2)

    with doc_col:
        doc_files = st.file_uploader(
            "Company documents (PDF, PPT, Word)",
            type=["pdf", "pptx", "ppt", "docx", "doc"],
            key="mb_doc_upload",
            accept_multiple_files=True,
        )
        if doc_files and api_key:
            if st.button(f"Extract from {len(doc_files)} file(s)", key="mb_extract_btn"):
                with st.spinner("Extracting with AI…"):
                    new_data, extracted_photo = _extract_with_claude(
                        api_key, [(f.read(), f.name) for f in doc_files], data
                    )
                new_found = found | {k for k in _FIELD_KEYS if new_data.get(k) and not data.get(k)}
                st.session_state["mb_data"]  = new_data
                st.session_state["mb_found"] = new_found
                if extracted_photo:
                    st.session_state["mb_photo"] = extracted_photo
                data  = new_data
                found = new_found
                st.success("✅ Extraction complete")
                st.rerun()
        elif doc_files and not api_key:
            st.warning("Add an API key above to enable AI extraction.")

    with photo_col:
        st.markdown("**Person Photo** (JPG / PNG)")
        photo_file = st.file_uploader(
            "Person photo", type=["jpg", "jpeg", "png"],
            key="mb_photo_upload", label_visibility="collapsed",
        )
        if photo_file:
            st.session_state["mb_photo"] = photo_file.read()
            st.image(st.session_state["mb_photo"], width=120)
        elif st.session_state["mb_photo"]:
            st.image(st.session_state["mb_photo"], width=120, caption="Current photo")
        else:
            st.markdown(
                f'<div class="mb-missing">No photo found — please upload one above '
                f'or it will be left blank in the brief.</div>',
                unsafe_allow_html=True,
            )

    # ── Step 3: Internet fill ─────────────────────────────────────────────────
    if api_key:
        missing_keys = [k for k in _FIELD_KEYS if not data.get(k)]
        if missing_keys:
            if st.button(f"🌐 Fill {len(missing_keys)} missing fields from internet", key="mb_web_btn"):
                with st.spinner("Searching the internet…"):
                    new_data = _search_online(
                        api_key, data.get("company_name", active_company),
                        data.get("full_name", ""), data.get("website", ""), data
                    )
                new_found = found | {k for k in _FIELD_KEYS if new_data.get(k) and not data.get(k)}
                st.session_state["mb_data"]  = new_data
                st.session_state["mb_found"] = new_found
                data  = new_data
                found = new_found
                st.rerun()

    # ── Step 4: Review & edit form ────────────────────────────────────────────
    st.markdown("### 3 — Review & Edit")
    st.markdown(
        '<p style="font-size:12px;color:#6b7280;">'
        '<span style="color:#991B1B;">⚠ Red fields</span> were not found automatically — please fill them manually. '
        '<span style="color:#065F46;">✓ Green fields</span> were auto-filled.</p>',
        unsafe_allow_html=True,
    )

    updated = _render_form(data, found)
    st.session_state["mb_data"] = updated

    # ── Step 5: Generate PPTX ─────────────────────────────────────────────────
    st.markdown("### 4 — Generate Brief")

    missing_req = [_FIELD_LABELS[k] for k in _REQUIRED if not updated.get(k)]
    if missing_req:
        st.warning(f"Required fields still missing: {', '.join(missing_req)}")

    gen_col, _ = st.columns([2, 3])
    with gen_col:
        if st.button("📊 Generate Minister Brief PPTX", type="primary", key="mb_gen_btn"):
            with st.spinner("Filling PPTX template…"):
                try:
                    pptx_bytes = _generate_pptx(
                        updated,
                        photo_bytes=st.session_state.get("mb_photo"),
                        logo_bytes=st.session_state.get("mb_logo"),
                    )
                    co_slug = re.sub(r"[^\w]", "_", updated.get("company_name", "Brief"))
                    fname   = f"MinisterBrief_{co_slug}_{date.today().strftime('%Y%m%d')}.pptx"
                    st.download_button(
                        "⬇️ Download Brief",
                        data=pptx_bytes,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key="mb_download",
                    )
                    st.success("✅ Brief ready — click Download above.")
                except Exception as e:
                    st.error(f"Failed to generate: {e}")


# ── Embedded render (called from Evaluation & Briefing tab) ───────────────────

def _ev_to_mb(ev_brief: dict) -> dict:
    """Map Evaluation & Briefing extracted JSON to Minister Brief data format."""
    data = _empty()
    data["company_name"]   = ev_brief.get("company", "")
    data["website"]        = ev_brief.get("companyDomain", "")
    data["revenue"]        = ev_brief.get("revenue", "")
    data["employee_count"] = ev_brief.get("employees", "")
    data["description"]    = ev_brief.get("strategicContext", "")
    data["full_name"]      = ev_brief.get("visitorName", "")
    data["position"]       = ev_brief.get("visitorTitle", "")
    data["meeting_reason"] = ev_brief.get("subject", "")
    data["attendees"]      = ev_brief.get("accompaniedBy", "")
    sectors = ev_brief.get("sectors", [])
    data["sectors"]        = ", ".join(s.get("title", "") for s in sectors if s.get("title"))
    saudi = ev_brief.get("saudiPresence", {})
    data["ksa_presence"]   = saudi.get("investments", "")
    projects               = saudi.get("majorProjects", [])
    data["in_saudi"]       = "; ".join(projects) if projects else ""
    return data


def render_embedded(dfs: dict, lang: str):
    """
    Render Minister Meeting Brief as the second tab inside Evaluation & Briefing.
    Auto-seeds fields from the ev_brief that was already extracted, then lets the
    user fill gaps via document upload, internet search, or manual entry.
    """
    import os as _os
    _init()
    _css()

    # ── Seed from ev_brief whenever it changes ──────────────────────────────
    ev_brief = st.session_state.get("ev_brief")
    seeded   = st.session_state.get("mb_ev_seeded_key", "")
    if ev_brief:
        ev_key = ev_brief.get("company", "") + "|" + ev_brief.get("visitorName", "")
        if ev_key != seeded:
            # CRM data first, then ev_brief on top (evaluation extraction wins)
            crm_data, crm_found = _prefill_from_crm(dfs or {}, ev_brief.get("company", ""))
            ev_data = _ev_to_mb(ev_brief)
            for k, v in ev_data.items():
                if v:
                    crm_data[k] = v
                    crm_found.add(k)
            # Also pull the photo extracted in Evaluation & Briefing
            if st.session_state.get("ev_photo_bytes") and not st.session_state.get("mb_photo"):
                st.session_state["mb_photo"] = st.session_state["ev_photo_bytes"]
            st.session_state["mb_data"]           = crm_data
            st.session_state["mb_found"]          = crm_found
            st.session_state["mb_active_company"] = ev_brief.get("company", "")
            st.session_state["mb_ev_seeded_key"]  = ev_key

    active_company = st.session_state.get("mb_active_company", "")
    data  = st.session_state["mb_data"]
    found = st.session_state["mb_found"]

    # ── If no company loaded yet: show company input (works with or without ev_brief) ──
    if not active_company:
        ev_available = bool(st.session_state.get("ev_brief"))
        if ev_available:
            st.info("Loading data from Evaluation & Briefing…")
            return
        else:
            st.markdown(
                '<p style="font-size:12px;color:#6b7280;margin-bottom:10px;">'
                'After generating in <strong>Evaluation &amp; Briefing</strong>, all fields auto-populate here. '
                'Or enter a company name below to start the brief manually.</p>',
                unsafe_allow_html=True,
            )
            inv = (dfs or {}).get("Investor Master", pd.DataFrame())
            companies = (
                sorted(inv["Company Name"].dropna().unique().tolist())
                if not inv.empty and "Company Name" in inv.columns
                else []
            )
            co_col, btn_col = st.columns([4, 1])
            with co_col:
                company_input = st.text_input(
                    "Company name",
                    placeholder="e.g. Barclays, Morgan Stanley, ACWA Power…",
                    key="mb_emb_co_input",
                    label_visibility="collapsed",
                )
            with btn_col:
                load_clicked = st.button(
                    "Load →", key="mb_emb_load_btn", type="primary", use_container_width=True
                )

            if company_input and companies:
                matches = [c for c in companies if company_input.lower() in c.lower()][:5]
                if matches:
                    st.markdown(
                        '<div style="font-size:11px;color:#6b7280;margin-bottom:4px;">CRM matches — click to load:</div>',
                        unsafe_allow_html=True,
                    )
                    chip_cols = st.columns(min(len(matches), 5))
                    for i, m in enumerate(matches):
                        if chip_cols[i].button(m, key=f"mb_emb_chip_{i}"):
                            d, f = _prefill_from_crm(dfs or {}, m)
                            st.session_state["mb_data"]           = d
                            st.session_state["mb_found"]          = f
                            st.session_state["mb_active_company"] = m
                            st.rerun()

            if load_clicked and st.session_state.get("mb_emb_co_input", "").strip():
                active = st.session_state["mb_emb_co_input"].strip()
                d, f = _prefill_from_crm(dfs or {}, active)
                st.session_state["mb_data"]           = d
                st.session_state["mb_found"]          = f
                st.session_state["mb_active_company"] = active
                st.rerun()

            # Nothing loaded yet — show input only
            return

    # ── API key (inherited from Evaluation & Briefing session) ─────────────
    api_key = (
        st.session_state.get("ev_api_key", "")
        or st.session_state.get("mb_api_key", "")
        or _os.environ.get("ANTHROPIC_API_KEY", "")
    ).strip()

    if not api_key:
        with st.expander("⚙️ API Key (required for AI extraction & internet fill)", expanded=True):
            k = st.text_input("Anthropic API Key", type="password", key="mb_emb_api_key_inp")
            if k:
                st.session_state["mb_api_key"] = k
                api_key = k

    # ── Status banner ───────────────────────────────────────────────────────
    n_found   = len(found)
    n_missing = sum(1 for k in _FIELD_KEYS if not data.get(k))
    ev_source = " + Evaluation & Briefing" if st.session_state.get("mb_ev_seeded_key") else ""
    st.markdown(
        f'<div style="background:#f0fdf4;border:1px solid #86efac;border-radius:8px;'
        f'padding:8px 14px;margin-bottom:14px;font-size:13px;">'
        f'<strong>{active_company}</strong> — auto-filled from CRM{ev_source}: '
        f'<strong style="color:{_GREEN}">{n_found} fields found</strong> · '
        f'<strong style="color:{_RED}">{n_missing} fields still missing</strong>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ── 1 — Upload additional document (optional) ───────────────────────────
    st.markdown("### 1 — Upload Additional Document (optional)")
    st.markdown(
        '<p style="font-size:12px;color:#6b7280;">Upload a company brief, PPT, PDF or Word doc '
        'to fill any remaining gaps automatically.</p>',
        unsafe_allow_html=True,
    )

    doc_col, photo_col = st.columns(2)

    with doc_col:
        doc_files = st.file_uploader(
            "Company documents (PDF, PPT, Word)",
            type=["pdf", "pptx", "ppt", "docx", "doc"],
            key="mb_emb_doc_upload",
            accept_multiple_files=True,
        )
        if doc_files and api_key:
            if st.button(f"Extract from {len(doc_files)} file(s)", key="mb_emb_extract_btn"):
                with st.spinner("Extracting with AI…"):
                    new_data, extracted_photo = _extract_with_claude(
                        api_key, [(f.read(), f.name) for f in doc_files], data
                    )
                new_found = found | {k for k in _FIELD_KEYS if new_data.get(k) and not data.get(k)}
                st.session_state["mb_data"]  = new_data
                st.session_state["mb_found"] = new_found
                if extracted_photo:
                    st.session_state["mb_photo"] = extracted_photo
                data  = new_data
                found = new_found
                st.success("✅ Extraction complete")
                st.rerun()
        elif doc_files and not api_key:
            st.warning("Enter an API key above to enable AI extraction.")

    with photo_col:
        st.markdown("**Person Photo** (JPG / PNG)")
        photo_file = st.file_uploader(
            "Person photo", type=["jpg", "jpeg", "png"],
            key="mb_emb_photo_upload", label_visibility="collapsed",
        )
        if photo_file:
            st.session_state["mb_photo"] = photo_file.read()
            st.image(st.session_state["mb_photo"], width=120)
        elif st.session_state.get("mb_photo"):
            src = "from evaluation brief" if st.session_state.get("ev_auto_photo") else "current photo"
            st.image(st.session_state["mb_photo"], width=120, caption=f"📸 {src}")
        else:
            st.markdown(
                '<div class="mb-missing">No photo found — upload one above '
                'or it will be left blank in the brief.</div>',
                unsafe_allow_html=True,
            )

    # ── 2 — Fill gaps from internet ─────────────────────────────────────────
    if api_key:
        missing_keys = [k for k in _FIELD_KEYS if not data.get(k)]
        if missing_keys:
            if st.button(
                f"🌐 Fill {len(missing_keys)} missing fields from internet",
                key="mb_emb_web_btn",
            ):
                with st.spinner("Searching the internet…"):
                    new_data = _search_online(
                        api_key, data.get("company_name", active_company),
                        data.get("full_name", ""), data.get("website", ""), data
                    )
                new_found = found | {k for k in _FIELD_KEYS if new_data.get(k) and not data.get(k)}
                st.session_state["mb_data"]  = new_data
                st.session_state["mb_found"] = new_found
                data  = new_data
                found = new_found
                st.rerun()

    # ── 3 — Review & edit ───────────────────────────────────────────────────
    st.markdown("### 2 — Review & Edit")
    st.markdown(
        '<p style="font-size:12px;color:#6b7280;">'
        '<span style="color:#991B1B;">⚠ Red fields</span> were not found — fill manually. '
        '<span style="color:#065F46;">✓ Green fields</span> were auto-filled.</p>',
        unsafe_allow_html=True,
    )
    updated = _render_form(data, found)
    st.session_state["mb_data"] = updated

    # ── 4 — Generate PPTX ───────────────────────────────────────────────────
    st.markdown("### 3 — Generate Brief")
    missing_req = [_FIELD_LABELS[k] for k in _REQUIRED if not updated.get(k)]
    if missing_req:
        st.warning(f"Required fields still missing: {', '.join(missing_req)}")

    gen_col, _ = st.columns([2, 3])
    with gen_col:
        if st.button("📊 Generate Minister Brief PPTX", type="primary", key="mb_emb_gen_btn"):
            with st.spinner("Filling PPTX template…"):
                try:
                    logo_bytes = (
                        st.session_state.get("mb_logo")
                        or st.session_state.get("ev_logo_bytes")
                    )
                    pptx_bytes = _generate_pptx(
                        updated,
                        photo_bytes=st.session_state.get("mb_photo"),
                        logo_bytes=logo_bytes,
                    )
                    co_slug = re.sub(r"[^\w]", "_", updated.get("company_name", "Brief"))
                    fname   = f"MinisterBrief_{co_slug}_{date.today().strftime('%Y%m%d')}.pptx"
                    st.download_button(
                        "⬇️ Download Brief",
                        data=pptx_bytes,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key="mb_emb_download",
                    )
                    st.success("✅ Brief ready — click Download above.")
                except Exception as e:
                    st.error(f"Failed to generate: {e}")
