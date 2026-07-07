
# Arabic Meeting Minutes Generator — محضر الاجتماع
# Matches the MISA MoM template (green/gold brand, RTL Arabic output)

import io
import base64
import streamlit as st
from datetime import date

_GREEN = "#1B5C3F"
_GOLD  = "#C9974A"

# ─────────────────────────────────────────────────────────────────────────────
# Session-state keys
# ─────────────────────────────────────────────────────────────────────────────
_SK = {
    "attendees":    "mom_attendees",
    "actions":      "mom_actions",
    "disc_points":  "mom_disc_points",
    "next_steps":   "mom_next_steps",
}

def _init_state():
    defaults = {
        "mom_attendees":   [{"name": "", "role": ""}],
        "mom_actions":     [{"num": "1", "item": "", "owner": "", "deliverable": "", "due": "", "measure": ""}],
        "mom_disc_points": [""],
        "mom_next_steps":  [""],
        # metadata
        "mom_date":        date.today().strftime("%d %B %Y"),
        "mom_ref":         "",
        "mom_subject":     "",
        "mom_objective":   "",
        "mom_priority":    "",
        "mom_en_notes":    "",
        "mom_prepared_by": "Minister Outreach Office, MISA",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ─────────────────────────────────────────────────────────────────────────────
# Main entry point
# ─────────────────────────────────────────────────────────────────────────────
def render(dfs: dict, lang: str):
    _init_state()

    st.markdown(
        f"<h2 style='color:{_GREEN};margin-bottom:2px;'>📄 Meeting Minutes Builder</h2>"
        f"<p style='color:#6b7280;font-size:13px;margin-top:0;'>"
        f"محضر الاجتماع — Arabic MoM Generator for MISA</p>",
        unsafe_allow_html=True,
    )

    tab_a, tab_b, tab_c = st.tabs([
        "A  —  English Summary",
        "B  —  Meeting Details",
        "C  —  Generate Minutes",
    ])

    with tab_a:
        _render_step_a()

    with tab_b:
        _render_step_b()

    with tab_c:
        _render_step_c()


# ─────────────────────────────────────────────────────────────────────────────
# Step A — English meeting summary (type/paste OR upload)
# ─────────────────────────────────────────────────────────────────────────────
def _render_step_a():
    st.markdown(
        "<p style='font-size:13px;color:#374151;margin-bottom:12px;'>"
        "Provide your English meeting notes by <b>typing / pasting</b> or "
        "<b>uploading a file</b> (.txt or .docx). The content will be included "
        "as an English page in bilingual output mode.</p>",
        unsafe_allow_html=True,
    )

    input_mode = st.radio(
        "Input method",
        ["✏️  Type / Paste", "📎  Upload File (.txt or .docx)"],
        horizontal=True,
        key="mom_input_mode",
        label_visibility="collapsed",
    )

    if "Upload" in input_mode:
        uploaded = st.file_uploader(
            "Upload meeting notes file",
            type=["txt", "docx", "pdf", "pptx", "xlsx", "csv", "md", "rtf"],
            key="mom_notes_file",
            label_visibility="collapsed",
            help="Supported: .txt .docx .pdf .pptx .xlsx .csv .md .rtf",
        )
        if uploaded is not None:
            extracted = _extract_text_from_file(uploaded)
            if extracted:
                st.session_state["mom_en_notes"] = extracted
                st.success(
                    f"✅ Extracted {len(extracted.splitlines())} lines from **{uploaded.name}**"
                )

        # Still show a read-only preview of whatever was loaded
        if st.session_state.get("mom_en_notes"):
            st.markdown(
                "<p style='font-size:12px;color:#6B7280;margin-top:10px;margin-bottom:4px;'>"
                "Extracted text (edit if needed):</p>",
                unsafe_allow_html=True,
            )
            edited = st.text_area(
                "Extracted text",
                value=st.session_state["mom_en_notes"],
                height=260,
                key="mom_en_notes_upload_edit",
                label_visibility="collapsed",
            )
            st.session_state["mom_en_notes"] = edited
        else:
            st.info("Upload a .txt or .docx file above to load your meeting notes.")
    else:
        edited = st.text_area(
            "English Meeting Notes",
            value=st.session_state.get("mom_en_notes", ""),
            height=300,
            placeholder=(
                "Paste or type your English meeting summary here…\n\n"
                "Example:\n"
                "Meeting date: 29 June 2026\n"
                "Subject: Sovereign AI Opportunity\n"
                "Key points discussed:\n"
                "• Sovereign AI vision to reduce dependency on foreign AI providers\n"
                "• Government-level sponsorship required\n"
                "Action items: Siddharth to submit technical proposal before next meeting"
            ),
            key="mom_en_notes_paste",
            label_visibility="collapsed",
        )
        st.session_state["mom_en_notes"] = edited

    st.info(
        "**Tip:** Fill in the meeting details in **Tab B** to build the structured Arabic document. "
        "Tab C lets you generate Arabic-only or both English + Arabic output."
    )


def _extract_text_from_file(uploaded_file) -> str:
    """Extract plain text from any supported file format."""
    import io as _io
    name = uploaded_file.name.lower()
    raw  = uploaded_file.read()

    try:
        # ── Plain text variants ───────────────────────────────────────────────
        if name.endswith((".txt", ".md", ".rtf")):
            try:
                return raw.decode("utf-8")
            except UnicodeDecodeError:
                return raw.decode("latin-1", errors="replace")

        # ── Word document (.docx) ─────────────────────────────────────────────
        elif name.endswith(".docx"):
            from docx import Document
            doc   = Document(_io.BytesIO(raw))
            lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            for table in doc.tables:
                for row in table.rows:
                    cells = [c.text.strip() for c in row.cells if c.text.strip()]
                    if cells:
                        lines.append(" | ".join(cells))
            return "\n".join(lines)

        # ── PDF (.pdf) ────────────────────────────────────────────────────────
        elif name.endswith(".pdf"):
            try:
                from pypdf import PdfReader
            except ImportError:
                st.error("pypdf not installed. Run: pip install pypdf")
                return ""
            reader = PdfReader(_io.BytesIO(raw))
            pages  = []
            for i, page in enumerate(reader.pages):
                text = page.extract_text() or ""
                if text.strip():
                    pages.append(f"--- Page {i+1} ---\n{text.strip()}")
            return "\n\n".join(pages)

        # ── PowerPoint (.pptx) ────────────────────────────────────────────────
        elif name.endswith(".pptx"):
            from pptx import Presentation
            prs   = Presentation(_io.BytesIO(raw))
            lines = []
            for slide_num, slide in enumerate(prs.slides, 1):
                slide_texts = []
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            t = para.text.strip()
                            if t:
                                slide_texts.append(t)
                if slide_texts:
                    lines.append(f"--- Slide {slide_num} ---")
                    lines.extend(slide_texts)
            return "\n".join(lines)

        # ── Excel (.xlsx / .xls) ──────────────────────────────────────────────
        elif name.endswith((".xlsx", ".xls")):
            import pandas as pd
            xl     = pd.ExcelFile(_io.BytesIO(raw))
            blocks = []
            for sheet in xl.sheet_names:
                df = xl.parse(sheet).fillna("").astype(str)
                df = df[df.apply(lambda r: r.str.strip().any(), axis=1)]
                if df.empty:
                    continue
                blocks.append(f"--- Sheet: {sheet} ---")
                blocks.append(df.to_string(index=False))
            return "\n\n".join(blocks)

        # ── CSV (.csv) ────────────────────────────────────────────────────────
        elif name.endswith(".csv"):
            import pandas as pd
            try:
                text = raw.decode("utf-8")
            except UnicodeDecodeError:
                text = raw.decode("latin-1", errors="replace")
            import io as _sio
            df = pd.read_csv(_sio.StringIO(text)).fillna("").astype(str)
            return df.to_string(index=False)

        else:
            st.warning(f"Unsupported file type: {name.split('.')[-1]}")

    except Exception as e:
        st.error(f"Could not read **{uploaded_file.name}**: {e}")

    return ""


# ─────────────────────────────────────────────────────────────────────────────
# Step B — Structured meeting details (Arabic input fields)
# ─────────────────────────────────────────────────────────────────────────────
def _render_step_b():
    st.markdown(
        "<p style='font-size:13px;color:#374151;margin-bottom:16px;'>"
        "Enter all meeting details below. Fields accept Arabic text directly. "
        "All inputs build the structured محضر الاجتماع document.</p>",
        unsafe_allow_html=True,
    )

    # ── Metadata ──────────────────────────────────────────────────────────────
    _section_header("📋 Meeting Metadata", "بيانات الاجتماع")

    mc1, mc2, mc3 = st.columns(3)
    st.session_state["mom_date"] = mc1.text_input(
        "Date / التاريخ", value=st.session_state["mom_date"], key="mom_date_inp"
    )
    st.session_state["mom_ref"] = mc2.text_input(
        "Reference No. / رقم المرجع",
        value=st.session_state.get("mom_ref", ""),
        placeholder="MISA/Company/2026-06",
        key="mom_ref_inp",
    )
    st.session_state["mom_prepared_by"] = mc3.text_input(
        "Prepared by / أعده",
        value=st.session_state.get("mom_prepared_by", "Minister Outreach Office, MISA"),
        key="mom_prep_inp",
    )

    st.session_state["mom_subject"] = st.text_input(
        "Subject / الموضوع",
        value=st.session_state.get("mom_subject", ""),
        placeholder="e.g. فرصة الذكاء الاصطناعي السيادي للمملكة العربية السعودية",
        key="mom_subj_inp",
    )

    # ── Attendees ─────────────────────────────────────────────────────────────
    _section_header("👥 Attendees", "الحضور")

    attendees = st.session_state["mom_attendees"]
    for i, att in enumerate(attendees):
        ac1, ac2, ac3 = st.columns([3, 3, 0.6])
        attendees[i]["name"] = ac1.text_input(
            "Name / الاسم", value=att["name"],
            key=f"att_name_{i}", label_visibility="collapsed" if i > 0 else "visible",
            placeholder="الاسم الكامل"
        )
        attendees[i]["role"] = ac2.text_input(
            "Role / المنصب", value=att["role"],
            key=f"att_role_{i}", label_visibility="collapsed" if i > 0 else "visible",
            placeholder="المسمى الوظيفي / الدور"
        )
        if ac3.button("✕", key=f"del_att_{i}", help="Remove row") and len(attendees) > 1:
            attendees.pop(i)
            st.rerun()

    if st.button("+ Add Attendee", key="add_att"):
        attendees.append({"name": "", "role": ""})
        st.rerun()

    # ── Meeting Objective ─────────────────────────────────────────────────────
    _section_header("🎯 Meeting Objective", "هدف الاجتماع")

    st.session_state["mom_objective"] = st.text_area(
        "Objective / الهدف",
        value=st.session_state.get("mom_objective", ""),
        height=90,
        placeholder="مناقشة المبادرة المقترحة وتقييم مدى انطباقها…",
        key="mom_obj_inp",
        label_visibility="collapsed",
    )

    # ── Key Discussion Points ─────────────────────────────────────────────────
    _section_header("💬 Key Discussion Points", "محاور النقاش الرئيسية")

    disc = st.session_state["mom_disc_points"]
    for i, pt in enumerate(disc):
        dc1, dc2 = st.columns([11, 1])
        disc[i] = dc1.text_input(
            f"Point {i+1}", value=pt,
            key=f"disc_{i}", label_visibility="collapsed",
            placeholder=f"النقطة {i+1}…",
        )
        if dc2.button("✕", key=f"del_disc_{i}", help="Remove") and len(disc) > 1:
            disc.pop(i)
            st.rerun()

    if st.button("+ Add Discussion Point", key="add_disc"):
        disc.append("")
        st.rerun()

    # ── Action Items ──────────────────────────────────────────────────────────
    _section_header("✅ Action Items", "بنود العمل")

    actions = st.session_state["mom_actions"]
    # Header labels once
    ah1, ah2, ah3, ah4, ah5, ah6, _del = st.columns([0.4, 3.2, 2, 2, 2, 2, 0.5])
    ah1.markdown("<small style='color:#6B7280;'>#</small>", unsafe_allow_html=True)
    ah2.markdown("<small style='color:#6B7280;'>Action Item / بند العمل</small>", unsafe_allow_html=True)
    ah3.markdown("<small style='color:#6B7280;'>Owner / المسؤول</small>", unsafe_allow_html=True)
    ah4.markdown("<small style='color:#6B7280;'>Deliverable / المخرج</small>", unsafe_allow_html=True)
    ah5.markdown("<small style='color:#6B7280;'>Due Date / الموعد</small>", unsafe_allow_html=True)
    ah6.markdown("<small style='color:#6B7280;'>Success Measure / مقياس النجاح</small>", unsafe_allow_html=True)

    for i, act in enumerate(actions):
        ac1, ac2, ac3, ac4, ac5, ac6, ac_del = st.columns([0.4, 3.2, 2, 2, 2, 2, 0.5])
        actions[i]["num"] = ac1.text_input("", value=act["num"], key=f"act_num_{i}",
                                            label_visibility="collapsed")
        actions[i]["item"] = ac2.text_input("", value=act["item"],
                                             key=f"act_item_{i}", label_visibility="collapsed",
                                             placeholder="وصف البند…")
        actions[i]["owner"] = ac3.text_input("", value=act["owner"],
                                              key=f"act_owner_{i}", label_visibility="collapsed",
                                              placeholder="المسؤول")
        actions[i]["deliverable"] = ac4.text_input("", value=act["deliverable"],
                                                    key=f"act_del_{i}", label_visibility="collapsed",
                                                    placeholder="المخرج")
        actions[i]["due"] = ac5.text_input("", value=act["due"],
                                            key=f"act_due_{i}", label_visibility="collapsed",
                                            placeholder="قبل الاجتماع التالي")
        actions[i]["measure"] = ac6.text_input("", value=act["measure"],
                                                key=f"act_meas_{i}", label_visibility="collapsed",
                                                placeholder="مقياس النجاح")
        if ac_del.button("✕", key=f"del_act_{i}", help="Remove") and len(actions) > 1:
            actions.pop(i)
            st.rerun()

    if st.button("+ Add Action Item", key="add_act"):
        actions.append({"num": str(len(actions) + 1), "item": "", "owner": "",
                        "deliverable": "", "due": "", "measure": ""})
        st.rerun()

    # ── Next Steps ────────────────────────────────────────────────────────────
    _section_header("🔜 Summary of Next Steps", "ملخص الخطوات التالية")

    steps = st.session_state["mom_next_steps"]
    for i, step in enumerate(steps):
        sc1, sc2 = st.columns([11, 1])
        steps[i] = sc1.text_input(
            f"Step {i+1}", value=step,
            key=f"step_{i}", label_visibility="collapsed",
            placeholder=f"الخطوة {i+1}…",
        )
        if sc2.button("✕", key=f"del_step_{i}", help="Remove") and len(steps) > 1:
            steps.pop(i)
            st.rerun()

    if st.button("+ Add Next Step", key="add_step"):
        steps.append("")
        st.rerun()

    # ── Immediate Priority ────────────────────────────────────────────────────
    _section_header("⚡ Immediate Priority", "الأولوية الفورية")

    st.session_state["mom_priority"] = st.text_area(
        "Priority / الأولوية",
        value=st.session_state.get("mom_priority", ""),
        height=80,
        placeholder="الأولوية الفورية: استلام ومراجعة المقترح التقني التفصيلي…",
        key="mom_prio_inp",
        label_visibility="collapsed",
    )

    st.success("✅ All sections saved automatically. Go to **Tab C → Generate Minutes** when ready.")


def _section_header(en: str, ar: str):
    st.markdown(
        f"<div style='background:{_GREEN};color:#fff;padding:6px 12px;"
        f"border-radius:6px;margin:16px 0 8px 0;font-size:13px;font-weight:700;'>"
        f"{en} &nbsp;<span style='color:#C9974A;font-size:12px;'>{ar}</span></div>",
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────────────────────────
# Step C — Generate & Preview
# ─────────────────────────────────────────────────────────────────────────────
def _render_step_c():
    st.markdown(
        "<p style='font-size:13px;color:#374151;margin-bottom:16px;'>"
        "Choose your output format, then preview or download the formatted document.</p>",
        unsafe_allow_html=True,
    )

    mode = st.radio(
        "Output language",
        ["Arabic Only — عربي فقط", "Bilingual — ثنائي اللغة (English + Arabic)"],
        horizontal=True,
        key="mom_output_mode",
    )
    bilingual = "Bilingual" in mode

    col_prev, col_dl = st.columns([1, 1])

    with col_prev:
        if st.button("👁 Preview Document", use_container_width=True, key="mom_preview_btn"):
            st.session_state["mom_show_preview"] = True

    with col_dl:
        html_bytes = _build_html(_collect_data(), bilingual=bilingual)
        fname = "MoM_Arabic.html" if not bilingual else "MoM_Bilingual.html"
        st.download_button(
            "⬇ Download HTML (Print-ready)",
            data=html_bytes,
            file_name=fname,
            mime="text/html",
            use_container_width=True,
            key="mom_dl_btn",
        )

    st.markdown(
        "<div style='background:#fffbeb;border:1px solid #fde68a;border-radius:6px;"
        "padding:8px 12px;font-size:12px;color:#92400e;margin:8px 0;'>"
        "💡 <b>To print:</b> Open the downloaded HTML file in your browser, then press "
        "<b>Ctrl+P</b> (or ⌘+P on Mac). Set margins to <b>None</b> and enable "
        "<b>Background graphics</b> for best results.</div>",
        unsafe_allow_html=True,
    )

    if st.session_state.get("mom_show_preview", False):
        st.markdown("---")
        st.markdown(
            f"<p style='font-size:12px;color:#6B7280;margin-bottom:4px;'>"
            f"Preview (scroll to see full document)</p>",
            unsafe_allow_html=True,
        )
        html_str = html_bytes.decode("utf-8")
        st.components.v1.html(html_str, height=900, scrolling=True)


# ─────────────────────────────────────────────────────────────────────────────
# Data collection helper
# ─────────────────────────────────────────────────────────────────────────────
def _collect_data() -> dict:
    return {
        "date":         st.session_state.get("mom_date",        date.today().strftime("%d %B %Y")),
        "ref":          st.session_state.get("mom_ref",         ""),
        "subject":      st.session_state.get("mom_subject",     ""),
        "objective":    st.session_state.get("mom_objective",   ""),
        "priority":     st.session_state.get("mom_priority",    ""),
        "prepared_by":  st.session_state.get("mom_prepared_by", "Minister Outreach Office, MISA"),
        "attendees":    st.session_state.get("mom_attendees",   []),
        "actions":      st.session_state.get("mom_actions",     []),
        "disc_points":  st.session_state.get("mom_disc_points", []),
        "next_steps":   st.session_state.get("mom_next_steps",  []),
        "en_notes":     st.session_state.get("mom_en_notes",    ""),
    }


# ─────────────────────────────────────────────────────────────────────────────
# HTML generation
# ─────────────────────────────────────────────────────────────────────────────
def _build_html(data: dict, bilingual: bool = False) -> bytes:
    arabic_page = _arabic_page(data)
    if bilingual:
        english_page = _english_page(data)
        body = english_page + '<div style="page-break-after:always;"></div>' + arabic_page
    else:
        body = arabic_page

    html = f"""<!DOCTYPE html>
<html lang="ar">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>محضر اجتماع — MISA</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: 'Arial', 'Tahoma', 'Segoe UI', sans-serif;
    font-size: 11pt;
    color: #1a1a1a;
    background: #f5f5f5;
  }}
  .page {{
    width: 210mm;
    min-height: 297mm;
    margin: 12mm auto;
    background: #fff;
    padding: 16mm 16mm 14mm 16mm;
    box-shadow: 0 2px 16px rgba(0,0,0,.12);
  }}
  /* ── Arabic page ── */
  .ar {{ direction: rtl; text-align: right; }}
  /* ── English page ── */
  .en {{ direction: ltr; text-align: left; }}

  /* ── Header ── */
  .mom-header {{
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    border-bottom: 3px solid {_GREEN};
    padding-bottom: 10px;
    margin-bottom: 14px;
  }}
  .mom-title-block {{ flex: 1; }}
  .mom-title {{
    font-size: 16pt;
    font-weight: 700;
    color: {_GREEN};
    line-height: 1.2;
  }}
  .mom-subtitle {{
    font-size: 10pt;
    color: {_GREEN};
    margin-top: 3px;
  }}
  .mom-meta {{
    text-align: left;
    font-size: 9pt;
  }}
  .ar .mom-meta {{ text-align: right; }}
  .confidential {{
    color: {_GOLD};
    font-weight: 700;
    font-size: 11pt;
    text-transform: uppercase;
    letter-spacing: 1px;
  }}
  .mom-date {{ color: {_GREEN}; font-weight: 600; margin-top: 3px; }}
  .mom-ref  {{ color: {_GREEN}; font-size: 9pt; margin-top: 2px; }}

  /* ── Subject ── */
  .subject-line {{
    background: #f0fdf4;
    border-right: 4px solid {_GOLD};
    padding: 7px 12px;
    font-size: 11pt;
    font-weight: 600;
    margin-bottom: 16px;
    color: {_GREEN};
  }}
  .en .subject-line {{
    border-right: none;
    border-left: 4px solid {_GOLD};
  }}

  /* ── Sections ── */
  .section {{ margin-bottom: 18px; }}
  .section-title {{
    font-size: 12pt;
    font-weight: 700;
    color: {_GREEN};
    border-bottom: 2px solid {_GOLD};
    padding-bottom: 4px;
    margin-bottom: 10px;
  }}
  .section-body {{
    font-size: 10.5pt;
    line-height: 1.6;
    color: #1f2937;
  }}

  /* ── Bullet lists ── */
  .bullet-list {{ padding-right: 20px; padding-left: 0; }}
  .en .bullet-list {{ padding-left: 20px; padding-right: 0; }}
  .bullet-list li {{
    margin-bottom: 5px;
    font-size: 10.5pt;
    color: #1f2937;
    line-height: 1.5;
  }}

  /* ── Tables ── */
  table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 9.5pt;
  }}
  th {{
    background: {_GREEN};
    color: #fff;
    font-weight: 700;
    padding: 7px 10px;
    text-align: right;
    border: 1px solid #d1fae5;
  }}
  .en th {{ text-align: left; }}
  td {{
    padding: 6px 10px;
    border: 1px solid #e5e7eb;
    vertical-align: top;
    color: #1f2937;
  }}
  tr:nth-child(even) td {{ background: #f9fafb; }}
  .cell-num {{
    width: 28px;
    text-align: center;
    font-weight: 700;
    color: {_GREEN};
  }}

  /* ── Priority box ── */
  .priority-box {{
    border: 1px solid {_GOLD};
    border-radius: 6px;
    padding: 10px 14px;
    background: #fffbeb;
    font-size: 10.5pt;
    color: #92400e;
    line-height: 1.5;
  }}

  /* ── Footer ── */
  .mom-footer {{
    border-top: 1px solid #d1d5db;
    padding-top: 8px;
    margin-top: 24px;
    display: flex;
    justify-content: space-between;
    font-size: 8.5pt;
    color: #6b7280;
  }}

  /* ── Print ── */
  @media print {{
    body {{ background: #fff; }}
    .page {{
      box-shadow: none;
      margin: 0;
      padding: 12mm 14mm 12mm 14mm;
      width: 100%;
    }}
  }}
</style>
</head>
<body>
{body}
</body>
</html>"""
    return html.encode("utf-8")


def _arabic_page(data: dict) -> str:
    att_rows = "".join(
        f"<tr><td>{_e(a['name'])}</td><td>{_e(a['role'])}</td></tr>"
        for a in data["attendees"] if a.get("name")
    )
    disc_items = "".join(
        f"<li>{_e(pt)}</li>" for pt in data["disc_points"] if pt.strip()
    )
    action_rows = "".join(
        f"<tr>"
        f"<td class='cell-num'>{_e(a['num'])}</td>"
        f"<td>{_e(a['item'])}</td>"
        f"<td>{_e(a['owner'])}</td>"
        f"<td>{_e(a['deliverable'])}</td>"
        f"<td>{_e(a['due'])}</td>"
        f"<td>{_e(a['measure'])}</td>"
        f"</tr>"
        for a in data["actions"] if a.get("item")
    )
    step_items = "".join(
        f"<li>{_e(s)}</li>" for s in data["next_steps"] if s.strip()
    )

    return f"""
<div class="page ar">
  <!-- Header -->
  <div class="mom-header">
    <div class="mom-meta">
      <div class="confidential">سري</div>
      <div class="mom-date">{_e(data['date'])}</div>
      <div class="mom-ref">{_e(data['ref'])}</div>
    </div>
    <div class="mom-title-block" style="text-align:right;">
      <div class="mom-title">محضر اجتماع (MoM)</div>
      <div class="mom-subtitle">وزارة الاستثمار — المملكة العربية السعودية</div>
    </div>
  </div>

  <!-- Subject -->
  <div class="subject-line">
    <span style="color:#6B7280;font-weight:400;">الموضوع:&nbsp;</span>{_e(data['subject'])}
  </div>

  <!-- Attendees -->
  <div class="section">
    <div class="section-title">الحضور</div>
    <table>
      <thead><tr><th>الاسم</th><th>المنصب / الدور</th></tr></thead>
      <tbody>{att_rows or "<tr><td colspan='2' style='color:#9ca3af;text-align:center;'>—</td></tr>"}</tbody>
    </table>
  </div>

  <!-- Objective -->
  <div class="section">
    <div class="section-title">هدف الاجتماع</div>
    <div class="section-body">{_e(data['objective']) or "<span style='color:#9ca3af;'>—</span>"}</div>
  </div>

  <!-- Discussion Points -->
  <div class="section">
    <div class="section-title">محاور النقاش الرئيسية</div>
    <ul class="bullet-list">{disc_items or "<li style='color:#9ca3af;'>—</li>"}</ul>
  </div>

  <!-- Action Items -->
  <div class="section">
    <div class="section-title">بنود العمل</div>
    <table>
      <thead>
        <tr>
          <th style="width:28px;">م</th>
          <th>بند العمل</th>
          <th>المسؤول</th>
          <th>المخرج</th>
          <th>تاريخ الاستحقاق</th>
          <th>مقياس النجاح</th>
        </tr>
      </thead>
      <tbody>{action_rows or "<tr><td colspan='6' style='color:#9ca3af;text-align:center;'>—</td></tr>"}</tbody>
    </table>
  </div>

  <!-- Next Steps -->
  <div class="section">
    <div class="section-title">ملخص الخطوات التالية</div>
    <ul class="bullet-list">{step_items or "<li style='color:#9ca3af;'>—</li>"}</ul>
  </div>

  <!-- Immediate Priority -->
  <div class="section">
    <div class="section-title">الأولوية الفورية</div>
    <div class="priority-box">{_e(data['priority']) or "<span style='color:#9ca3af;'>—</span>"}</div>
  </div>

  <!-- Footer -->
  <div class="mom-footer">
    <span>للاستخدام الداخلي فقط — وزارة الاستثمار</span>
    <span>أعده: {_e(data['prepared_by'])}</span>
  </div>
</div>"""


def _english_page(data: dict) -> str:
    en_notes = data.get("en_notes", "").strip()
    # Render the English notes as simple paragraphs / bullets
    lines_html = ""
    for line in en_notes.splitlines():
        stripped = line.strip()
        if not stripped:
            lines_html += "<br>"
        elif stripped.startswith(("• ", "- ", "* ")):
            lines_html += f"<li>{_e(stripped[2:])}</li>"
        else:
            lines_html += f"<p style='margin:4px 0;'>{_e(stripped)}</p>"

    att_rows = "".join(
        f"<tr><td>{_e(a['name'])}</td><td>{_e(a['role'])}</td></tr>"
        for a in data["attendees"] if a.get("name")
    )

    return f"""
<div class="page en">
  <!-- Header -->
  <div class="mom-header">
    <div class="mom-title-block">
      <div class="mom-title">Minutes of Meeting (MoM)</div>
      <div class="mom-subtitle">Ministry of Investment of Saudi Arabia (MISA)</div>
    </div>
    <div class="mom-meta" style="text-align:right;">
      <div class="confidential">CONFIDENTIAL</div>
      <div class="mom-date">{_e(data['date'])}</div>
      <div class="mom-ref">{_e(data['ref'])}</div>
    </div>
  </div>

  <!-- Subject -->
  <div class="subject-line">
    <span style="color:#6B7280;font-weight:400;">Subject:&nbsp;</span>{_e(data['subject'])}
  </div>

  <!-- Attendees -->
  <div class="section">
    <div class="section-title">Attendees</div>
    <table>
      <thead><tr><th style="text-align:left;">Name</th><th style="text-align:left;">Role</th></tr></thead>
      <tbody>{att_rows or "<tr><td colspan='2' style='color:#9ca3af;text-align:center;'>—</td></tr>"}</tbody>
    </table>
  </div>

  <!-- English Notes -->
  <div class="section">
    <div class="section-title">Meeting Notes</div>
    <div class="section-body" style="line-height:1.7;">{lines_html or "<span style='color:#9ca3af;'>No English notes provided.</span>"}</div>
  </div>

  <!-- Footer -->
  <div class="mom-footer">
    <span>Prepared by: {_e(data['prepared_by'])}</span>
    <span>For internal use only — Ministry of Investment of Saudi Arabia</span>
  </div>
</div>"""


def _e(val) -> str:
    """HTML-escape a value safely."""
    import html
    if not val:
        return ""
    return html.escape(str(val))
