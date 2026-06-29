
# Ministry of Investment — Investor Relations CRM
# Single-user web application for the Relationship Manager.
# Launch: streamlit run app.py

import io
import json
import os
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

# ── Page config (must be first Streamlit call) ──────────────────────────────
st.set_page_config(
    page_title="MoI Investor CRM",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Load custom CSS ─────────────────────────────────────────────────────────
_CSS_PATH = Path(__file__).parent / "assets" / "styles.css"
if _CSS_PATH.exists():
    st.markdown(f"<style>{_CSS_PATH.read_text()}</style>", unsafe_allow_html=True)

# ── Internal imports ─────────────────────────────────────────────────────────
from config.translations import t
from config.settings     import MISA_GREEN, MISA_GOLD

from modules.data_loader   import load_excel, validate_schema, get_summary
from modules.persistence   import save_session, load_session
from modules.dashboard     import (
    render_action_advisor,
    render_kpi_cards,
    render_investment_flow,
    render_alerts,
    render_active_opportunities_panel,
    render_progress_chart,
)
from modules.investors     import render as render_investors
from modules.outreach      import render as render_outreach
from modules.meetings      import render as render_meetings
from modules.opportunities import render as render_opportunities
from modules.actions       import render as render_actions, render_action_summary_widget
from modules.tasks         import render as render_tasks
from modules.report_builder      import render as render_report_builder
from modules.evaluation          import render as render_evaluation

try:
    from exports.pptx_generator import generate_pptx, generate_pptx_company
    _PPTX_AVAILABLE = True
except Exception:
    _PPTX_AVAILABLE = False
    generate_pptx         = None
    generate_pptx_company = None

try:
    from exports.pdf_generator import generate_pdf
    _PDF_AVAILABLE = True
except Exception:
    _PDF_AVAILABLE = False
    generate_pdf = None

from exports.excel_exporter import export_status_excel, generate_template


# ── Session state initialisation ─────────────────────────────────────────────
def _init_state():
    defaults = {
        "lang":     "en",
        "dfs":      None,
        "last_upload_name": None,
        "last_upload_time": None,
        "summary":  None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()

# ── Auto-save on every rerun (runs before main so st.rerun() can't skip it) ──
if st.session_state.get("dfs") is not None:
    try:
        save_session(st.session_state["dfs"])
    except Exception:
        pass

# ── Auto-load saved session on first run ─────────────────────────────────────
if st.session_state["dfs"] is None:
    _saved = load_session()
    if _saved is not None:
        st.session_state["dfs"]     = _saved
        st.session_state["summary"] = get_summary(_saved)
        st.session_state.setdefault("last_upload_time", "Auto-loaded")


# ── Language helpers ─────────────────────────────────────────────────────────
def lang() -> str:
    return st.session_state["lang"]


def T(key: str) -> str:
    return t(key, lang())


# ── Sidebar ───────────────────────────────────────────────────────────────────
def render_sidebar():
    with st.sidebar:
        # Logo area
        st.markdown("""
        <div style="text-align:center;padding:16px 8px 8px 8px;">
          <div style="font-size:28px;margin-bottom:4px;">🏛️</div>
          <div style="color:#C9974A;font-weight:700;font-size:14px;">
            Ministry of Investment
          </div>
          <div style="color:rgba(255,255,255,0.6);font-size:11px;">
            وزارة الاستثمار
          </div>
        </div>
        <hr style="border-color:rgba(255,255,255,0.2);margin:8px 0 16px 0;">
        """, unsafe_allow_html=True)

        # Language toggle
        current_lang = lang()
        toggle_label = T("language_toggle")
        if st.button(f"🌐 {toggle_label}", use_container_width=True, key="lang_toggle"):
            st.session_state["lang"] = "ar" if current_lang == "en" else "en"
            st.rerun()

        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

        # Navigation
        nav_options = {
            T("nav_dashboard"):         "dashboard",
            T("nav_outreach"):          "outreach",
            T("nav_investors"):         "investors",
            T("nav_opportunities"):     "opportunities",
            T("nav_report_builder"):    "report_builder",
            T("nav_evaluation"):        "evaluation",
            T("nav_actions"):           "actions",
            T("nav_tasks"):             "tasks",
            T("nav_export"):            "export",
        }
        selected_label = st.radio(
            "Navigation",
            list(nav_options.keys()),
            label_visibility="collapsed",
        )
        page = nav_options[selected_label]

        st.markdown("<hr style='border-color:rgba(255,255,255,0.2);margin:16px 0;'>", unsafe_allow_html=True)

        # Upload section
        st.markdown(f"<div style='color:rgba(255,255,255,0.8);font-size:12px;margin-bottom:6px;'>📂 {T('upload_excel')}</div>",
                    unsafe_allow_html=True)
        uploaded = st.file_uploader(
            T("upload_excel"),
            type=["xlsx"],
            label_visibility="collapsed",
            key="excel_uploader",
        )
        if uploaded is not None:
            _handle_upload(uploaded)

        # Clear saved data button
        if st.session_state["dfs"] is not None:
            if st.button("🗑 Clear & Re-upload", use_container_width=True, key="clear_btn",
                         help="Delete saved session and reset — then upload your Excel again"):
                from modules.persistence import SAVE_PATH
                try:
                    if SAVE_PATH.exists():
                        SAVE_PATH.unlink()
                except Exception:
                    pass
                st.session_state["dfs"]     = None
                st.session_state["summary"] = None
                st.rerun()

        # Data status
        if st.session_state["dfs"] is not None:
            summary = st.session_state["summary"] or {}
            st.markdown(f"""
            <div style="background:rgba(255,255,255,0.1);border-radius:6px;
                        padding:10px;margin-top:8px;font-size:11px;">
              <div style="color:#C9974A;font-weight:700;margin-bottom:4px;">
                ✓ {T('data_loaded')}
              </div>
              <div>📊 {summary.get('investors',0)} {T('investors_count')}</div>
              <div>🤝 {summary.get('meetings',0)} {T('meetings_count')}</div>
              <div>🎯 {summary.get('opportunities',0)} {T('opportunities_count')}</div>
              <div>✅ {summary.get('actions',0)} {T('actions_count')}</div>
              <div style="margin-top:4px;color:rgba(255,255,255,0.5);">
                {T('last_updated')}: {st.session_state.get('last_upload_time','—')}
              </div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div style="background:rgba(255,255,255,0.08);border-radius:6px;
                        padding:10px;margin-top:8px;font-size:11px;
                        color:rgba(255,255,255,0.5);">
              No data loaded.<br>Upload an Excel file above.
            </div>
            """, unsafe_allow_html=True)

        # ── Persistence status ────────────────────────────────────────────────
        if st.session_state.get("dfs") is not None:
            save_status = st.session_state.get("save_status", "")
            save_time   = st.session_state.get("save_time", "")

            if save_status == "ok":
                st.markdown(f"""
                <div style="background:rgba(16,185,129,0.15);border:1px solid rgba(16,185,129,0.35);
                            border-radius:6px;padding:8px 10px;margin-top:8px;">
                  <div style="color:#6EE7B7;font-size:11px;font-weight:600;">
                    Data saved automatically
                  </div>
                  <div style="color:rgba(255,255,255,0.45);font-size:10px;margin-top:2px;">
                    Last saved at {save_time} · restores on next launch
                  </div>
                </div>""", unsafe_allow_html=True)
            elif save_status.startswith("error"):
                err_msg = save_status.replace("error: ", "")
                st.markdown(f"""
                <div style="background:rgba(220,38,38,0.15);border:1px solid rgba(220,38,38,0.35);
                            border-radius:6px;padding:8px 10px;margin-top:8px;">
                  <div style="color:#FCA5A5;font-size:11px;font-weight:600;">Save failed</div>
                  <div style="color:rgba(255,255,255,0.45);font-size:10px;">{err_msg[:60]}</div>
                </div>""", unsafe_allow_html=True)
            else:
                st.markdown("""
                <div style="font-size:10px;color:rgba(255,255,255,0.35);
                            margin-top:6px;text-align:center;">Saving…</div>
                """, unsafe_allow_html=True)

        return page


def _handle_upload(uploaded_file):
    import io as _io

    raw_bytes = uploaded_file.read()

    try:
        xl_peek   = pd.ExcelFile(_io.BytesIO(raw_bytes))
        sheet_names = xl_peek.sheet_names
    except Exception as e:
        st.error(f"Cannot read Excel file: {e}")
        return

    _OUTREACH_SHEETS = {"الشركات الأجنبية", "الشركات المحلية"}
    _CRM_SHEETS      = {"Investor Master", "Action Items", "Opportunity Pipeline"}
    has_outreach = bool(_OUTREACH_SHEETS & set(sheet_names))
    has_crm      = bool(_CRM_SHEETS & set(sheet_names))

    # ── Pure outreach Excel (Arabic sheets only) ──────────────────────────
    if has_outreach and not has_crm:
        from modules.outreach import load_outreach_excel, _merge_outreach
        outreach_df = load_outreach_excel(_io.BytesIO(raw_bytes))
        if outreach_df is None or outreach_df.empty:
            st.warning("No company data found in the Outreach Excel.")
            return
        if st.session_state["dfs"] is None:
            st.session_state["dfs"] = {}
        existing = st.session_state["dfs"].get("Outreach Tracker", pd.DataFrame())
        merged   = _merge_outreach(existing, outreach_df) if not existing.empty else outreach_df
        st.session_state["dfs"]["Outreach Tracker"] = merged
        st.session_state["summary"]          = get_summary(st.session_state["dfs"])
        st.session_state["last_upload_name"] = uploaded_file.name
        st.session_state["last_upload_time"] = date.today().strftime("%d %b %Y")
        save_session(st.session_state["dfs"])
        n_f = len(merged[merged["Company Type"] == "Foreign"]) if "Company Type" in merged.columns else 0
        n_l = len(merged[merged["Company Type"] == "Local"])   if "Company Type" in merged.columns else 0
        st.success(f"✅ Outreach Excel loaded: **{len(merged)} companies** ({n_f} foreign, {n_l} local)")
        return

    # ── Standard CRM Excel (may also contain outreach sheets) ────────────
    with st.spinner(T("loading")):
        dfs = load_excel(_io.BytesIO(raw_bytes))
    if dfs is None:
        st.error(T("error_upload"))
        return

    if has_outreach:
        try:
            from modules.outreach import load_outreach_excel, _merge_outreach
            outreach_df = load_outreach_excel(_io.BytesIO(raw_bytes))
            if outreach_df is not None and not outreach_df.empty:
                existing = dfs.get("Outreach Tracker", pd.DataFrame())
                dfs["Outreach Tracker"] = _merge_outreach(existing, outreach_df) if not existing.empty else outreach_df
        except Exception:
            pass

    warnings = validate_schema(dfs)
    if warnings:
        st.warning("Schema warnings:\n" + "\n".join(f"• {w}" for w in warnings))

    st.session_state["dfs"]              = dfs
    st.session_state["summary"]          = get_summary(dfs)
    st.session_state["last_upload_name"] = uploaded_file.name
    st.session_state["last_upload_time"] = date.today().strftime("%d %b %Y")
    save_session(dfs)
    st.success(f"✅ {T('success_upload')}: **{uploaded_file.name}**")


# ── Page: Dashboard ───────────────────────────────────────────────────────────
def render_dashboard():
    dfs = st.session_state["dfs"]

    # Header
    st.markdown(f"""
    <div class="misa-header">
      <div>
        <div style="font-size:20px;font-weight:700;color:white;">
          {'وزارة الاستثمار — إدارة علاقات المستثمرين' if lang()=='ar' else 'Ministry of Investment — Investor Relations CRM'}
        </div>
        <div class="subtitle">
          {'لوحة التحكم التنفيذية' if lang()=='ar' else 'Executive Dashboard — Relationship Manager View'}
        </div>
      </div>
      <div style="text-align:right;color:rgba(255,255,255,0.7);font-size:12px;">
        {date.today().strftime('%d %B %Y')}
      </div>
    </div>
    """, unsafe_allow_html=True)

    if dfs is None:
        _render_no_data_welcome()
        return

    # ── Company filter + action summary bar ───────────────────────────────
    _inv = dfs.get("Investor Master", pd.DataFrame())
    _act = dfs.get("Action Items", pd.DataFrame())
    _companies = sorted(_inv["Company Name"].dropna().unique().tolist()) if "Company Name" in _inv.columns else []

    filter_col, summary_col = st.columns([2, 3])
    with filter_col:
        selected_company = st.selectbox(
            "View company:",
            ["All Companies"] + _companies,
            key="dash_company_filter",
        )
    with summary_col:
        if not _act.empty and "Status" in _act.columns:
            _act_f = _act if selected_company == "All Companies" else (
                _act[_act["Company Name"] == selected_company] if "Company Name" in _act.columns else _act
            )
            _n_tot  = len(_act_f)
            _n_done = int(_act_f["Status"].str.lower().str.contains("complet", na=False).sum())
            _n_prog = int(_act_f["Status"].str.lower().str.contains("in progress|inprogress", na=False).sum())
            _n_due  = max(_n_tot - _n_done - _n_prog, 0)
            _next_due = ""
            if "Due Date" in _act_f.columns:
                _pending = _act_f[~_act_f["Status"].str.lower().str.contains("complet", na=False)]
                _dates   = pd.to_datetime(_pending["Due Date"], errors="coerce").dropna()
                _future  = _dates[_dates >= pd.Timestamp(date.today())]
                if not _future.empty:
                    _next_due = " · Next due " + _future.min().strftime("%d %b").lstrip("0")
            st.markdown(
                f"<div style='margin-top:28px;font-size:13px;color:#555;'>"
                f"<b>{_n_tot}</b> actions total — "
                f"<span style='color:#1B5C3F'><b>{_n_done}</b> completed</span> · "
                f"<span style='color:#C9974A'><b>{_n_prog}</b> in progress</span> · "
                f"<b>{_n_due}</b> due{_next_due}"
                f"</div>",
                unsafe_allow_html=True,
            )

    # Build filtered view
    if selected_company == "All Companies":
        view_dfs = dfs
    else:
        view_dfs = {}
        for sheet, df in dfs.items():
            if isinstance(df, pd.DataFrame) and "Company Name" in df.columns:
                view_dfs[sheet] = df[df["Company Name"] == selected_company].reset_index(drop=True)
            else:
                view_dfs[sheet] = df

    # ── Today's Briefing — action advisor + progress chart ───────────────
    adv_col, chart_col = st.columns([3, 1])
    with adv_col:
        render_action_advisor(view_dfs, lang())
    with chart_col:
        st.markdown("**Action Progress**")
        render_progress_chart(view_dfs, lang())

    # ── Row 1 + Row 2: KPI strips ──────────────────────────────────────────
    render_kpi_cards(view_dfs, lang())

    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    st.markdown("---")

    # ── Active Investment Opportunities (minister view) ────────────────────
    render_active_opportunities_panel(view_dfs, lang())

    st.markdown("---")

    # ── Strategic Alerts ──────────────────────────────────────────────────
    st.markdown("#### Strategic Alerts")
    render_alerts(view_dfs, lang())

    st.markdown("---")

    # ── Action Items — attention panel ───────────────────────────────────
    render_action_summary_widget(view_dfs, lang())

    st.markdown("---")

    # ── Investment Flow ───────────────────────────────────────────────────
    st.markdown(f"#### {T('dash_investment_flow')}")
    render_investment_flow(view_dfs, lang())


def _render_no_data_welcome():
    st.markdown("""
    <div style="background:white;border-radius:12px;padding:40px;
                text-align:center;margin-top:40px;
                border:2px dashed #C9974A;">
      <div style="font-size:48px;margin-bottom:16px;">📊</div>
      <h2 style="color:#1B5C3F;margin-bottom:8px;">Welcome to the Investor Relations CRM</h2>
      <p style="color:#6B6B6B;font-size:14px;max-width:500px;margin:0 auto;">
        Upload your Excel data file using the sidebar to begin.
        The tool accepts both the enhanced 5-sheet format and the
        legacy per-investor Action Item Tracker format.
      </p>
      <div style="margin-top:24px;padding:16px;background:#F7F7F2;
                  border-radius:8px;display:inline-block;text-align:left;">
        <div style="font-weight:600;color:#1B5C3F;margin-bottom:8px;">
          Expected sheets (new format):
        </div>
        <div style="font-size:13px;color:#4A4A4A;">
          1. Investor Master &nbsp;|&nbsp;
          2. Meeting Log &nbsp;|&nbsp;
          3. Opportunity Pipeline &nbsp;|&nbsp;
          4. Action Items &nbsp;|&nbsp;
          5. RM Tasks
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)


# ── Page: Export & Reports ────────────────────────────────────────────────────
def render_export():
    st.markdown(f"### {T('nav_export')}")

    dfs = st.session_state["dfs"]

    col1, col2, col3, col4 = st.columns(4)

    # ── PowerPoint ───────────────────────────────────────────────────────────
    with col1:
        st.markdown(f"""
        <div style="background:white;border-radius:10px;padding:20px;
                    border:1px solid #E0E0E0;min-height:200px;">
          <div style="font-size:32px;margin-bottom:8px;">📊</div>
          <div style="font-weight:700;color:#1B5C3F;margin-bottom:4px;">
            {T('export_pptx')}
          </div>
          <div style="font-size:12px;color:#6B6B6B;margin-bottom:12px;">
            Executive Account Dashboard + Opportunities Flow.<br>
            Named: MoI_Investor_Status_YYYY-MM-DD.pptx
          </div>
        </div>
        """, unsafe_allow_html=True)
        report_lang_pptx = st.selectbox("Language", ["English", "Arabic"], key="pptx_lang")
        if not _PPTX_AVAILABLE:
            st.warning("PPTX export requires python-pptx. Run: python -m pip install python-pptx")
        elif st.button(T("generate_report") + " (PPTX)", use_container_width=True, disabled=(dfs is None)):
            with st.spinner(T("generating")):
                pptx_lang = "ar" if report_lang_pptx == "Arabic" else "en"
                pptx_bytes = generate_pptx(dfs or {}, lang=pptx_lang)
            filename = f"MoI_Investor_Status_{date.today().strftime('%Y-%m-%d')}.pptx"
            st.download_button(
                T("download") + " PPTX",
                data=pptx_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )

    # ── PDF ──────────────────────────────────────────────────────────────────
    with col2:
        st.markdown(f"""
        <div style="background:white;border-radius:10px;padding:20px;
                    border:1px solid #E0E0E0;min-height:200px;">
          <div style="font-size:32px;margin-bottom:8px;">📄</div>
          <div style="font-weight:700;color:#1B5C3F;margin-bottom:4px;">
            {T('export_pdf')}
          </div>
          <div style="font-size:12px;color:#6B6B6B;margin-bottom:12px;">
            KPIs, portfolio, opportunities, overdue actions.<br>
            Clean single-document for Minister briefing.
          </div>
        </div>
        """, unsafe_allow_html=True)
        report_lang_pdf = st.selectbox("Language", ["English", "Arabic"], key="pdf_lang")
        if not _PDF_AVAILABLE:
            st.warning("PDF export requires additional packages. Run: python -m pip install reportlab")
        elif st.button(T("generate_report") + " (PDF)", use_container_width=True, disabled=(dfs is None)):
            with st.spinner(T("generating")):
                pdf_lang  = "ar" if report_lang_pdf == "Arabic" else "en"
                pdf_bytes = generate_pdf(dfs or {}, lang=pdf_lang)
            filename = f"MoI_Executive_Summary_{date.today().strftime('%Y-%m-%d')}.pdf"
            st.download_button(
                T("download") + " PDF",
                data=pdf_bytes,
                file_name=filename,
                mime="application/pdf",
                use_container_width=True,
            )

    # ── Excel Status Export ───────────────────────────────────────────────────
    with col3:
        st.markdown(f"""
        <div style="background:white;border-radius:10px;padding:20px;
                    border:1px solid #E0E0E0;min-height:200px;">
          <div style="font-size:32px;margin-bottom:8px;">📋</div>
          <div style="font-weight:700;color:#1B5C3F;margin-bottom:4px;">
            {T('export_excel')}
          </div>
          <div style="font-size:12px;color:#6B6B6B;margin-bottom:12px;">
            Full data export: investors, meetings, opportunities,
            actions, tasks — formatted for archiving.
          </div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("")
        if st.button(T("export_excel"), use_container_width=True, disabled=(dfs is None)):
            with st.spinner(T("generating")):
                xl_bytes = export_status_excel(dfs or {})
            filename = f"MoI_Status_Export_{date.today().strftime('%Y-%m-%d')}.xlsx"
            st.download_button(
                T("download") + " Excel",
                data=xl_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    # ── Excel Template ────────────────────────────────────────────────────────
    with col4:
        st.markdown(f"""
        <div style="background:white;border-radius:10px;padding:20px;
                    border:1px solid #E0E0E0;min-height:200px;">
          <div style="font-size:32px;margin-bottom:8px;">📝</div>
          <div style="font-weight:700;color:#1B5C3F;margin-bottom:4px;">
            {T('export_template')}
          </div>
          <div style="font-size:12px;color:#6B6B6B;margin-bottom:12px;">
            Enhanced 5-sheet Excel template with dropdowns,
            validation, and pre-filled sample data.
          </div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("")
        if st.button(T("export_template"), use_container_width=True):
            with st.spinner(T("generating")):
                tmpl_bytes = generate_template()
            filename = f"MoI_CRM_Template_{date.today().strftime('%Y-%m-%d')}.xlsx"
            st.download_button(
                T("download") + " Template",
                data=tmpl_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    # ── Per-company PowerPoint ────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("#### 🏢 Company-Specific PowerPoint")

    if dfs is not None:
        investors_df = dfs.get("Investor Master", pd.DataFrame())
        companies = sorted(investors_df["Company Name"].dropna().unique().tolist()) if not investors_df.empty and "Company Name" in investors_df.columns else []
    else:
        companies = []

    if companies:
        cc1, cc2 = st.columns([2, 1])
        selected_co = cc1.selectbox("Select Company", companies, key="pptx_company_select")
        co_lang     = cc2.selectbox("Language", ["English", "Arabic"], key="pptx_company_lang")
        if not _PPTX_AVAILABLE:
            st.warning("PPTX export requires python-pptx. Run: python -m pip install python-pptx")
        elif st.button("Generate Company PPTX", use_container_width=False, disabled=(dfs is None)):
            with st.spinner(T("generating")):
                co_lang_code = "ar" if co_lang == "Arabic" else "en"
                co_pptx = generate_pptx_company(dfs, selected_co, lang=co_lang_code)
            co_filename = f"MoI_{selected_co.replace(' ','_')}_{date.today().strftime('%Y-%m-%d')}.pptx"
            st.download_button(
                f"Download — {selected_co}",
                data=co_pptx,
                file_name=co_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
    else:
        st.info("Upload data first to generate company reports.")

    # ── Export notes ──────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("""
    **Export notes:**
    - All exports use the data currently loaded in the session.
    - Upload an updated Excel file (sidebar) before generating reports to ensure data is current.
    - PPTX and PDF are available in English or Arabic. Excel exports are data-only.
    - The Excel Template includes all 5 sheets with dropdowns, validation, and sample data from real investors.
    """)


# ── Main router ───────────────────────────────────────────────────────────────
def main():
    page = render_sidebar()
    dfs  = st.session_state["dfs"]

    if page == "dashboard":
        render_dashboard()

    elif page == "outreach":
        render_outreach(dfs, lang())

    elif page == "investors":
        if dfs is None:
            _require_data()
        else:
            render_investors(dfs, lang())

    elif page == "meetings":
        if dfs is None:
            _require_data()
        else:
            render_meetings(dfs, lang())

    elif page == "opportunities":
        if dfs is None:
            _require_data()
        else:
            render_opportunities(dfs, lang())

    elif page == "report_builder":
        if dfs is None:
            _require_data()
        else:
            render_report_builder(dfs, lang())

    elif page == "evaluation":
        render_evaluation()


    elif page == "actions":
        if dfs is None:
            _require_data()
        else:
            render_actions(dfs, lang())

    elif page == "tasks":
        if dfs is None:
            _require_data()
        else:
            render_tasks(dfs, lang())

    elif page == "export":
        render_export()


def _require_data():
    st.info(f"⬆️ {T('no_data')}")


if __name__ == "__main__":
    main()
