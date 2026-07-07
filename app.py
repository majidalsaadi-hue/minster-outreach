
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
from modules.contacts      import render as render_contacts
from modules.deals         import render as render_deals
from modules.company_directory import render as render_company_directory
from modules.report_builder      import render as render_report_builder
from modules.evaluation          import render as render_evaluation
from modules.mom_builder         import render as render_mom_builder

_PPTX_IMPORT_ERROR = None
try:
    from exports.pptx_generator import generate_pptx, generate_pptx_company, generate_pptx_all_companies_dashboard
    _PPTX_AVAILABLE = True
except Exception as _e:
    _PPTX_AVAILABLE = False
    _PPTX_IMPORT_ERROR = str(_e)
    generate_pptx                       = None
    generate_pptx_company               = None
    generate_pptx_all_companies_dashboard = None

_PDF_IMPORT_ERROR = None
try:
    from exports.pdf_generator import generate_pdf
    _PDF_AVAILABLE = True
except Exception as _e:
    _PDF_AVAILABLE = False
    _PDF_IMPORT_ERROR = str(_e)
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

        # Global search
        dfs_for_search = st.session_state.get("dfs")
        if dfs_for_search is not None:
            gsq = st.text_input(
                "Global Search",
                placeholder="🔍 Search all data…",
                key="global_search_q",
                label_visibility="collapsed",
            )
            if gsq:
                st.session_state["_gsq_active"] = gsq.strip()
            elif "global_search_q" in st.session_state and not st.session_state.get("global_search_q", "").strip():
                st.session_state.pop("_gsq_active", None)
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

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
            T("nav_meetings"):          "meetings",
            T("nav_opportunities"):     "opportunities",
            T("nav_report_builder"):    "report_builder",
            T("nav_evaluation"):        "evaluation",
            T("nav_actions"):           "actions",
            T("nav_tasks"):             "tasks",
            "Contacts":                 "contacts",
            "Deals":                    "deals",
            "📄 MoM Builder":           "mom_builder",
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
              {"<div>📜 " + str(summary.get('deals',0)) + " deals</div>" if summary.get('deals',0) else ""}
              {"<div>👤 " + str(summary.get('contacts',0)) + " contacts</div>" if summary.get('contacts',0) else ""}
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

    # Preserve existing RM Tasks if the uploaded Excel has no Tasks sheet
    existing_dfs = st.session_state.get("dfs") or {}
    existing_tasks = existing_dfs.get("RM Tasks", pd.DataFrame())
    uploaded_tasks = dfs.get("RM Tasks", pd.DataFrame())
    if not existing_tasks.empty and uploaded_tasks.empty:
        dfs["RM Tasks"] = existing_tasks

    st.session_state["dfs"]              = dfs
    st.session_state["summary"]          = get_summary(dfs)
    st.session_state["last_upload_name"] = uploaded_file.name
    st.session_state["last_upload_time"] = date.today().strftime("%d %b %Y")
    save_session(dfs)
    st.success(f"✅ {T('success_upload')}: **{uploaded_file.name}**")


# ── Overdue action banner ─────────────────────────────────────────────────────
def _render_overdue_banner(dfs: dict):
    actions = dfs.get("Action Items", pd.DataFrame())
    if actions.empty or "Due Date" not in actions.columns or "Status" not in actions.columns:
        return

    today = date.today()
    pending = actions[~actions["Status"].str.lower().str.contains("complet|cancel", na=False)]
    if pending.empty:
        return

    dates = pd.to_datetime(pending["Due Date"], errors="coerce")
    overdue = pending[dates.notna() & (dates.dt.date < today)]
    if overdue.empty:
        return

    n = len(overdue)
    # Show top 3 most overdue
    top = overdue.copy()
    top["_days"] = (pd.to_datetime(today) - pd.to_datetime(top["Due Date"], errors="coerce")).dt.days
    top = top.nlargest(3, "_days")
    items_html = "".join(
        f'<span style="background:rgba(255,255,255,0.25);border-radius:4px;'
        f'padding:2px 8px;margin-right:6px;font-size:11px;">'
        f'{str(r.get("Company Name","?"))[:20]} · {str(r.get("Action Description","?"))[:35]}…'
        f' ({int(r["_days"])}d)</span>'
        for _, r in top.iterrows()
    )
    dismiss_key = f"_overdue_dismissed_{date.today().isoformat()}"
    if st.session_state.get(dismiss_key):
        return

    col_banner, col_x = st.columns([10, 1])
    with col_banner:
        st.markdown(
            f'<div style="background:#DC2626;color:#fff;border-radius:8px;'
            f'padding:8px 14px;margin-bottom:10px;display:flex;'
            f'align-items:center;flex-wrap:wrap;gap:6px;">'
            f'<span style="font-size:13px;font-weight:700;">🔴 {n} overdue action{"s" if n!=1 else ""}</span>'
            f'&nbsp;{items_html}'
            f'</div>',
            unsafe_allow_html=True,
        )
    with col_x:
        if st.button("✕", key=f"dismiss_overdue_{date.today().isoformat()}", help="Dismiss for today"):
            st.session_state[dismiss_key] = True
            st.rerun()


# ── Global search renderer ────────────────────────────────────────────────────
def _render_global_search(dfs: dict, query: str):
    query = query.strip()
    if not query:
        return

    st.markdown(
        f'<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:8px;'
        f'padding:10px 14px;margin-bottom:16px;">'
        f'<span style="font-size:13px;font-weight:600;color:#92400e;">🔍 Search results for: </span>'
        f'<span style="font-size:13px;color:#1c1917;">"{query}"</span>'
        f'&nbsp;<span style="font-size:11px;color:#9CA3AF;">(clear search box to dismiss)</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    q_low = query.lower()
    found_any = False

    _SECTIONS = [
        ("Investor Master",      "🏢 Investors",      ["Company Name", "Sector", "Country", "Relationship Manager", "Account Manager"]),
        ("Meeting Log",          "🤝 Meetings",        ["Company Name", "Meeting Type", "Meeting Objective", "Key Discussion Points"]),
        ("Action Items",         "📋 Actions",         ["Company Name", "Action Description", "Assigned To", "Status"]),
        ("Opportunity Pipeline", "🎯 Opportunities",  ["Company Name", "Opportunity Name", "Opportunity Stage"]),
        ("Deal Progress",        "📜 Deals",           ["Company Name", "Deal Name", "Deal Stage"]),
    ]

    for sheet, label, search_cols in _SECTIONS:
        df = dfs.get(sheet, pd.DataFrame())
        if df.empty:
            continue
        cols_present = [c for c in search_cols if c in df.columns]
        if not cols_present:
            continue

        mask = df[cols_present].apply(
            lambda col: col.fillna("").astype(str).str.lower().str.contains(q_low, regex=False)
        ).any(axis=1)
        hits = df[mask]
        if hits.empty:
            continue

        found_any = True
        st.markdown(f"**{label}** — {len(hits)} match{'es' if len(hits) != 1 else ''}")

        display_cols = cols_present[:4]
        st.dataframe(
            hits[display_cols].head(8).reset_index(drop=True),
            use_container_width=True,
            hide_index=True,
        )

    if not found_any:
        st.info(f'No results found for "{query}" across any sheet.')

    st.markdown("---")


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

    # ── Filters bar (company + date range) ────────────────────────────────
    _inv = dfs.get("Investor Master", pd.DataFrame())
    _act = dfs.get("Action Items", pd.DataFrame())
    _mtg = dfs.get("Meeting Log",   pd.DataFrame())
    _companies = sorted(_inv["Company Name"].dropna().unique().tolist()) if "Company Name" in _inv.columns else []

    fc1, fc2, fc3, fc4 = st.columns([2, 1.2, 1.2, 0.6])
    with fc1:
        selected_company = st.selectbox(
            "View company:",
            ["All Companies"] + _companies,
            key="dash_company_filter",
        )
    with fc2:
        dash_from = st.date_input("From", value=None, key="dash_date_from", label_visibility="visible")
    with fc3:
        dash_to   = st.date_input("To",   value=None, key="dash_date_to",   label_visibility="visible")
    with fc4:
        if st.button("Clear", key="dash_date_clear", help="Reset date range"):
            st.session_state["dash_date_from"] = None
            st.session_state["dash_date_to"]   = None
            st.rerun()

    # Action summary strip (uses unfiltered actions for overall counts)
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
        _range_note = ""
        if dash_from or dash_to:
            _range_note = (
                f' &nbsp;·&nbsp; <span style="color:#7C3AED;">📅 '
                f'{"From " + dash_from.strftime("%d %b") if dash_from else ""}'
                f'{" – " if dash_from and dash_to else ""}'
                f'{"To " + dash_to.strftime("%d %b") if dash_to else ""}'
                f'</span>'
            )
        st.markdown(
            f"<div style='font-size:13px;color:#555;margin-bottom:4px;'>"
            f"<b>{_n_tot}</b> actions total — "
            f"<span style='color:#1B5C3F'><b>{_n_done}</b> completed</span> · "
            f"<span style='color:#C9974A'><b>{_n_prog}</b> in progress</span> · "
            f"<b>{_n_due}</b> due{_next_due}{_range_note}"
            f"</div>",
            unsafe_allow_html=True,
        )

    # Build filtered view (company + optional date range)
    def _filter_by_date(df: pd.DataFrame, date_col: str) -> pd.DataFrame:
        if df.empty or date_col not in df.columns:
            return df
        d_col = pd.to_datetime(df[date_col], errors="coerce")
        mask = pd.Series(True, index=df.index)
        if dash_from:
            mask &= d_col.dt.date >= dash_from
        if dash_to:
            mask &= d_col.dt.date <= dash_to
        return df[mask].reset_index(drop=True)

    if selected_company == "All Companies":
        view_dfs = dict(dfs)
    else:
        view_dfs = {}
        for sheet, df in dfs.items():
            if isinstance(df, pd.DataFrame) and "Company Name" in df.columns:
                view_dfs[sheet] = df[df["Company Name"] == selected_company].reset_index(drop=True)
            else:
                view_dfs[sheet] = df

    if dash_from or dash_to:
        view_dfs["Meeting Log"]  = _filter_by_date(view_dfs.get("Meeting Log",  pd.DataFrame()), "Meeting Date")
        view_dfs["Action Items"] = _filter_by_date(view_dfs.get("Action Items", pd.DataFrame()), "Due Date")

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
            err_detail = f" — {_PPTX_IMPORT_ERROR}" if _PPTX_IMPORT_ERROR else ""
            st.warning(f"PPTX generator failed to load{err_detail}")
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
            err_detail = f" — {_PPTX_IMPORT_ERROR}" if _PPTX_IMPORT_ERROR else ""
            st.warning(f"PPTX generator failed to load{err_detail}")
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

    # ── All-companies dashboard ─────────────────────────────────────────────
    st.markdown("---")
    st.markdown("#### 📊 All Companies Dashboard")
    st.markdown("""<div style="font-size:13px;color:#374151">
        Strategic infographic deck — cover page with sector/country/progress charts,
        then one page per company.
    </div>""", unsafe_allow_html=True)
    all_lang_pptx = st.selectbox("Language", ["English", "Arabic"], key="pptx_all_lang")
    if generate_pptx_all_companies_dashboard is None:
        err_detail = f" — {_PPTX_IMPORT_ERROR}" if _PPTX_IMPORT_ERROR else ""
        st.warning(f"PPTX generator failed to load{err_detail}")
    elif st.button("📥 Generate All Companies Dashboard", use_container_width=True, disabled=(dfs is None)):
        with st.spinner("Building all-companies dashboard…"):
            all_lang_code = "ar" if all_lang_pptx == "Arabic" else "en"
            all_pptx = generate_pptx_all_companies_dashboard(dfs or {}, lang=all_lang_code)
        all_filename = f"MoI_AllCompanies_Dashboard_{date.today().strftime('%Y-%m-%d')}.pptx"
        st.download_button(
            "⬇️ Download All Companies Dashboard (.pptx)",
            data=all_pptx,
            file_name=all_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            key="dl_all_companies_pptx",
        )

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

    # Global search banner — shown on top of any page when search is active
    gsq = st.session_state.get("_gsq_active", "").strip()
    if gsq and dfs is not None:
        _render_global_search(dfs, gsq)

    # Overdue action banner
    if dfs is not None and page != "dashboard":
        _render_overdue_banner(dfs)

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
        render_evaluation(dfs, lang())


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

    elif page == "contacts":
        if dfs is None:
            _require_data()
        else:
            render_contacts(dfs, lang())

    elif page == "deals":
        if dfs is None:
            _require_data()
        else:
            render_deals(dfs, lang())

    elif page == "mom_builder":
        render_mom_builder(dfs or {}, lang())

    elif page == "export":
        render_export()


def _require_data():
    st.info(f"⬆️ {T('no_data')}")


if __name__ == "__main__":
    main()
