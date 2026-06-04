
# Opportunity Pipeline — strategic pipeline board with stage kanban.

import hashlib
import pandas as pd
import plotly.express as px
import streamlit as st
from datetime import date

from config.settings import (
    SECTORS, OPPORTUNITY_TYPES, OPPORTUNITY_SOURCES, OPPORTUNITY_STAGES,
    CONFIDENCE_LEVELS, MISA_GREEN, MISA_GOLD,
)
from config.translations import t
from modules.persistence import save_session

_GREEN  = "#1B5C3F"
_GOLD   = "#C9974A"
_RED    = "#DC2626"
_AMBER  = "#D97706"
_BLUE   = "#1D4ED8"

OPP_STATUSES = ["Active", "Under Review", "Blocked", "Converted to Deal", "Dropped"]

_STAGES = [
    "Exploration",
    "Due Diligence",
    "Active Negotiation",
    "Committed",
    "Post-Investment",
]
_STAGE_COLORS = {
    "Exploration":        "#6B7280",
    "Due Diligence":      "#1D4ED8",
    "Active Negotiation": "#D97706",
    "Committed":          "#059669",
    "Post-Investment":    _GREEN,
}
_CONF_COLORS = {
    "High":        ("#D1FAE5", "#065F46"),
    "Medium":      ("#FEF3C7", "#92400E"),
    "Low":         ("#FEE2E2", "#991B1B"),
    "Speculative": ("#F3F4F6", "#374151"),
}


def render(dfs: dict, lang: str):
    opps      = dfs.get("Opportunity Pipeline", pd.DataFrame())
    investors = dfs.get("Investor Master",      pd.DataFrame())

    st.markdown(
        f"<h2 style='color:{_GREEN};margin-bottom:2px;'>Opportunity Pipeline</h2>"
        f"<p style='color:#6b7280;font-size:13px;margin-top:0;'>"
        f"Track, manage and advance every investment opportunity</p>",
        unsafe_allow_html=True,
    )

    # ── Add opportunity form ──────────────────────────────────────────────────
    with st.expander("+ Add New Opportunity", expanded=False):
        _add_form(dfs, investors, lang)

    if opps.empty:
        st.markdown("""
        <div style="background:#f0fdf4;border:2px dashed #86efac;border-radius:12px;
                    padding:40px;text-align:center;margin-top:16px;">
          <div style="font-size:40px;margin-bottom:12px;">🎯</div>
          <h3 style="color:#1B5C3F;margin-bottom:8px;">No opportunities yet</h3>
          <p style="color:#6b7280;font-size:13px;">
            Add an opportunity above, or upload your Excel tracker — the CRM will
            extract opportunities from the Action Items sheets automatically.
          </p>
        </div>""", unsafe_allow_html=True)
        return

    # Normalise stage column — map unknown stages to Exploration
    if "Opportunity Stage" in opps.columns:
        opps["Opportunity Stage"] = opps["Opportunity Stage"].apply(
            lambda s: s if str(s) in _STAGES else "Exploration"
        )

    # ── Filters ───────────────────────────────────────────────────────────────
    fc1, fc2, fc3, fc4 = st.columns([2.5, 1.5, 1.5, 1.5])
    companies  = sorted(opps["Company Name"].dropna().unique()) if "Company Name" in opps.columns else []
    with fc1:
        q = st.text_input("Search", placeholder="Search company or opportunity…",
                          key="opp_search", label_visibility="collapsed")
    with fc2:
        sel_stage = st.selectbox("Stage", ["All stages"] + _STAGES,
                                 key="opp_stage", label_visibility="collapsed")
    with fc3:
        conf_opts = ["All confidence"] + CONFIDENCE_LEVELS
        sel_conf = st.selectbox("Confidence", conf_opts,
                                key="opp_conf", label_visibility="collapsed")
    with fc4:
        sel_status = st.selectbox("Status", ["All statuses"] + OPP_STATUSES,
                                  key="opp_status", label_visibility="collapsed")

    view = opps.copy()
    if q:
        mask = (
            view.get("Company Name", pd.Series(dtype=str)).fillna("").str.contains(q, case=False, na=False)
            | view.get("Opportunity Name", pd.Series(dtype=str)).fillna("").str.contains(q, case=False, na=False)
        )
        view = view[mask]
    if sel_stage  != "All stages":    view = view[view.get("Opportunity Stage",   pd.Series()) == sel_stage]
    if sel_conf   != "All confidence": view = view[view.get("Confidence Level",    pd.Series()) == sel_conf]
    if sel_status != "All statuses":  view = view[view.get("Opportunity Status",  pd.Series()) == sel_status]

    if view.empty:
        st.warning("No opportunities match the current filters.")
        return

    # ── KPI strip ─────────────────────────────────────────────────────────────
    total_val   = _sum_col(view, "Est. Value (SAR)")
    active_cnt  = int((view.get("Opportunity Status", pd.Series()) == "Active").sum()) if "Opportunity Status" in view.columns else len(view)
    committed   = int((view.get("Opportunity Stage", pd.Series()) == "Committed").sum()) if "Opportunity Stage" in view.columns else 0
    high_conf   = int((view.get("Confidence Level",  pd.Series()) == "High").sum())      if "Confidence Level"  in view.columns else 0

    k1, k2, k3, k4 = st.columns(4)
    _kpi(k1, str(len(view)),         "Total Opportunities",  _GREEN)
    _kpi(k2, _fmt_sar(total_val),    "Pipeline Value",       _GOLD)
    _kpi(k3, str(committed),         "Committed",            "#059669")
    _kpi(k4, str(high_conf),         "High Confidence",      _BLUE)

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # ── Stage pipeline bar ────────────────────────────────────────────────────
    _render_stage_bar(view)

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    # ── Pipeline board ────────────────────────────────────────────────────────
    _render_board(view)


# ── Stage progress bar ────────────────────────────────────────────────────────

def _render_stage_bar(df: pd.DataFrame):
    if "Opportunity Stage" not in df.columns:
        return
    counts = df["Opportunity Stage"].value_counts().to_dict()
    total  = sum(counts.values()) or 1

    segs_html = ""
    for stage in _STAGES:
        cnt = counts.get(stage, 0)
        pct = cnt / total * 100
        if pct < 1:
            continue
        color = _STAGE_COLORS.get(stage, "#6B7280")
        segs_html += f"""
        <div style="flex:{pct};background:{color};height:100%;
                    display:flex;align-items:center;justify-content:center;
                    font-size:10px;color:#fff;font-weight:600;white-space:nowrap;
                    overflow:hidden;padding:0 4px;min-width:0;"
             title="{stage}: {cnt}">
          {cnt}
        </div>"""

    legend_html = ""
    for stage in _STAGES:
        cnt = counts.get(stage, 0)
        color = _STAGE_COLORS.get(stage, "#6B7280")
        legend_html += f"""
        <span style="display:inline-flex;align-items:center;gap:4px;margin-right:14px;
                     font-size:11px;color:#374151;">
          <span style="width:10px;height:10px;border-radius:2px;
                       background:{color};display:inline-block;flex-shrink:0;"></span>
          {stage} ({cnt})
        </span>"""

    st.markdown(f"""
    <div style="border-radius:6px;overflow:hidden;height:28px;display:flex;margin-bottom:8px;">
      {segs_html}
    </div>
    <div style="margin-bottom:4px;">{legend_html}</div>
    """, unsafe_allow_html=True)


# ── Kanban board ──────────────────────────────────────────────────────────────

def _render_board(df: pd.DataFrame):
    stage_col = "Opportunity Stage" if "Opportunity Stage" in df.columns else None
    if stage_col is None:
        _render_list(df)
        return

    active_stages = [s for s in _STAGES if s in df[stage_col].values]
    if not active_stages:
        _render_list(df)
        return

    cols = st.columns(len(active_stages))
    for col_idx, stage in enumerate(active_stages):
        stage_df = df[df[stage_col] == stage]
        color    = _STAGE_COLORS.get(stage, "#6B7280")
        with cols[col_idx]:
            val = _sum_col(stage_df, "Est. Value (SAR)")
            st.markdown(f"""
            <div style="background:{color}18;border:1px solid {color}40;
                        border-radius:8px;padding:8px 10px;margin-bottom:8px;">
              <div style="color:{color};font-size:11px;font-weight:700;
                          letter-spacing:.4px;text-transform:uppercase;">{stage}</div>
              <div style="display:flex;justify-content:space-between;align-items:center;
                          margin-top:3px;">
                <span style="color:#374151;font-size:12px;font-weight:600;">
                  {len(stage_df)} opp{'s' if len(stage_df)!=1 else ''}
                </span>
                <span style="color:{color};font-size:11px;font-weight:600;">
                  {_fmt_sar(val)}
                </span>
              </div>
            </div>""", unsafe_allow_html=True)

            for _, row in stage_df.iterrows():
                _opp_card(row)


def _render_list(df: pd.DataFrame):
    for _, row in df.iterrows():
        _opp_card(row)


# ── Opportunity card ──────────────────────────────────────────────────────────

def _opp_card(row):
    company   = str(row.get("Company Name",       "?"))
    opp_name  = str(row.get("Opportunity Name",   "—"))
    stage     = str(row.get("Opportunity Stage",  "—"))
    conf      = str(row.get("Confidence Level",   "—"))
    status    = str(row.get("Opportunity Status", "Active"))
    val       = row.get("Est. Value (SAR)", None)
    am        = str(row.get("Assigned AM", "") or row.get("Account Manager", "") or "—")
    opp_type  = str(row.get("Opportunity Type",   ""))
    sector    = str(row.get("Sector", ""))

    color        = _STAGE_COLORS.get(stage, "#6B7280")
    conf_bg, conf_fg = _CONF_COLORS.get(conf, ("#F3F4F6", "#374151"))
    val_str      = _fmt_sar(float(val)) if val and str(val) not in ("nan", "None", "") else ""
    avatar       = _initials(company)
    avatar_color = _company_color(company)

    status_badge = ""
    if status.lower() in ("blocked", "dropped"):
        status_badge = f'<span style="background:#FEE2E2;color:#991B1B;padding:1px 6px;border-radius:4px;font-size:9px;font-weight:600;margin-left:4px;">{status.upper()}</span>'
    elif status.lower() == "converted to deal":
        status_badge = f'<span style="background:#D1FAE5;color:#065F46;padding:1px 6px;border-radius:4px;font-size:9px;font-weight:600;margin-left:4px;">DEAL</span>'

    sector_tag = f'<span style="background:#f8fafc;color:#6b7280;padding:1px 5px;border-radius:4px;font-size:9px;">{sector}</span>' if sector and sector != "nan" else ""
    type_tag   = f'<span style="background:#ede9fe;color:#7c3aed;padding:1px 5px;border-radius:4px;font-size:9px;">{opp_type}</span>' if opp_type and opp_type not in ("nan", "Opportunity") else ""

    st.markdown(f"""
    <div style="border:1px solid #e5e7eb;border-left:3px solid {color};
                border-radius:8px;padding:10px;margin-bottom:6px;background:#fff;">
      <div style="display:flex;align-items:center;gap:7px;margin-bottom:6px;">
        <div style="width:28px;height:28px;border-radius:6px;background:{avatar_color};
                    display:flex;align-items:center;justify-content:center;
                    font-size:10px;font-weight:700;color:#fff;flex-shrink:0;">{avatar}</div>
        <div style="flex:1;min-width:0;">
          <div style="font-size:11px;font-weight:700;color:{_GREEN};
                      white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">{company}</div>
        </div>
        {f'<span style="font-size:12px;font-weight:700;color:{_GOLD};">{val_str}</span>' if val_str else ""}
      </div>
      <div style="font-size:12px;font-weight:600;color:#111827;margin-bottom:5px;
                  line-height:1.3;">{opp_name}{status_badge}</div>
      <div style="display:flex;flex-wrap:wrap;gap:4px;margin-bottom:5px;">
        <span style="background:{conf_bg};color:{conf_fg};padding:1px 6px;
                     border-radius:4px;font-size:9px;font-weight:600;">{conf}</span>
        {sector_tag}
        {type_tag}
      </div>
      {f'<div style="font-size:10px;color:#9ca3af;">AM: {am}</div>' if am != "—" else ""}
    </div>""", unsafe_allow_html=True)


# ── Add opportunity form ──────────────────────────────────────────────────────

def _add_form(dfs: dict, investors: pd.DataFrame, lang: str):
    company_options = [""] + (
        sorted(investors["Company Name"].dropna().unique().tolist())
        if not investors.empty and "Company Name" in investors.columns else []
    )
    with st.form("add_opp_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        company  = c1.selectbox("Company", company_options)
        opp_name = c2.text_input("Opportunity Name")
        c3, c4 = st.columns(2)
        opp_type = c3.selectbox("Type", OPPORTUNITY_TYPES)
        source   = c4.selectbox("Source", OPPORTUNITY_SOURCES)
        c5, c6 = st.columns(2)
        sector   = c5.selectbox("Sector", [""] + SECTORS)
        stage    = c6.selectbox("Stage", _STAGES)
        c7, c8 = st.columns(2)
        conf     = c7.selectbox("Confidence Level", CONFIDENCE_LEVELS)
        status   = c8.selectbox("Status", OPP_STATUSES)
        c9, c10 = st.columns(2)
        est_val  = c9.number_input("Est. Value (SAR)", min_value=0.0, step=1_000_000.0)
        target_d = c10.date_input("Target Closure", value=None)
        c11, c12 = st.columns(2)
        assigned = c11.text_input("Assigned AM")
        escalate = c12.selectbox("Escalation Required", ["No", "Yes"])
        blockers = st.text_area("Blockers / Notes")

        if st.form_submit_button("Add Opportunity", use_container_width=True):
            if not company or not opp_name:
                st.warning("Company and opportunity name are required.")
                return
            opps   = dfs.get("Opportunity Pipeline", pd.DataFrame())
            inv_id = _get_investor_id(investors, company)
            new_id = _next_id(opps, "Opportunity ID", "OPP")
            new_row = {
                "Opportunity ID":     new_id,
                "Investor ID":        inv_id,
                "Company Name":       company,
                "Opportunity Name":   opp_name,
                "Opportunity Type":   opp_type,
                "Opportunity Source": source,
                "Sector":             sector,
                "Opportunity Stage":  stage,
                "Confidence Level":   conf,
                "Opportunity Status": status,
                "Est. Value (SAR)":   est_val if est_val > 0 else None,
                "Assigned AM":        assigned,
                "Start Date":         date.today(),
                "Target Closure Date":target_d,
                "Blockers":           blockers,
                "Escalation Required":escalate,
                "Last Updated":       date.today(),
            }
            dfs["Opportunity Pipeline"] = pd.concat(
                [opps, pd.DataFrame([new_row])], ignore_index=True
            )
            save_session(dfs)
            st.success(f"Opportunity {new_id} added for {company}")
            st.rerun()


# ── Helpers ───────────────────────────────────────────────────────────────────

def _kpi(col, value: str, label: str, color: str):
    col.markdown(f"""
    <div style="background:{color};padding:14px 10px;border-radius:8px;
                text-align:center;">
      <div style="color:rgba(255,255,255,0.75);font-size:9px;font-weight:600;
                  letter-spacing:.5px;text-transform:uppercase;margin-bottom:4px;">{label}</div>
      <div style="color:#fff;font-size:22px;font-weight:700;line-height:1.1;">{value}</div>
    </div>""", unsafe_allow_html=True)


def _initials(name: str) -> str:
    parts = name.split()
    if len(parts) >= 2:
        return (parts[0][0] + parts[-1][0]).upper()
    return name[:2].upper() if name else "?"


def _company_color(name: str) -> str:
    h = int(hashlib.md5(name.encode()).hexdigest()[:6], 16)
    r = max(40, min(int((h >> 16) & 0xFF), 160))
    g = max(40, min(int((h >> 8)  & 0xFF), 160))
    b = max(40, min(int(h & 0xFF),          160))
    return f"#{r:02X}{g:02X}{b:02X}"


def _sum_col(df: pd.DataFrame, col: str) -> float:
    if df.empty or col not in df.columns:
        return 0.0
    return pd.to_numeric(df[col], errors="coerce").fillna(0).sum()


def _fmt_sar(val: float) -> str:
    if not val or (isinstance(val, float) and pd.isna(val)):
        return "—"
    if val >= 1e9:
        return f"SAR {val/1e9:.1f}B"
    if val >= 1e6:
        return f"SAR {val/1e6:.0f}M"
    if val > 0:
        return f"SAR {val:,.0f}"
    return "—"


def _get_investor_id(investors: pd.DataFrame, company: str) -> str:
    if investors.empty or "Company Name" not in investors.columns:
        return ""
    m = investors[investors["Company Name"] == company]
    return str(m.iloc[0].get("Investor ID", "")) if not m.empty else ""


def _next_id(df: pd.DataFrame, id_col: str, prefix: str) -> str:
    if df.empty or id_col not in df.columns:
        return f"{prefix}-001"
    nums = []
    for v in df[id_col].dropna():
        try:
            nums.append(int(str(v).split("-")[-1]))
        except ValueError:
            pass
    return f"{prefix}-{(max(nums)+1 if nums else 1):03d}"
