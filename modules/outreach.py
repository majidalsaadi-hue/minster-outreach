
# Outreach Tracker — strategic company pipeline through 4-step ministerial engagement workflow.

import io
import re
from datetime import date, datetime

import numpy as np
import pandas as pd
import streamlit as st

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from modules.persistence import save_session

_GREEN  = "#1B5C3F"
_GOLD   = "#C9974A"
_RED    = "#DC2626"
_AMBER  = "#D97706"
_BLUE   = "#1D4ED8"
_GRAY   = "#6B7280"

_MATURITY_LEVELS = ["L0", "L1", "L2", "L3", "L4", "L5"]

_MATURITY_COLORS = {
    "L0": "#9CA3AF", "L1": "#60A5FA", "L2": "#34D399",
    "L3": "#FBBF24", "L4": "#F97316", "L5": "#1B5C3F",
}

_MATURITY_DESCRIPTIONS = {
    "L0": "Idea",
    "L1": "Studied Concept (Invest Saudi published)",
    "L2": "Investment Opportunity (investor engaged, LOI/NDA)",
    "L3": "Deal – Negotiation (term sheet)",
    "L4": "Structured Deal (all signed, licensed)",
    "L5": "Closed (operational, capex started)",
}

_PRIORITY_COLORS = {"A": _RED, "B": _AMBER, "C": _BLUE}

_OPP_PRIORITY_MAP = {
    "عالية جداً": "Very High",
    "عالية":      "High",
    "متوسطة":     "Medium",
    "منخفضة":     "Low",
}

_OPP_PRIORITY_REVERSE = {v: k for k, v in _OPP_PRIORITY_MAP.items()}

_COL_MAP = {
    0:  "Seq",
    1:  "Company Name",
    2:  "Sector",
    3:  "Country",
    4:  "Company Size",
    5:  "Classification",
    6:  "Score M1",
    7:  "Score M2",
    8:  "Score M3",
    9:  "Score M4",
    10: "Score M5",
    11: "Score M6",
    12: "Total Score",
    13: "Priority",
    14: "Proposed Channel",
    15: "Work Track",
    16: "Batch",
    17: "RM",
    18: "AM",
    19: "Company Contact",
    20: "Phone",
    21: "Email",
    22: "Investment Opportunity",
    23: "Next Step",
    24: "Assessment Notes",
    25: "Maturity Level",
    26: "Maturity Description",
    27: "Opportunity Priority",
    28: "Previous Minister Meeting",
}

_ARABIC_COL_HEADERS = [
    "#", "اسم الشركة", "القطاع", "الدولة/المدينة", "الحجم",
    "نوع التصنيف", "م١", "م٢", "م٣", "م٤", "م٥", "م٦",
    "الدرجة الكلية ٪", "الأولوية", "القناة المقترحة", "مسار العمل",
    "الدفعة", "مدير العلاقة من مكتب معاليه", "مدير الحساب من MISA",
    "النظير من الشركة FOC", "هاتف التواصل", "البريد الإلكتروني",
    "الفرصة الاستثمارية", "الخطوة التالية", "ملاحظات التقييم",
    "مستوى النضوج L0→L5", "وصف النضوج والملاحظة",
    "أولوية الفرصة", "اجتماع سابق مع معاليه",
    # Workflow tracking added by CRM
    "هل عرض AM؟", "تاريخ العرض",
    "اجتماع مع معاليه (CRM)", "تاريخ الاجتماع",
]

_OUTPUT_COLS = [
    "Outreach ID", "Company Type", "Company Name", "Sector", "Country",
    "Company Size", "Classification",
    "Score M1", "Score M2", "Score M3", "Score M4", "Score M5", "Score M6",
    "Total Score", "Priority", "Proposed Channel", "Work Track", "Batch",
    "RM", "AM", "Company Contact", "Phone", "Email",
    "Investment Opportunity", "Next Step", "Assessment Notes",
    "Maturity Level", "Maturity Description", "Opportunity Priority",
    "Previous Minister Meeting",
    "AM Presented", "AM Presented Date", "Minister Engaged", "Minister Engaged Date",
    "Last Updated",
]


# ── Public API ─────────────────────────────────────────────────────────────────

def load_outreach_excel(uploaded_file) -> pd.DataFrame:
    xf = pd.ExcelFile(uploaded_file, engine="openpyxl")
    frames = []

    sheet_types = {
        "الشركات الأجنبية": "Foreign",
        "الشركات المحلية":  "Local",
    }

    for sheet_name, company_type in sheet_types.items():
        if sheet_name not in xf.sheet_names:
            continue
        raw = xf.parse(sheet_name, header=None)

        rows = []
        for idx in range(6, len(raw)):
            row = raw.iloc[idx]
            col0 = row.iloc[0] if len(row) > 0 else None
            col1 = row.iloc[1] if len(row) > 1 else None

            if col0 is None or col1 is None:
                continue

            col0_str = str(col0).strip()
            col1_str = str(col1).strip()

            if col1_str in ("", "nan"):
                continue
            if col0_str.startswith("►") or col0_str.startswith("⚠"):
                continue

            try:
                int(float(col0_str))
            except (ValueError, TypeError):
                continue

            record = {"Company Type": company_type}
            for col_idx, col_name in _COL_MAP.items():
                val = row.iloc[col_idx] if col_idx < len(row) else ""
                record[col_name] = _clean_cell(val)

            record["Maturity Level"]        = _normalize_maturity(record.get("Maturity Level", ""))
            record["Opportunity Priority"]  = _normalize_opp_priority(record.get("Opportunity Priority", ""))
            record["Previous Minister Meeting"] = _normalize_minister_meeting(record.get("Previous Minister Meeting", ""))

            rows.append(record)

        if rows:
            frames.append(pd.DataFrame(rows))

    if not frames:
        return _empty_outreach_df()

    combined = pd.concat(frames, ignore_index=True)
    combined["AM Presented"]       = "No"
    combined["AM Presented Date"]  = ""
    combined["Minister Engaged"]   = "No"
    combined["Minister Engaged Date"] = ""
    combined["Last Updated"]       = ""

    combined["Outreach ID"] = [f"OUT-{i+1:03d}" for i in range(len(combined))]

    for col in _OUTPUT_COLS:
        if col not in combined.columns:
            combined[col] = ""

    return combined[_OUTPUT_COLS].reset_index(drop=True)


def export_outreach_excel(dfs: dict) -> bytes:
    df = dfs.get("Outreach Tracker", pd.DataFrame())
    wb = openpyxl.Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    for company_type, sheet_name in [
        ("Foreign", "الشركات الأجنبية"),
        ("Local",   "الشركات المحلية"),
    ]:
        subset = df[df["Company Type"] == company_type].reset_index(drop=True) \
                 if not df.empty and "Company Type" in df.columns else pd.DataFrame()
        ws = wb.create_sheet(sheet_name)
        _write_outreach_sheet(ws, subset, company_type)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def render(dfs: dict, lang: str):
    tracker = dfs.get("Outreach Tracker", pd.DataFrame()) if dfs else pd.DataFrame()

    hcol, bcol = st.columns([5, 3])
    hcol.markdown(
        f"<h2 style='color:{_GREEN};margin-bottom:2px;'>Outreach Tracker</h2>",
        unsafe_allow_html=True,
    )

    with bcol:
        b1, b2, b3 = st.columns(3)
        imp_file = b1.file_uploader(
            "Import Outreach Excel",
            type=["xlsx"],
            label_visibility="collapsed",
            key="outreach_import_uploader",
            help="Import an Outreach Excel file (Arabic format)",
        )
        export_clicked = b2.button(
            "📥 Export Outreach Excel",
            use_container_width=True,
            key="outreach_export_btn",
            disabled=(tracker.empty),
        )
        sync_clicked = b3.button(
            "🔗 Sync to CRM",
            use_container_width=True,
            key="outreach_sync_btn",
            disabled=(tracker.empty or dfs is None),
            help="Copy RM, AM, contact info and investment opportunities into Investor Master and Opportunity Pipeline",
        )

    if imp_file is not None:
        _handle_import(imp_file, dfs)
        tracker = dfs.get("Outreach Tracker", pd.DataFrame())

    if export_clicked and not tracker.empty:
        xl = export_outreach_excel(dfs)
        fname = f"Outreach_{date.today().strftime('%Y-%m-%d')}.xlsx"
        st.download_button(
            "⬇ Download Outreach Excel",
            data=xl,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="outreach_dl",
        )

    if sync_clicked and not tracker.empty:
        n_inv, n_opp, n_act = _sync_to_crm(dfs)
        save_session(dfs)
        st.success(
            f"✅ Sync complete — {n_inv} investor records updated, "
            f"{n_opp} opportunities created, {n_act} action items added from Next Steps."
        )

    if tracker.empty:
        _render_import_prompt()
        return

    _render_funnel(tracker)
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    filtered = _render_filters(tracker)
    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

    _render_editor(filtered, tracker, dfs)


# ── Import handling ────────────────────────────────────────────────────────────

def _handle_import(uploaded_file, dfs: dict):
    with st.spinner("Parsing Outreach Excel…"):
        try:
            new_df = load_outreach_excel(uploaded_file)
        except Exception as exc:
            st.error(f"Import failed: {exc}")
            return

    if new_df.empty:
        st.warning("No data rows found in the uploaded file. Check the sheet names and row structure.")
        return

    existing = dfs.get("Outreach Tracker", pd.DataFrame())
    if existing.empty:
        merged = new_df
    else:
        merged = _merge_outreach(existing, new_df)

    if dfs is None:
        st.error("No active session — upload a CRM Excel file first.")
        return

    dfs["Outreach Tracker"] = merged
    save_session(dfs)
    st.success(f"✅ Imported {len(merged)} companies ({len(new_df[new_df['Company Type']=='Foreign'])} foreign, {len(new_df[new_df['Company Type']=='Local'])} local)")
    st.rerun()


def _merge_outreach(existing: pd.DataFrame, new_df: pd.DataFrame) -> pd.DataFrame:
    workflow_cols = ["AM Presented", "AM Presented Date", "Minister Engaged", "Minister Engaged Date"]
    preserve = existing.set_index("Company Name")[workflow_cols] \
               if "Company Name" in existing.columns else pd.DataFrame()

    result = new_df.copy()
    if not preserve.empty:
        for col in workflow_cols:
            if col in preserve.columns:
                result[col] = result["Company Name"].map(
                    preserve[col]
                ).fillna(result[col] if col in result.columns else "")

    existing_names = set(existing["Company Name"].dropna().tolist()) \
                     if "Company Name" in existing.columns else set()
    new_names      = set(new_df["Company Name"].dropna().tolist()) \
                     if "Company Name" in new_df.columns else set()
    truly_new_mask = ~new_df["Company Name"].isin(existing_names) \
                     if "Company Name" in new_df.columns else pd.Series([True]*len(new_df))

    return result.reset_index(drop=True)


# ── Workflow funnel ────────────────────────────────────────────────────────────

def _render_funnel(df: pd.DataFrame):
    n_total = len(df)

    has_opp = df["Investment Opportunity"].fillna("").str.strip().ne("").sum() \
              if "Investment Opportunity" in df.columns else 0

    has_am_rm = (
        df["AM"].fillna("").str.strip().ne("") &
        df["RM"].fillna("").str.strip().ne("")
    ).sum() if ("AM" in df.columns and "RM" in df.columns) else 0

    am_presented = (
        df["AM Presented"].fillna("").str.strip().str.lower() == "yes"
    ).sum() if "AM Presented" in df.columns else 0

    min_engaged = (
        df["Minister Engaged"].fillna("").str.strip().str.lower() == "yes"
    ).sum() if "Minister Engaged" in df.columns else 0

    steps = [
        ("1  Has Opportunity", has_opp,    _GREEN, "Investment opportunity defined"),
        ("2  AM/RM Assigned",  has_am_rm,  _BLUE,  "Both AM and RM assigned"),
        ("3  AM Presented",    am_presented, _AMBER, "AM presented to company"),
        ("4  Minister Engaged",min_engaged, _GOLD,  "Minister has engaged"),
    ]

    cols = st.columns(4)
    for col, (label, count, color, desc) in zip(cols, steps):
        pct = f"{int(count/n_total*100)}%" if n_total > 0 else "0%"
        col.markdown(
            f'<div style="background:#fff;border:1px solid #E5E7EB;border-top:4px solid {color};'
            f'border-radius:8px;padding:12px 16px;text-align:center;">'
            f'<div style="color:{color};font-size:28px;font-weight:700;">{count}</div>'
            f'<div style="color:#374151;font-size:10px;font-weight:700;margin:2px 0;">{label}</div>'
            f'<div style="color:#9CA3AF;font-size:9px;">{desc}</div>'
            f'<div style="color:{color};font-size:11px;font-weight:600;margin-top:4px;">{pct} of {n_total}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )


# ── Filters ────────────────────────────────────────────────────────────────────

def _render_filters(df: pd.DataFrame) -> pd.DataFrame:
    fc1, fc2, fc3, fc4, fc5 = st.columns([1.5, 1.2, 1.5, 1.8, 2.5])

    type_opts = ["All"] + sorted(df["Company Type"].dropna().unique().tolist()) \
                if "Company Type" in df.columns else ["All"]
    f_type = fc1.selectbox("Company Type", type_opts,
                           label_visibility="collapsed", key="out_type",
                           placeholder="All Types")

    pri_opts = ["All"] + sorted([
        p for p in df["Priority"].dropna().unique().tolist() if p
    ]) if "Priority" in df.columns else ["All"]
    f_pri = fc2.selectbox("Priority", pri_opts,
                          label_visibility="collapsed", key="out_pri",
                          placeholder="All Priorities")

    mat_opts = ["All"] + _MATURITY_LEVELS
    f_mat = fc3.selectbox("Maturity Level", mat_opts,
                          label_visibility="collapsed", key="out_mat",
                          placeholder="All Levels")

    gap_opts = [
        "All",
        "Missing Opportunity",
        "Missing AM or RM",
        "AM Not Presented",
        "Minister Not Engaged",
    ]
    f_gap = fc4.selectbox("Workflow Gap", gap_opts,
                          label_visibility="collapsed", key="out_gap",
                          placeholder="All Gaps")

    q = fc5.text_input("Search", placeholder="Search company, sector, country…",
                       label_visibility="collapsed", key="out_q")

    filtered = df.copy()

    if f_type != "All" and "Company Type" in filtered.columns:
        filtered = filtered[filtered["Company Type"] == f_type]

    if f_pri != "All" and "Priority" in filtered.columns:
        filtered = filtered[filtered["Priority"] == f_pri]

    if f_mat != "All" and "Maturity Level" in filtered.columns:
        filtered = filtered[filtered["Maturity Level"] == f_mat]

    if f_gap != "All":
        if f_gap == "Missing Opportunity":
            filtered = filtered[filtered.get("Investment Opportunity", pd.Series(dtype=str)).fillna("").str.strip() == ""]
        elif f_gap == "Missing AM or RM":
            no_am = filtered.get("AM", pd.Series(dtype=str)).fillna("").str.strip() == ""
            no_rm = filtered.get("RM", pd.Series(dtype=str)).fillna("").str.strip() == ""
            filtered = filtered[no_am | no_rm]
        elif f_gap == "AM Not Presented":
            filtered = filtered[
                filtered.get("AM Presented", pd.Series(dtype=str)).fillna("").str.strip().str.lower() != "yes"
            ]
        elif f_gap == "Minister Not Engaged":
            filtered = filtered[
                filtered.get("Minister Engaged", pd.Series(dtype=str)).fillna("").str.strip().str.lower() != "yes"
            ]

    if q:
        mask = (
            filtered.get("Company Name", pd.Series(dtype=str)).fillna("").str.contains(q, case=False, na=False)
            | filtered.get("Sector", pd.Series(dtype=str)).fillna("").str.contains(q, case=False, na=False)
            | filtered.get("Country", pd.Series(dtype=str)).fillna("").str.contains(q, case=False, na=False)
        )
        filtered = filtered[mask]

    return filtered


# ── Data editor ────────────────────────────────────────────────────────────────

def _render_editor(filtered: pd.DataFrame, full_tracker: pd.DataFrame, dfs: dict):
    display_cols = [c for c in [
        "Outreach ID", "Company Type", "Company Name", "Priority",
        "Sector", "Country", "Total Score",
        "RM", "AM", "Investment Opportunity", "Next Step",
        "Maturity Level",
        "AM Presented", "AM Presented Date",
        "Minister Engaged", "Minister Engaged Date",
    ] if c in filtered.columns]

    display = filtered[display_cols].copy().reset_index(drop=True)

    col_cfg = {
        "Outreach ID":           st.column_config.TextColumn("ID", width="small",  disabled=True),
        "Company Type":          st.column_config.TextColumn("Type", width="small", disabled=True),
        "Company Name":          st.column_config.TextColumn("Company", disabled=True),
        "Priority":              st.column_config.TextColumn("Pri.", width="small", disabled=True),
        "Sector":                st.column_config.TextColumn("Sector", disabled=True),
        "Country":               st.column_config.TextColumn("Country", disabled=True),
        "Total Score":           st.column_config.NumberColumn("Score", format="%.1f%%", disabled=True),
        "RM":                    st.column_config.TextColumn("RM"),
        "AM":                    st.column_config.TextColumn("AM"),
        "Investment Opportunity":st.column_config.TextColumn("Opportunity", width="large"),
        "Next Step":             st.column_config.TextColumn("Next Step"),
        "Maturity Level":        st.column_config.SelectboxColumn(
                                     "Maturity", options=_MATURITY_LEVELS, width="small"),
        "AM Presented":          st.column_config.SelectboxColumn(
                                     "AM Presented", options=["Yes", "No"], width="small"),
        "AM Presented Date":     st.column_config.DateColumn("AM Presented Date"),
        "Minister Engaged":      st.column_config.SelectboxColumn(
                                     "Min. Engaged", options=["Yes", "No"], width="small"),
        "Minister Engaged Date": st.column_config.DateColumn("Min. Engaged Date"),
    }

    st.markdown(
        '<div style="font-size:11px;color:#9CA3AF;margin-bottom:4px;">'
        '✏️ Click any editable cell to update. Press <strong>Save Changes</strong> to persist.'
        '</div>',
        unsafe_allow_html=True,
    )

    edited = st.data_editor(
        display,
        key="outreach_editor",
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config=col_cfg,
    )

    sa_col, _ = st.columns([1.5, 5])
    if sa_col.button("💾 Save Changes", type="primary", use_container_width=True, key="out_save"):
        n_saved = _merge_edits_back(dfs, edited, display, filtered)
        if n_saved:
            st.success(f"✅ Saved {n_saved} change{'s' if n_saved!=1 else ''}")
            st.rerun()
        else:
            st.info("No changes detected.")

    st.caption(f"{len(display)} companies shown")


def _merge_edits_back(dfs: dict, edited: pd.DataFrame, original: pd.DataFrame,
                      filtered: pd.DataFrame) -> int:
    editable_cols = [
        "RM", "AM", "Investment Opportunity", "Next Step",
        "Maturity Level", "AM Presented", "AM Presented Date",
        "Minister Engaged", "Minister Engaged Date",
    ]
    check_cols = [c for c in editable_cols if c in edited.columns and c in original.columns]
    if not check_cols:
        return 0

    full = dfs.get("Outreach Tracker", pd.DataFrame()).copy()
    if full.empty or "Outreach ID" not in full.columns:
        return 0

    n_changed = 0
    orig_reset = original.reset_index(drop=True)
    edit_reset = edited.reset_index(drop=True)

    for i in range(min(len(orig_reset), len(edit_reset))):
        orig_row = orig_reset.iloc[i]
        edit_row = edit_reset.iloc[i]
        out_id   = orig_row.get("Outreach ID")
        if not out_id:
            continue
        changed = any(
            str(orig_row.get(c, "")) != str(edit_row.get(c, ""))
            for c in check_cols
        )
        if not changed:
            continue

        mask = full["Outreach ID"] == out_id
        for c in check_cols:
            if c not in full.columns:
                continue
            val = edit_row.get(c)
            if c.endswith("Date"):
                if val is not None and str(val) not in ("", "nan", "NaT", "None"):
                    try:
                        ts = pd.to_datetime(val, errors="coerce")
                        full.loc[mask, c] = ts.date() if pd.notna(ts) else ""
                    except Exception:
                        full.loc[mask, c] = ""
                else:
                    full.loc[mask, c] = ""
            else:
                full.loc[mask, c] = val

        full.loc[mask, "Last Updated"] = date.today().isoformat()
        n_changed += 1

    if n_changed:
        dfs["Outreach Tracker"] = full
        save_session(dfs)
    return n_changed


# ── Sync to CRM ────────────────────────────────────────────────────────────────

def _sync_to_crm(dfs: dict):
    tracker = dfs.get("Outreach Tracker", pd.DataFrame())
    if tracker.empty:
        return 0, 0, 0

    investors = dfs.get("Investor Master",      pd.DataFrame()).copy()
    opps      = dfs.get("Opportunity Pipeline", pd.DataFrame()).copy()
    actions   = dfs.get("Action Items",         pd.DataFrame()).copy()

    n_inv = n_opp = n_act = 0

    for _, row in tracker.iterrows():
        company = str(row.get("Company Name", "")).strip()
        if not company:
            continue

        am      = str(row.get("AM",              "") or "").strip()
        rm      = str(row.get("RM",              "") or "").strip()
        phone   = str(row.get("Phone",           "") or "").strip()
        email   = str(row.get("Email",           "") or "").strip()
        contact = str(row.get("Company Contact", "") or "").strip()
        sector  = str(row.get("Sector",          "") or "").strip()
        country = str(row.get("Country",         "") or "").strip()
        size    = str(row.get("Company Size",    "") or "").strip()
        opp_text  = str(row.get("Investment Opportunity", "") or "").strip()
        next_step = str(row.get("Next Step",              "") or "").strip()
        mat_level = str(row.get("Maturity Level",         "") or "").strip()
        opp_pri   = str(row.get("Opportunity Priority",   "") or "").strip()

        # ── Investor Master ──────────────────────────────────────────────────
        if not investors.empty and "Company Name" in investors.columns:
            mask = investors["Company Name"] == company
            if mask.any():
                # Only overwrite if the new value is non-empty (never blank-out existing data)
                if am:      investors.loc[mask, "Account Manager"]       = am
                if rm:      investors.loc[mask, "Relationship Manager"]  = rm
                if phone:   investors.loc[mask, "Rep Phone"]             = phone
                if email:   investors.loc[mask, "Rep Email"]             = email
                if contact: investors.loc[mask, "Company Rep"]           = contact
                if size:    investors.loc[mask, "Company Size (Global)"] = size
            else:
                investors = pd.concat([investors, pd.DataFrame([{
                    "Investor ID":         _next_id(investors, "Investor ID", "INV"),
                    "Company Name":        company,
                    "Country":             country,
                    "Sector":             sector,
                    "Relationship Manager":rm,
                    "Account Manager":     am,
                    "Rep Phone":           phone,
                    "Rep Email":           email,
                    "Company Rep":         contact,
                    "Journey Stage":       "Qualification",
                    "Relationship Status": "Active",
                    "Investor Tier":       "Tier 2 — High Potential",
                }])], ignore_index=True)
            n_inv += 1
        else:
            investors = pd.DataFrame([{
                "Investor ID": _next_id(investors, "Investor ID", "INV"),
                "Company Name": company, "Country": country, "Sector": sector,
                "Relationship Manager": rm, "Account Manager": am,
                "Journey Stage": "Qualification",
            }])
            n_inv += 1

        # ── Opportunity Pipeline ──────────────────────────────────────────────
        if opp_text:
            co_opps = opps[opps["Company Name"] == company] \
                      if not opps.empty and "Company Name" in opps.columns else pd.DataFrame()
            # Keep both if text differs; skip only if exact same text already exists
            already = (
                not co_opps.empty and "Opportunity Name" in co_opps.columns and
                co_opps["Opportunity Name"].str.strip().str[:80].isin([opp_text[:80]]).any()
            )
            if not already:
                opps = pd.concat([opps, pd.DataFrame([{
                    "Opportunity ID":     _next_id(opps, "Opportunity ID", "OPP"),
                    "Company Name":       company,
                    "Investor ID":        "",
                    "Opportunity Name":   opp_text[:200],
                    "Sector":             sector,
                    "Opportunity Stage":  _maturity_to_stage(mat_level),
                    "Opportunity Status": "Active",
                    "Notes":              f"L-Level: {mat_level} | Priority: {opp_pri}",
                    "Last Updated":       date.today().isoformat(),
                }])], ignore_index=True)
                n_opp += 1

        # ── Action Items (from Next Step) ─────────────────────────────────────
        if next_step:
            co_acts = actions[actions["Company Name"] == company] \
                      if not actions.empty and "Company Name" in actions.columns else pd.DataFrame()
            already_act = (
                not co_acts.empty and "Action Description" in co_acts.columns and
                co_acts["Action Description"].str.strip().str[:80].isin([next_step[:80]]).any()
            )
            if not already_act:
                actions = pd.concat([actions, pd.DataFrame([{
                    "Action ID":          _next_id(actions, "Action ID", "ACT"),
                    "Company Name":       company,
                    "Investor ID":        "",
                    "Action Description": next_step[:200],
                    "Assigned To":        am or rm,
                    "Type of Engagement": "Outreach",
                    "Priority":           _opp_priority_to_action_priority(opp_pri),
                    "Status":             "Not Started",
                    "AM Input":           "",
                    "Remarks":            f"From Outreach — {mat_level}",
                    "Last Updated":       date.today().isoformat(),
                }])], ignore_index=True)
                n_act += 1

    dfs["Investor Master"]      = investors
    dfs["Opportunity Pipeline"] = opps
    dfs["Action Items"]         = actions
    return n_inv, n_opp, n_act


def _opp_priority_to_action_priority(opp_pri: str) -> str:
    return {"Very High": "High", "High": "High", "Medium": "Medium", "Low": "Low"}.get(opp_pri, "Medium")


def _maturity_to_stage(level: str) -> str:
    mapping = {
        "L0": "Early", "L1": "Early", "L2": "Development",
        "L3": "Development", "L4": "Ready", "L5": "Execution",
    }
    return mapping.get(level.strip().upper(), "Early")


# ── Excel export helpers ───────────────────────────────────────────────────────

def _write_outreach_sheet(ws, df: pd.DataFrame, company_type: str = "Foreign"):
    """Write one company sheet in the original Arabic tracker template format."""
    N_COLS = len(_ARABIC_COL_HEADERS)
    last_col = get_column_letter(N_COLS)

    green_fill  = PatternFill("solid", fgColor="1B5C3F")
    gold_fill   = PatternFill("solid", fgColor="C9974A")
    yellow_fill = PatternFill("solid", fgColor="FFF9C4")
    alt_fill    = PatternFill("solid", fgColor="F0F7F4")
    white_fill  = PatternFill("solid", fgColor="FFFFFF")
    white_font  = Font(color="FFFFFF", bold=True, size=9)
    dark_font   = Font(size=9)
    thin        = Side(style="thin")
    gold_side   = Side(style="medium", color="C9974A")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)
    gold_border = Border(left=gold_side, right=gold_side, top=gold_side, bottom=gold_side)
    center_al   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_al     = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    right_al    = Alignment(horizontal="right",  vertical="center", wrap_text=True)
    today       = date.today()

    # ── Row 1: Title ──────────────────────────────────────────────────────────
    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]
    c.value     = ("جدول التقييم الموحّد — الشركات الأجنبية المستهدفة"
                   if company_type == "Foreign"
                   else "جدول التقييم الموحّد — الشركات المحلية المستهدفة")
    c.fill      = green_fill
    c.font      = Font(color="FFFFFF", bold=True, size=14)
    c.alignment = center_al
    ws.row_dimensions[1].height = 28

    # ── Row 2: Subtitle ───────────────────────────────────────────────────────
    ws.merge_cells(f"A2:{last_col}2")
    c = ws["A2"]
    c.value     = (f"الشركاء الاستراتيجيون والمستثمرون الكبار  |  "
                   f"إعداد: مكتب التواصل التنفيذي  |  {today.strftime('%B %Y')}")
    c.fill      = green_fill
    c.font      = Font(color="FFFFFF", size=10)
    c.alignment = center_al
    ws.row_dimensions[2].height = 18

    # ── Row 3: Empty ──────────────────────────────────────────────────────────
    ws.row_dimensions[3].height = 6

    # ── Row 4: Group headers ──────────────────────────────────────────────────
    groups = [
        (1,  6,  "معلومات الشركة"),
        (7,  12, "معايير التقييم — الدرجة من ٥"),
        (13, 17, "نتيجة التقييم"),
        (18, 18, "مدير العلاقة\n(مكتب معاليه)"),
        (19, 20, "جهات التواصل"),
        (21, 22, "بيانات الاتصال"),
        (23, 29, "الفرصة والمتابعة ونضوج الصفقة"),
        (30, N_COLS, "متابعة التواصل (CRM)"),
    ]
    for sc, ec, label in groups:
        if sc < ec:
            ws.merge_cells(start_row=4, start_column=sc, end_row=4, end_column=ec)
        cell = ws.cell(row=4, column=sc, value=label)
        cell.fill      = green_fill
        cell.font      = Font(color="FFFFFF", bold=True, size=9)
        cell.alignment = center_al
    ws.row_dimensions[4].height = 26

    # ── Row 5: Weight note ────────────────────────────────────────────────────
    ws.merge_cells(f"A5:{last_col}5")
    c = ws["A5"]
    c.value     = "أوزان معايير التقييم:  م1=25%  |  م2=20%  |  م3=20%  |  م4=20%  |  م5=10%  |  م6=5%"
    c.fill      = PatternFill("solid", fgColor="EAF4EE")
    c.font      = Font(color="1B5C3F", size=9, italic=True)
    c.alignment = center_al
    ws.row_dimensions[5].height = 16

    # ── Row 6: Column headers ─────────────────────────────────────────────────
    for ci, hdr in enumerate(_ARABIC_COL_HEADERS, start=1):
        cell           = ws.cell(row=6, column=ci, value=hdr)
        cell.fill      = green_fill
        cell.font      = white_font
        cell.alignment = center_al
        cell.border    = border
    ws.row_dimensions[6].height = 36

    # ── Row 7: Warning note ───────────────────────────────────────────────────
    ws.merge_cells(f"A7:{last_col}7")
    c = ws["A7"]
    c.value     = "⚠ العمود R (مدير العلاقة من مكتب معاليه) فارغ بإطار ذهبي — يُعبَّأ يدوياً"
    c.fill      = yellow_fill
    c.font      = Font(color="996600", size=9, italic=True)
    c.alignment = right_al
    ws.row_dimensions[7].height = 16

    # ── Row 8: Empty ──────────────────────────────────────────────────────────
    ws.row_dimensions[8].height = 6

    if df.empty:
        _set_col_widths(ws)
        ws.sheet_view.rightToLeft = True
        return

    # Sort by priority (A→B→C→D→other) then company name
    _prio_order = {"A": 0, "B": 1, "C": 2, "D": 3}
    df = df.copy()
    df["_ps"] = df["Priority"].map(_prio_order).fillna(9)
    df = df.sort_values(["_ps", "Company Name"]).drop(columns=["_ps"])

    current_row     = 9
    current_priority = None
    seq             = 1

    _section_labels = {
        "A": "► المستوى A — أساسي",
        "B": "► المستوى B — مهم",
        "C": "► المستوى C — استكشافي",
        "D": "► المستوى D — منخفض",
    }
    _pri_cell_fills = {
        "A": PatternFill("solid", fgColor="D9EAD3"),
        "B": PatternFill("solid", fgColor="FCE5CD"),
        "C": PatternFill("solid", fgColor="CFE2F3"),
        "D": PatternFill("solid", fgColor="F4CCCC"),
    }

    for _, row in df.iterrows():
        pri = str(row.get("Priority", "") or "").strip()

        if pri != current_priority:
            ws.merge_cells(f"A{current_row}:{last_col}{current_row}")
            sec           = ws[f"A{current_row}"]
            sec.value     = _section_labels.get(pri, f"► المستوى {pri}")
            sec.fill      = PatternFill("solid", fgColor="D9EAD3")
            sec.font      = Font(color="1B5C3F", bold=True, size=10)
            sec.alignment = right_al
            ws.row_dimensions[current_row].height = 18
            current_priority = pri
            current_row += 1

        opp_pri_ar = _OPP_PRIORITY_REVERSE.get(
            str(row.get("Opportunity Priority", "") or ""),
            str(row.get("Opportunity Priority", "") or "")
        )
        mat_level = str(row.get("Maturity Level", "") or "").strip()
        rm_val    = str(row.get("RM", "") or "").strip()

        values = [
            seq,
            str(row.get("Company Name",          "") or ""),
            str(row.get("Sector",                 "") or ""),
            str(row.get("Country",                "") or ""),
            str(row.get("Company Size",           "") or ""),
            str(row.get("Classification",         "") or ""),
            str(row.get("Score M1",               "") or ""),
            str(row.get("Score M2",               "") or ""),
            str(row.get("Score M3",               "") or ""),
            str(row.get("Score M4",               "") or ""),
            str(row.get("Score M5",               "") or ""),
            str(row.get("Score M6",               "") or ""),
            str(row.get("Total Score",            "") or ""),
            pri,
            str(row.get("Proposed Channel",       "") or ""),
            str(row.get("Work Track",             "") or ""),
            str(row.get("Batch",                  "") or ""),
            rm_val,
            str(row.get("AM",                     "") or ""),
            str(row.get("Company Contact",        "") or ""),
            str(row.get("Phone",                  "") or ""),
            str(row.get("Email",                  "") or ""),
            str(row.get("Investment Opportunity", "") or ""),
            str(row.get("Next Step",              "") or ""),
            str(row.get("Assessment Notes",       "") or ""),
            mat_level,
            str(row.get("Maturity Description",   "") or ""),
            opp_pri_ar,
            str(row.get("Previous Minister Meeting", "") or ""),
            str(row.get("AM Presented",           "No") or "No"),
            str(row.get("AM Presented Date",      "") or ""),
            str(row.get("Minister Engaged",       "No") or "No"),
            str(row.get("Minister Engaged Date",  "") or ""),
        ]

        row_fill = alt_fill if seq % 2 == 0 else white_fill

        for ci, val in enumerate(values, start=1):
            cell           = ws.cell(row=current_row, column=ci, value=val)
            cell.font      = dark_font
            cell.border    = border
            cell.alignment = center_al if ci in (1, 7, 8, 9, 10, 11, 12, 13, 14) else left_al

            if ci == 14:                             # Priority
                cell.fill = _pri_cell_fills.get(pri, row_fill)
            elif ci == 26:                           # Maturity Level
                if mat_level in _MATURITY_COLORS:
                    cell.fill = PatternFill("solid", fgColor=_MATURITY_COLORS[mat_level].lstrip("#"))
                    cell.font = Font(color="FFFFFF", bold=True, size=9)
                    cell.alignment = center_al
                else:
                    cell.fill = row_fill
            elif ci == 18:                           # RM — gold border when empty
                cell.fill   = (yellow_fill if not rm_val else row_fill)
                cell.border = (gold_border if not rm_val else border)
            elif ci >= 30:                           # Workflow tracking cols
                cell.fill      = PatternFill("solid", fgColor="EAF4EE")
                cell.alignment = center_al
            else:
                cell.fill = row_fill

        ws.row_dimensions[current_row].height = 45
        current_row += 1
        seq += 1

    _set_col_widths(ws)
    ws.sheet_view.rightToLeft = True
    ws.freeze_panes = "C9"


def _set_col_widths(ws):
    widths = [
        5, 28, 18, 16, 14, 16,          # #, Company, Sector, Country, Size, Classification
        5, 5, 5, 5, 5, 5,               # M1-M6
        9, 7, 16, 14, 12,               # Score, Priority, Channel, Track, Batch
        18, 16, 20, 16, 24,             # RM, AM, Contact, Phone, Email
        36, 26, 24,                     # Opportunity, Next Step, Notes
        10, 22, 14, 22,                 # Maturity Level, Description, Opp Priority, Min Meeting
        12, 14, 14, 14,                 # Workflow: AM Presented, Date, Min Engaged, Date
    ]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ── Import prompt ──────────────────────────────────────────────────────────────

def _render_import_prompt():
    st.markdown(
        f'<div style="background:#fff;border-radius:12px;padding:40px;'
        f'text-align:center;margin-top:40px;border:2px dashed {_GOLD};">'
        f'<div style="font-size:40px;margin-bottom:12px;">📋</div>'
        f'<h3 style="color:{_GREEN};margin-bottom:8px;">No Outreach Data Loaded</h3>'
        f'<p style="color:#6B7280;font-size:14px;max-width:480px;margin:0 auto;">'
        f'Use the <strong>Import Outreach Excel</strong> button above to load the Arabic-format '
        f'strategic company tracking file (الشركات الأجنبية / الشركات المحلية).'
        f'</p>'
        f'<div style="margin-top:20px;padding:14px;background:#F0F7F4;border-radius:8px;'
        f'display:inline-block;text-align:left;">'
        f'<div style="font-weight:600;color:{_GREEN};margin-bottom:6px;font-size:12px;">'
        f'Expected Excel structure:</div>'
        f'<div style="font-size:12px;color:#374151;">'
        f'Sheet: الشركات الأجنبية (Foreign) &nbsp;|&nbsp; Sheet: الشركات المحلية (Local)<br>'
        f'Header row 6 · Data rows from row 10</div>'
        f'</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


# ── Cell normalisation ─────────────────────────────────────────────────────────

def _clean_cell(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    if s in ("—", "nan", "None", "NaT", "-"):
        return ""
    return s


def _normalize_maturity(raw: str) -> str:
    if not raw:
        return ""
    m = re.match(r"(L[0-5])", raw.strip(), re.IGNORECASE)
    return m.group(1).upper() if m else raw.strip()


def _normalize_opp_priority(raw: str) -> str:
    if not raw:
        return ""
    for ar_val, en_val in _OPP_PRIORITY_MAP.items():
        if ar_val in raw:
            return en_val
    return raw.strip()


def _normalize_minister_meeting(raw: str) -> str:
    if not raw:
        return "No"
    stripped = raw.strip()
    if stripped.startswith("نعم"):
        rest = stripped[3:].strip().lstrip("—").lstrip("-").strip()
        return f"Yes — {rest}" if rest else "Yes"
    if stripped in ("—", ""):
        return "No"
    return stripped


# ── Helpers ────────────────────────────────────────────────────────────────────

def _empty_outreach_df() -> pd.DataFrame:
    return pd.DataFrame(columns=_OUTPUT_COLS)


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
