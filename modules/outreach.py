
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

_ARABIC_HEADERS = [
    "#", "اسم الشركة", "القطاع", "الدولة/المدينة", "الحجم",
    "نوع التصنيف", "م١", "م٢", "م٣", "م٤", "م٥", "م٦",
    "الدرجة الكلية ٪", "الأولوية", "القناة المقترحة", "مسار العمل",
    "الدفعة", "مدير العلاقة من مكتب معاليه", "مدير الحساب من MISA",
    "النظير من الشركة FOC", "هاتف التواصل", "البريد الإلكتروني",
    "الفرصة الاستثمارية", "الخطوة التالية", "ملاحظات التقييم",
    "مستوى النضوج L0→L5", "وصف النضوج والملاحظة",
    "أولوية الفرصة", "اجتماع سابق مع معاليه",
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

    for company_type, sheet_title in [("Foreign", "Foreign Companies"), ("Local", "Local Companies")]:
        subset = df[df["Company Type"] == company_type].reset_index(drop=True) \
                 if not df.empty and "Company Type" in df.columns else pd.DataFrame()
        ws = wb.create_sheet(sheet_title)
        _write_outreach_sheet(ws, subset)

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
        n_inv, n_opp = _sync_to_crm(dfs)
        save_session(dfs)
        st.success(f"✅ Sync complete — {n_inv} investor records updated, {n_opp} opportunities created.")

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
        return 0, 0

    investors = dfs.get("Investor Master", pd.DataFrame()).copy()
    opps      = dfs.get("Opportunity Pipeline", pd.DataFrame()).copy()

    n_inv_updated = 0
    n_opp_created = 0

    for _, row in tracker.iterrows():
        company = str(row.get("Company Name", "")).strip()
        if not company:
            continue

        am      = str(row.get("AM",             "") or "")
        rm      = str(row.get("RM",             "") or "")
        phone   = str(row.get("Phone",          "") or "")
        email   = str(row.get("Email",          "") or "")
        contact = str(row.get("Company Contact","") or "")
        opp_text= str(row.get("Investment Opportunity","") or "").strip()

        if not investors.empty and "Company Name" in investors.columns:
            match_mask = investors["Company Name"] == company
            if match_mask.any():
                if am:
                    investors.loc[match_mask, "Account Manager"]     = am
                if rm:
                    investors.loc[match_mask, "Relationship Manager"]= rm
                if phone:
                    investors.loc[match_mask, "Rep Phone"]           = phone
                if email:
                    investors.loc[match_mask, "Rep Email"]           = email
                if contact:
                    investors.loc[match_mask, "Company Rep"]         = contact
                n_inv_updated += 1
            else:
                new_inv = {
                    "Investor ID":         _next_id(investors, "Investor ID", "INV"),
                    "Company Name":        company,
                    "Country":             str(row.get("Country", "") or ""),
                    "Sector":              str(row.get("Sector",  "") or ""),
                    "Relationship Manager":rm,
                    "Account Manager":     am,
                    "Rep Phone":           phone,
                    "Rep Email":           email,
                    "Company Rep":         contact,
                    "Journey Stage":       "Qualification",
                }
                investors = pd.concat(
                    [investors, pd.DataFrame([new_inv])], ignore_index=True
                )
                n_inv_updated += 1
        else:
            investors = pd.DataFrame([{
                "Investor ID":         "INV-001",
                "Company Name":        company,
                "Country":             str(row.get("Country", "") or ""),
                "Sector":              str(row.get("Sector",  "") or ""),
                "Relationship Manager":rm,
                "Account Manager":     am,
                "Journey Stage":       "Qualification",
            }])
            n_inv_updated += 1

        if opp_text:
            existing_opps = opps[opps["Company Name"] == company] \
                            if not opps.empty and "Company Name" in opps.columns \
                            else pd.DataFrame()
            if existing_opps.empty:
                new_opp = {
                    "Opportunity ID":     _next_id(opps, "Opportunity ID", "OPP"),
                    "Company Name":       company,
                    "Investor ID":        "",
                    "Opportunity Name":   opp_text[:120],
                    "Sector":             str(row.get("Sector", "") or ""),
                    "Opportunity Stage":  _maturity_to_stage(str(row.get("Maturity Level", "") or "")),
                    "Opportunity Status": "Active",
                    "Last Updated":       date.today().isoformat(),
                }
                opps = pd.concat(
                    [opps, pd.DataFrame([new_opp])], ignore_index=True
                )
                n_opp_created += 1

    dfs["Investor Master"]      = investors
    dfs["Opportunity Pipeline"] = opps
    return n_inv_updated, n_opp_created


def _maturity_to_stage(level: str) -> str:
    mapping = {
        "L0": "Early", "L1": "Early", "L2": "Development",
        "L3": "Development", "L4": "Ready", "L5": "Execution",
    }
    return mapping.get(level.strip().upper(), "Early")


# ── Excel export helpers ───────────────────────────────────────────────────────

def _write_outreach_sheet(ws, df: pd.DataFrame):
    green_fill  = PatternFill("solid", fgColor="1B5C3F")
    gold_fill   = PatternFill("solid", fgColor="C9974A")
    alt_fill    = PatternFill("solid", fgColor="F0F7F4")
    white_fill  = PatternFill("solid", fgColor="FFFFFF")
    white_font  = Font(color="FFFFFF", bold=True, size=9)
    dark_font   = Font(size=9)
    center_al   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_al     = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin        = Side(style="thin")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)

    header_cols = [
        "Outreach ID", "Company Name", "Sector", "Country", "Company Size",
        "Priority", "Total Score", "Maturity Level",
        "RM", "AM", "Company Contact", "Phone", "Email",
        "Investment Opportunity", "Next Step",
        "Opportunity Priority", "Previous Minister Meeting",
        "AM Presented", "AM Presented Date",
        "Minister Engaged", "Minister Engaged Date",
    ]

    for col_idx, col_name in enumerate(header_cols, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill      = green_fill
        cell.font      = white_font
        cell.alignment = center_al
        cell.border    = border

    ws.row_dimensions[1].height = 22

    if df.empty:
        return

    for row_idx, (_, row) in enumerate(df.iterrows(), start=2):
        row_fill = alt_fill if row_idx % 2 == 0 else white_fill
        mat_level = str(row.get("Maturity Level", "") or "").strip()

        for col_idx, col_name in enumerate(header_cols, start=1):
            val = row.get(col_name, "")
            if val in (None, "nan", "NaT"):
                val = ""
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font      = dark_font
            cell.alignment = center_al if col_idx <= 2 else left_al
            cell.border    = border

            if col_name == "Maturity Level" and mat_level in _MATURITY_COLORS:
                hex_color = _MATURITY_COLORS[mat_level].lstrip("#")
                cell.fill = PatternFill("solid", fgColor=hex_color)
                cell.font = Font(color="FFFFFF", bold=True, size=9)
            elif col_name == "Priority":
                pri = str(val).strip()
                hex_c = {"A": "DC2626", "B": "D97706", "C": "1D4ED8"}.get(pri)
                if hex_c:
                    cell.fill = PatternFill("solid", fgColor=hex_c)
                    cell.font = Font(color="FFFFFF", bold=True, size=9)
                else:
                    cell.fill = row_fill
            else:
                cell.fill = row_fill

        ws.row_dimensions[row_idx].height = 18

    col_widths = [10, 30, 18, 14, 12, 7, 10, 10, 16, 16, 20, 16, 24, 36, 26, 14, 22, 12, 16, 12, 16]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "C2"


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
