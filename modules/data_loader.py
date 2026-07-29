
# Data loading, validation, and parsing.
# Handles both the new enhanced schema and legacy per-investor sheets.

import io
import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime, date
from pathlib import Path
from data.schema import SCHEMA, LEGACY_SHEET_PREFIXES, LEGACY_HEADER_ROW_MARKERS

# Local auto-save file — sits next to app.py
_SAVE_PATH = Path(__file__).parent.parent / "crm_data.xlsx"


# ── Public API ───────────────────────────────────────────────────────────────

def save_session(dfs: dict):
    """Write the current session data to the local auto-save file."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        from openpyxl.utils import get_column_letter
        wb = Workbook()
        wb.remove(wb.active)
        for sheet_name, df in dfs.items():
            ws = wb.create_sheet(sheet_name)
            if df.empty:
                continue
            # Write header
            ws.append(list(df.columns))
            fill = PatternFill("solid", fgColor="1B5C3F")
            font = Font(bold=True, color="FFFFFF")
            for cell in ws[1]:
                cell.fill = fill
                cell.font = font
            # Write data
            for _, row in df.iterrows():
                ws.append([_clean_val(v) for v in row])
        wb.save(_SAVE_PATH)
    except Exception as exc:
        st.warning(f"Auto-save failed: {exc}")


def load_session() -> dict | None:
    """Load the auto-saved session file if it exists."""
    if not _SAVE_PATH.exists():
        return None
    try:
        raw = pd.ExcelFile(_SAVE_PATH)
        dfs = {}
        for sheet, spec in SCHEMA.items():
            if sheet in raw.sheet_names:
                df = raw.parse(sheet)
                df = _clean_df(df, spec.get("dtypes", {}))
                dfs[sheet] = df
        if "RM Tasks" not in dfs:
            dfs["RM Tasks"] = _empty_tasks_df()
        if "Contacts" not in dfs:
            dfs["Contacts"] = _empty_contacts_df()
        return dfs if dfs else None
    except Exception:
        return None


def _clean_val(v):
    if isinstance(v, float) and pd.isna(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.date()
    return v


# ── Public API ───────────────────────────────────────────────────────────────

def load_excel(uploaded_file) -> dict:
    """
    Parse an uploaded Excel file and return a dict of DataFrames keyed by
    sheet name.  Accepts both the new 5-sheet schema and the legacy
    per-investor action-item format, converting the latter automatically.
    Returns None and surfaces an error on st if the file is invalid.
    """
    try:
        raw = pd.ExcelFile(uploaded_file)
        sheet_names = raw.sheet_names

        if _is_legacy_format(sheet_names):
            return _load_legacy(raw, sheet_names)
        else:
            return _load_standard(raw)

    except Exception as exc:
        st.error(f"Failed to read Excel file: {exc}")
        return None


def validate_schema(dfs: dict) -> list[str]:
    """
    Validate loaded DataFrames against the expected schema.
    Returns a list of warning strings (empty = no issues).
    """
    warnings = []
    for sheet, spec in SCHEMA.items():
        if sheet not in dfs:
            if sheet != "RM Tasks":       # RM Tasks may be absent on first load
                warnings.append(f"Sheet '{sheet}' is missing.")
            continue
        df = dfs[sheet]
        for col in spec["required"]:
            if col not in df.columns:
                warnings.append(f"Sheet '{sheet}': required column '{col}' is missing.")
    return warnings


def get_summary(dfs: dict) -> dict:
    """Return high-level counts for the upload confirmation banner."""
    summary = {
        "investors":     0,
        "meetings":      0,
        "opportunities": 0,
        "actions":       0,
        "tasks":         0,
        "deals":         0,
        "contacts":      0,
    }
    if "Investor Master" in dfs:
        summary["investors"] = len(dfs["Investor Master"])
    if "Meeting Log" in dfs:
        summary["meetings"] = len(dfs["Meeting Log"])
    if "Opportunity Pipeline" in dfs:
        summary["opportunities"] = len(dfs["Opportunity Pipeline"])
    if "Action Items" in dfs:
        summary["actions"] = len(dfs["Action Items"])
    if "RM Tasks" in dfs:
        summary["tasks"] = len(dfs["RM Tasks"])
    if "Deal Progress" in dfs:
        summary["deals"] = len(dfs["Deal Progress"])
    if "Contacts" in dfs:
        summary["contacts"] = len(dfs["Contacts"])
    return summary


# ── Internal helpers ─────────────────────────────────────────────────────────

def _is_legacy_format(sheet_names: list[str]) -> bool:
    return any(s.startswith(prefix) for s in sheet_names for prefix in LEGACY_SHEET_PREFIXES)


def _load_standard(raw: pd.ExcelFile) -> dict:
    dfs = {}
    for sheet, spec in SCHEMA.items():
        if sheet in raw.sheet_names:
            df = raw.parse(sheet)
            df = _clean_df(df, spec.get("dtypes", {}))
            dfs[sheet] = df
    if "RM Tasks" not in dfs:
        dfs["RM Tasks"] = _empty_tasks_df()
    if "Contacts" not in dfs:
        dfs["Contacts"] = _empty_contacts_df()
    return dfs


def _load_legacy(raw: pd.ExcelFile, sheet_names: list[str]) -> dict:
    """
    Convert legacy per-investor sheets into the new 5-sheet structure.
    Best-effort mapping; missing fields are left blank.
    """
    action_rows   = []
    investor_rows = []
    opp_rows      = []
    investor_counter = 1
    opp_id_counter   = 1
    _OPP_SKIP = {"opportunity", "opportunities", "type", "n/a", "none", "",
                 "action item", "id", "assigned to", "remarks", "am input",
                 "priority", "progress", "start date", "due date", "type of engagement"}

    _STANDARD_SHEET_NAMES = set(SCHEMA.keys())

    for sheet_name in sheet_names:
        # Skip known standard CRM sheets; process everything else as a
        # potential legacy action-tracker tab (e.g. "CoC – Chamber of Commerce")
        if sheet_name in _STANDARD_SHEET_NAMES:
            continue

        df_raw = raw.parse(sheet_name, header=None)
        # Derive a fallback company name from the sheet name
        company = _extract_company_from_sheet_name(sheet_name)

        # Skip portfolio-level "deal" sheets — they have no per-investor header
        type_cell = _sstr(_find_header_value(df_raw, 10, 14, ["type"]))
        if type_cell.lower() == "deal":
            continue

        # MISA team — exact-label scan handles column shifts between file versions
        manager  = _find_header_value(df_raw, 11, 18, ["outreach"], exact=True)
        am_name  = _find_header_value(df_raw, 11, 18, ["am"],       exact=True)
        rm_name  = _find_header_value(df_raw, 11, 18, ["rm"],       exact=True)

        # Investor-side fields — keyword scan (column position varies by sheet width)
        company_name_xl = _find_header_value(df_raw, 11, 18, ["company name"])
        rep_name  = _find_header_value(df_raw, 11, 18, ["rep"])
        rep_pos   = _find_header_value(df_raw, 11, 18, ["postion", "position"])
        rep_email = _find_header_value(df_raw, 11, 18, ["email"])
        rep_phone = _find_header_value(df_raw, 11, 18, ["phone"])
        website   = _find_header_value(df_raw, 11, 18, ["website"])
        country   = _find_header_value(df_raw, 11, 18, ["country"])
        sector    = _find_header_value(df_raw, 11, 18, ["sector"])
        last_upd  = _find_header_value(df_raw, 11, 18, ["last updated"])
        next_mtg  = _find_header_value(df_raw, 11, 18, ["next meeting"])
        min_act   = _find_header_value(df_raw, 11, 22, ["minister action", "immediate action", "minister action required"])
        blocker   = _find_header_value(df_raw, 11, 22, ["blocker", "blocker level"])
        priority_hdr = _find_header_value(df_raw, 11, 18, ["priority"])
        file_owner   = _find_header_value(df_raw, 11, 18, ["file owner", "file_owner", "fileowner"])

        # Override sheet-name-derived company with the name in the Excel header
        if company_name_xl:
            company = _sstr(company_name_xl)

        inv_id = f"INV-{investor_counter:03d}"
        _min_act_str  = _sstr(min_act)  if min_act  else "None Required"
        _blocker_str  = _sstr(blocker)  if blocker  else "None"
        investor_rows.append({
            "Investor ID":                inv_id,
            "Company Name":               company,
            "Country":                    _sstr(country),
            "Sector":                     _sstr(sector),
            "Investor Tier":              "Tier 2 — High Potential",
            "Relationship Manager":       _sstr(rm_name),
            "Account Manager":            _sstr(am_name) or "TBD",
            "Outreach Manager":           _sstr(manager),
            "Journey Stage":              "Opportunity Matching",
            "Relationship Status":        "Active",
            "Est. Investment Value (SAR)":None,
            "Actual Commitment (SAR)":    None,
            "Last Meeting Date":          None,
            "Next Meeting Date":          _to_date(next_mtg),
            "Last Updated":               _to_date(last_upd),
            "Escalation Flag":            "None",
            "Notes":                      "",
            "Website":                    _sstr(website),
            "Company Rep":                _sstr(rep_name),
            "Rep Position":               _sstr(rep_pos),
            "Rep Email":                  _clean_email(rep_email),
            "Rep Phone":                  _sstr(rep_phone),
            "Minister Action Required":   _min_act_str,
            "Blocker Level":              _blocker_str,
            "Priority Classification":    _sstr(priority_hdr),
            "File Owner":                 _sstr(file_owner),
        })

        # ── Extract opportunities from header block (rows 14-19, col H) ──────
        for row_i in range(13, 19):          # 0-indexed rows 13-18 → Excel 14-19
            try:
                cell_val = df_raw.iloc[row_i, 7]   # col H = index 7
            except (IndexError, KeyError):
                continue
            if cell_val is None or (isinstance(cell_val, float) and pd.isna(cell_val)):
                continue
            sval = str(cell_val).strip()
            if not sval or sval.lower() in _OPP_SKIP:
                continue
            # Skip row-label artefacts
            if sval[:2] in ("4 ", "¦ ", "¹ ") or (sval and sval[0] in "¦¹"):
                continue
            opp_rows.append({
                "Opportunity ID":     f"OPP-{opp_id_counter:03d}",
                "Investor ID":        inv_id,
                "Company Name":       company,
                "Opportunity Name":   sval,
                "Sector":             _sstr(sector),
                "Opportunity Stage":  "Exploration",
                "Opportunity Status": "Active",
                "Opportunity Type":   "Opportunity",
                "Opportunity Source": "Excel Tracker Import",
                "Last Updated":       np.datetime64(datetime.now()),
            })
            opp_id_counter += 1

        # Find the data header row (contains "Action Item")
        header_row_idx = _find_header_row(df_raw)
        if header_row_idx is None:
            investor_counter += 1
            continue

        df_data = raw.parse(sheet_name, header=header_row_idx)
        df_data.columns = [str(c).strip() for c in df_data.columns]
        df_data = df_data.dropna(how="all")

        col_map = _build_legacy_col_map(df_data.columns)

        for _, row in df_data.iterrows():
            action_id_val = row.get(col_map.get("ID", ""), None)
            if pd.isna(action_id_val) or action_id_val == "ID":
                continue
            try:
                action_num = int(float(action_id_val))
            except (ValueError, TypeError):
                continue

            desc       = _safe_get(row, col_map, "Action Item", "")
            eng_type   = _safe_get(row, col_map, "Type of Engagement", "")
            sector     = _safe_get(row, col_map, "Sector", "")
            progress   = _normalise_progress(_safe_get(row, col_map, "Progress", 0))
            raw_status = _sstr(_safe_get(row, col_map, "Status", ""))
            # Use the explicit Status cell if present; otherwise derive from progress
            status     = _normalise_status(raw_status) if raw_status else _derive_status_from_progress(progress)

            action_rows.append({
                "Action ID":          f"ACT-{inv_id}-{action_num:03d}",
                "Investor ID":        inv_id,
                "Company Name":       company,
                "Meeting ID":         "",
                "Opportunity ID":     "",
                "Action Description": desc,
                "Assigned To":        _safe_get(row, col_map, "Assigned to", ""),
                "Department":         "",
                "Sector":             sector,
                "Type of Engagement": eng_type,
                "Start Date":         _to_date(_safe_get(row, col_map, "Start Date", None)),
                "Due Date":           _to_date(_safe_get(row, col_map, "Due Date", None)),
                "Priority":           _safe_get(row, col_map, "Priority", "Medium"),
                "Progress":           progress,
                "Status":             status,
                "Escalation Flag":    "None",
                "Remarks":            _safe_get(row, col_map, "Remarks", ""),
                "AM Input":           _safe_get(row, col_map, "AM Input", ""),
                "RM Input":           _safe_get(row, col_map, "RM Input", ""),
                "To Be In Dashboard": _safe_get(row, col_map, "To Be In Dashboard", ""),
                "Outcome":            "",
                "Next Action":        "",
                "Next Action Date":   None,
                "Last Updated":       _to_date(last_upd),
                "Updated By":         "",
            })


        investor_counter += 1

    investors_df = pd.DataFrame(investor_rows) if investor_rows else _empty_investor_df()
    actions_df   = pd.DataFrame(action_rows)   if action_rows   else _empty_actions_df()
    opps_df      = pd.DataFrame(opp_rows)      if opp_rows      else _empty_opportunity_df()

    return {
        "Investor Master":      investors_df,
        "Meeting Log":          _empty_meeting_df(),
        "Opportunity Pipeline": opps_df,
        "Action Items":         actions_df,
        "RM Tasks":             _empty_tasks_df(),
        "Contacts":             _empty_contacts_df(),
    }


# ── Column-name normalisation ─────────────────────────────────────────────────

def _build_legacy_col_map(cols: list[str]) -> dict:
    """Map logical field names to actual column names in the legacy sheet."""
    mapping = {}
    lookup = {c.lower().strip(): c for c in cols}
    candidates = {
        "ID":                ["id", "#", "no", "num"],
        "Action Item":       ["action item", "action", "task", "description"],
        "Assigned to":       ["assigned to", "owner", "responsible"],
        "Sector":            ["sector"],
        "Type of Engagement":["type of engagement", "type", "engagement type"],
        "Start Date":        ["start date"],
        "Due Date":          ["due date", "deadline", "target date"],
        "Priority":          ["priority"],
        "Progress":          ["progress", "completion", "complete"],
        "Status":            ["status", "update status", "update the status"],
        "Remarks":           ["remarks", "notes", "comments"],
        "AM Input":          ["am input", "am notes", "account manager input"],
        "RM Input":          ["rm input", "rm notes", "relationship manager input"],
        "To Be In Dashboard":["to be in dashbaord", "to be in dashboard", "in dashboard"],
    }
    for logical, options in candidates.items():
        for opt in options:
            if opt in lookup:
                mapping[logical] = lookup[opt]
                break
    return mapping


def _find_header_row(df_raw: pd.DataFrame) -> int | None:
    for i, row in df_raw.iterrows():
        vals = [str(v).strip().lower() for v in row if not pd.isna(v)]
        if any(m.lower() in vals for m in LEGACY_HEADER_ROW_MARKERS):
            return i
    return None


def _extract_company_from_sheet_name(sheet_name: str) -> str:
    for prefix in LEGACY_SHEET_PREFIXES:
        if sheet_name.startswith(prefix):
            return sheet_name[len(prefix):].strip()
    return sheet_name.strip()


# ── Type coercion helpers ─────────────────────────────────────────────────────

def _clean_df(df: pd.DataFrame, dtypes: dict) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    for col, dtype in dtypes.items():
        if col in df.columns and dtype == "date":
            df[col] = pd.to_datetime(df[col], errors="coerce")
        elif col in df.columns and dtype == "float":
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def _to_date(val) -> date | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, (datetime, date)):
        return val if isinstance(val, date) else val.date()
    try:
        return pd.to_datetime(val).date()
    except Exception:
        return None


def _normalise_progress(val) -> str:
    try:
        s = str(val).replace("%", "").strip() if val is not None else "0"
        f = float(s)
    except (TypeError, ValueError):
        return "0%"
    pct = round(f * 100) if f <= 1.0 else round(f)
    pct = max(0, min(100, pct))
    return f"{pct}%"


def _derive_status_from_progress(progress_str: str) -> str:
    """Derive status from progress percentage: 100% → Completed, >5% → In Progress, ≤5% → Not Started."""
    try:
        pct = float(str(progress_str).rstrip("%"))
    except (ValueError, TypeError):
        pct = 0.0
    if pct >= 100:
        return "Completed"
    elif pct > 5:
        return "In Progress"
    else:
        return "Not Started"


def _normalise_status(val: str) -> str:
    mapping = {
        "inprogress":  "In Progress",
        "in progress": "In Progress",
        "notstarted":  "Not Started",
        "not started": "Not Started",
        "due":         "Not Started",
        "completed":   "Completed",
        "blocked":     "Blocked",
        "cancelled":   "Cancelled",
    }
    return mapping.get(str(val).lower().strip(), str(val).strip())


def _safe_cell(df: pd.DataFrame, row: int, col: int):
    try:
        val = df.iloc[row, col]
        return None if pd.isna(val) else val
    except (IndexError, KeyError):
        return None


def _safe_get(row, col_map: dict, key: str, default):
    col = col_map.get(key)
    if col is None:
        return default
    val = row.get(col, default)
    if pd.isna(val) if not isinstance(val, str) else False:
        return default
    return val


# ── Empty DataFrame factories ─────────────────────────────────────────────────

def _empty_investor_df() -> pd.DataFrame:
    cols = [c for spec in [SCHEMA["Investor Master"]]
            for c in spec["required"] + spec["optional"]]
    return pd.DataFrame(columns=cols)


def _empty_meeting_df() -> pd.DataFrame:
    cols = [c for spec in [SCHEMA["Meeting Log"]]
            for c in spec["required"] + spec["optional"]]
    return pd.DataFrame(columns=cols)


def _empty_opportunity_df() -> pd.DataFrame:
    cols = [c for spec in [SCHEMA["Opportunity Pipeline"]]
            for c in spec["required"] + spec["optional"]]
    return pd.DataFrame(columns=cols)


def _empty_actions_df() -> pd.DataFrame:
    cols = [c for spec in [SCHEMA["Action Items"]]
            for c in spec["required"] + spec["optional"]]
    return pd.DataFrame(columns=cols)


def _empty_tasks_df() -> pd.DataFrame:
    cols = [c for spec in [SCHEMA["RM Tasks"]]
            for c in spec["required"] + spec["optional"]]
    return pd.DataFrame(columns=cols)


def _empty_contacts_df() -> pd.DataFrame:
    cols = [c for spec in [SCHEMA["Contacts"]]
            for c in spec["required"] + spec["optional"]]
    return pd.DataFrame(columns=cols)


def _find_header_value(df: pd.DataFrame, start_row: int, end_row: int, keywords: list, exact: bool = False):
    """Scan rows start_row..end_row for a cell matching any keyword; return the adjacent right cell.
    exact=True requires the full cell text to equal one of the keywords."""
    kw_lower = [k.lower() for k in keywords]
    for row_i in range(start_row, min(end_row, len(df))):
        for col_i in range(df.shape[1] - 1):
            val = df.iloc[row_i, col_i]
            if val is None or (isinstance(val, float) and pd.isna(val)):
                continue
            cell_str = str(val).replace("\xa0", " ").lower().strip().lstrip("¦¹ ")
            matched = (cell_str in kw_lower) if exact else any(kw in cell_str for kw in kw_lower)
            if matched:
                right = df.iloc[row_i, col_i + 1]
                if right is not None and not (isinstance(right, float) and pd.isna(right)):
                    return right
    return None


def _sstr(val) -> str:
    """Convert any cell value to clean string; empty string for None/NaN."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).replace("\xa0", " ").strip()


def _clean_email(val) -> str:
    """Strip 'Email: ' or 'Email:\t' prefix that appears in raw tracker cells."""
    import re
    s = _sstr(val)
    s = re.sub(r'^email\s*:\s*', '', s, flags=re.IGNORECASE).strip()
    return s
