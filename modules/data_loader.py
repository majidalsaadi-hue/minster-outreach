
# Data loading, validation, and parsing.
# Handles both the new enhanced schema and legacy per-investor sheets.

import io
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
    _OPP_SKIP = {"opportunity", "opportunities", "type", "n/a", "none", ""}

    for sheet_name in sheet_names:
        if not any(sheet_name.startswith(p) for p in LEGACY_SHEET_PREFIXES):
            continue

        df_raw = raw.parse(sheet_name, header=None)
        company = _extract_company_from_sheet_name(sheet_name)

        # Extract header metadata from rows 13-16 (0-indexed: 12-15)
        manager = _safe_cell(df_raw, 13, 11)
        am_name  = _safe_cell(df_raw, 14, 11)
        rm_name  = _safe_cell(df_raw, 15, 11)
        last_upd = _safe_cell(df_raw, 12, 16)
        next_mtg = _safe_cell(df_raw, 15, 16)

        inv_id = f"INV-{investor_counter:03d}"
        investor_rows.append({
            "Investor ID":              inv_id,
            "Company Name":             company,
            "Country":                  "",
            "Sector":                   "",
            "Investor Tier":            "Tier 2 — High Potential",
            "Relationship Manager":     rm_name or "",
            "Account Manager":          am_name or "TBD",
            "Outreach Manager":         manager or "",
            "Journey Stage":            "Opportunity Matching",
            "Relationship Status":      "Active",
            "Est. Investment Value (SAR)": None,
            "Actual Commitment (SAR)":   None,
            "Last Meeting Date":        None,
            "Next Meeting Date":        _to_date(next_mtg),
            "Last Updated":             _to_date(last_upd),
            "Escalation Flag":          "None",
            "Notes":                    "",
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
                "Sector":             "",
                "Opportunity Stage":  "Exploration",
                "Opportunity Status": "Active",
                "Opportunity Type":   "Opportunity",
                "Opportunity Source": "Excel Tracker Import",
                "Last Updated":       date.today(),
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

            desc    = _safe_get(row, col_map, "Action Item", "")
            eng_type = _safe_get(row, col_map, "Type of Engagement", "")
            sector  = _safe_get(row, col_map, "Sector", "")

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
                "Progress":           _normalise_progress(_safe_get(row, col_map, "Progress", 0)),
                "Status":             _normalise_status(_safe_get(row, col_map, "Status", "Not Started")),
                "Escalation Flag":    "None",
                "Remarks":            _safe_get(row, col_map, "Remarks", ""),
                "Outcome":            "",
                "Next Action":        "",
                "Next Action Date":   None,
                "Last Updated":       _to_date(last_upd),
                "Updated By":         "",
            })

            # Action items with Opportunity engagement type → also add to pipeline
            if str(eng_type).strip().lower() == "opportunity" and str(desc).strip():
                existing_names = {o["Opportunity Name"] for o in opp_rows}
                if desc not in existing_names:
                    opp_rows.append({
                        "Opportunity ID":     f"OPP-{opp_id_counter:03d}",
                        "Investor ID":        inv_id,
                        "Company Name":       company,
                        "Opportunity Name":   str(desc).strip(),
                        "Sector":             sector,
                        "Opportunity Stage":  "Exploration",
                        "Opportunity Status": "Active",
                        "Opportunity Type":   "Opportunity",
                        "Opportunity Source": "Action Item Import",
                        "Last Updated":       date.today(),
                    })
                    opp_id_counter += 1

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
    }


# ── Column-name normalisation ─────────────────────────────────────────────────

def _build_legacy_col_map(cols: list[str]) -> dict:
    """Map logical field names to actual column names in the legacy sheet."""
    mapping = {}
    lookup = {c.lower().strip(): c for c in cols}
    candidates = {
        "ID":                ["id"],
        "Action Item":       ["action item"],
        "Assigned to":       ["assigned to"],
        "Sector":            ["sector"],
        "Type of Engagement":["type of engagement"],
        "Start Date":        ["start date"],
        "Due Date":          ["due date"],
        "Priority":          ["priority"],
        "Progress":          ["progress"],
        "Status":            ["status"],
        "Remarks":           ["remarks"],
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
        f = float(val)
    except (TypeError, ValueError):
        return "0%"
    pct = round(f * 100) if f <= 1.0 else round(f)
    # Snap to nearest 25
    snapped = round(pct / 25) * 25
    snapped = max(0, min(100, snapped))
    return f"{snapped}%"


def _normalise_status(val: str) -> str:
    mapping = {
        "inprogress":  "In Progress",
        "in progress": "In Progress",
        "notstarted":  "Not Started",
        "not started": "Not Started",
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
