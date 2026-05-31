
# Auto-save and auto-load for local session persistence.

from pathlib import Path
import pandas as pd

from data.schema import SCHEMA

SAVE_PATH = Path(__file__).parent.parent / "crm_data.xlsx"


def save_session(dfs: dict):
    """Write all current DataFrames to the local auto-save file."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        from openpyxl.utils import get_column_letter

        wb = Workbook()
        wb.remove(wb.active)

        for sheet_name, df in dfs.items():
            ws = wb.create_sheet(sheet_name)
            if df.empty:
                ws.append(list(df.columns) if len(df.columns) > 0 else ["(empty)"])
                continue
            ws.append(list(df.columns))
            fill = PatternFill("solid", fgColor="1B5C3F")
            font = Font(bold=True, color="FFFFFF")
            for cell in ws[1]:
                cell.fill = fill
                cell.font = font
            for _, row in df.iterrows():
                ws.append([_clean(v) for v in row])

        wb.save(SAVE_PATH)
    except Exception:
        pass  # Never block the UI on save failure


def load_session() -> dict | None:
    """Load the auto-saved file on startup. Returns None if not found."""
    if not SAVE_PATH.exists():
        return None
    try:
        raw = pd.ExcelFile(SAVE_PATH)
        dfs = {}
        for sheet, spec in SCHEMA.items():
            if sheet in raw.sheet_names:
                df = raw.parse(sheet)
                df.columns = [str(c).strip() for c in df.columns]
                dfs[sheet] = df
        if not dfs:
            return None
        if "RM Tasks" not in dfs:
            from modules.data_loader import _empty_tasks_df
            dfs["RM Tasks"] = _empty_tasks_df()
        return dfs
    except Exception:
        return None


def _clean(v):
    if isinstance(v, float) and pd.isna(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.date()
    return v
