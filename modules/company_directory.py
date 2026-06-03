
# Company Directory — at-a-glance list of all companies with key meeting info.

import pandas as pd
import streamlit as st
from datetime import date


_GREEN      = "#1B5C3F"
_GOLD       = "#C9974A"
_RED        = "#DC2626"
_AMBER      = "#D97706"
_BLUE       = "#1D4ED8"


def render(dfs: dict, lang: str = "en"):
    st.markdown(
        f"<h2 style='color:{_GREEN};margin-bottom:4px;'>Company Directory</h2>"
        f"<p style='color:#6b7280;margin-top:0;'>All tracked companies — sector, country & meeting schedule</p>",
        unsafe_allow_html=True,
    )

    investors = dfs.get("Investor Master", pd.DataFrame())
    meetings  = dfs.get("Meeting Log",     pd.DataFrame())

    if investors.empty:
        st.info("No investor data loaded. Upload an Excel tracker to populate this view.")
        return

    # ── Build base table from Investor Master ─────────────────────────────────
    keep_inv = ["Company Name", "Sector", "Country", "Investor Tier",
                "Relationship Status", "Next Meeting Date", "Account Manager",
                "Relationship Manager"]
    inv = investors[[c for c in keep_inv if c in investors.columns]].copy()

    # ── Aggregate meeting stats from Meeting Log ──────────────────────────────
    if not meetings.empty and "Company Name" in meetings.columns:
        mtg = meetings.copy()
        if "Meeting Date" in mtg.columns:
            mtg["Meeting Date"] = pd.to_datetime(mtg["Meeting Date"], errors="coerce")
        agg = (
            mtg.groupby("Company Name", as_index=False)
               .agg(
                   Meetings=("Meeting ID", "count") if "Meeting ID" in mtg.columns
                            else ("Company Name", "count"),
                   Last_Meeting=("Meeting Date", "max") if "Meeting Date" in mtg.columns
                                else ("Company Name", "first"),
               )
        )
        agg.rename(columns={"Last_Meeting": "Last Meeting Date",
                             "Meetings":     "# Meetings"}, inplace=True)
        if "Last Meeting Date" in agg.columns:
            agg["Last Meeting Date"] = pd.to_datetime(
                agg["Last Meeting Date"], errors="coerce"
            ).dt.date
        inv = inv.merge(agg, on="Company Name", how="left")
    else:
        inv["# Meetings"]       = 0
        inv["Last Meeting Date"] = None

    if "Next Meeting Date" in inv.columns:
        inv["Next Meeting Date"] = pd.to_datetime(
            inv["Next Meeting Date"], errors="coerce"
        ).dt.date

    # ── Fill blanks ───────────────────────────────────────────────────────────
    inv["# Meetings"]       = inv.get("# Meetings",       pd.Series()).fillna(0).astype(int)
    inv["Sector"]           = inv.get("Sector",           pd.Series()).fillna("—")
    inv["Country"]          = inv.get("Country",          pd.Series()).fillna("—")
    inv["Investor Tier"]    = inv.get("Investor Tier",    pd.Series()).fillna("—")
    inv["Relationship Status"] = inv.get("Relationship Status", pd.Series()).fillna("—")

    # ── Filters row ───────────────────────────────────────────────────────────
    fc1, fc2, fc3, fc4 = st.columns([2, 1.5, 1.5, 1.5])
    with fc1:
        q = st.text_input("Search company", placeholder="Type to filter…",
                          key="cd_search", label_visibility="collapsed")
    with fc2:
        sectors = ["All sectors"] + sorted(
            s for s in inv["Sector"].unique() if s and s != "—"
        )
        sel_sector = st.selectbox("Sector", sectors, key="cd_sector",
                                  label_visibility="collapsed")
    with fc3:
        countries = ["All countries"] + sorted(
            c for c in inv["Country"].unique() if c and c != "—"
        )
        sel_country = st.selectbox("Country", countries, key="cd_country",
                                   label_visibility="collapsed")
    with fc4:
        statuses = ["All statuses"] + sorted(
            s for s in inv["Relationship Status"].unique() if s and s != "—"
        )
        sel_status = st.selectbox("Status", statuses, key="cd_status",
                                  label_visibility="collapsed")

    # Apply filters
    view = inv.copy()
    if q:
        view = view[view["Company Name"].str.contains(q, case=False, na=False)]
    if sel_sector != "All sectors":
        view = view[view["Sector"] == sel_sector]
    if sel_country != "All countries":
        view = view[view["Country"] == sel_country]
    if sel_status != "All statuses":
        view = view[view["Relationship Status"] == sel_status]

    view = view.sort_values("Company Name").reset_index(drop=True)

    # ── Summary chips ─────────────────────────────────────────────────────────
    today     = date.today()
    total     = len(view)
    no_mtg    = int((view["# Meetings"] == 0).sum())
    has_next  = view["Next Meeting Date"].notna().sum() if "Next Meeting Date" in view.columns else 0
    overdue   = 0
    if "Next Meeting Date" in view.columns:
        overdue = int(
            view["Next Meeting Date"].apply(
                lambda d: isinstance(d, date) and d < today
            ).sum()
        )

    s1, s2, s3, s4 = st.columns(4)
    _chip(s1, str(total),   "Companies",          _GREEN)
    _chip(s2, str(no_mtg),  "No meetings yet",    _AMBER if no_mtg else _GREEN)
    _chip(s3, str(has_next),"Next meeting set",   _BLUE)
    _chip(s4, str(overdue), "Overdue next mtg",   _RED if overdue else _GREEN)

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    # ── Table ─────────────────────────────────────────────────────────────────
    if view.empty:
        st.warning("No companies match the current filters.")
        return

    # Render as styled HTML table
    rows_html = []
    for _, row in view.iterrows():
        company  = str(row.get("Company Name", "—"))
        sector   = str(row.get("Sector",  "—"))
        country  = str(row.get("Country", "—"))
        tier     = str(row.get("Investor Tier", "—"))
        status   = str(row.get("Relationship Status", "Active"))
        n_mtg    = int(row.get("# Meetings", 0))
        last_mtg = row.get("Last Meeting Date")
        next_mtg = row.get("Next Meeting Date")
        am       = str(row.get("Account Manager", "—"))
        rm       = str(row.get("Relationship Manager", "—"))

        last_str = _fmt_date(last_mtg) if last_mtg and str(last_mtg) not in ("NaT", "None", "nan") else "—"
        next_str, next_color = _next_date_fmt(next_mtg, today)

        status_badge = _status_badge(status)
        mtg_badge    = (
            f'<span style="background:#f3f4f6;color:#374151;padding:1px 7px;'
            f'border-radius:10px;font-size:11px;">{n_mtg}</span>'
        )

        am_rm = f"{am}" if am != "—" else ""
        if rm != "—":
            am_rm = (am_rm + f" / {rm}") if am_rm else rm

        rows_html.append(f"""
        <tr>
          <td style="font-weight:600;color:{_GREEN};">{company}</td>
          <td>{sector}</td>
          <td>{country}</td>
          <td>{status_badge}</td>
          <td style="text-align:center;">{mtg_badge}</td>
          <td style="color:#374151;">{last_str}</td>
          <td style="color:{next_color};font-weight:{'600' if next_color != '#374151' else '400'};">{next_str}</td>
          <td style="font-size:11px;color:#6b7280;">{am_rm}</td>
        </tr>""")

    table_html = f"""
    <style>
      .cd-table {{
        width: 100%; border-collapse: collapse;
        font-size: 13px; font-family: system-ui, sans-serif;
      }}
      .cd-table th {{
        background: {_GREEN}; color: #fff;
        padding: 9px 12px; text-align: left;
        font-size: 12px; font-weight: 600; letter-spacing: .4px;
        position: sticky; top: 0;
      }}
      .cd-table td {{
        padding: 8px 12px; border-bottom: 1px solid #f0f0f0;
        vertical-align: middle;
      }}
      .cd-table tr:nth-child(even) td {{ background: #fafafa; }}
      .cd-table tr:hover td {{ background: #f0fdf4; }}
    </style>
    <div style="overflow-x:auto;border:1px solid #e5e7eb;border-radius:8px;
                max-height:600px;overflow-y:auto;">
      <table class="cd-table">
        <thead>
          <tr>
            <th>Company</th>
            <th>Sector</th>
            <th>Country</th>
            <th>Status</th>
            <th style="text-align:center;"># Meetings</th>
            <th>Last Meeting</th>
            <th>Next Meeting</th>
            <th>AM / RM</th>
          </tr>
        </thead>
        <tbody>
          {''.join(rows_html)}
        </tbody>
      </table>
    </div>
    """
    st.markdown(table_html, unsafe_allow_html=True)
    st.caption(f"{total} compan{'y' if total == 1 else 'ies'} shown")


# ── Helpers ───────────────────────────────────────────────────────────────────

def _chip(col, value: str, label: str, color: str):
    col.markdown(f"""
    <div style="background:{color}18;border:1px solid {color}40;border-radius:8px;
                padding:10px 14px;text-align:center;">
      <div style="font-size:22px;font-weight:700;color:{color};">{value}</div>
      <div style="font-size:11px;color:#6b7280;">{label}</div>
    </div>""", unsafe_allow_html=True)


def _fmt_date(d) -> str:
    try:
        if isinstance(d, date):
            return d.strftime("%d %b %Y")
        dt = pd.to_datetime(d)
        return dt.strftime("%d %b %Y")
    except Exception:
        return str(d)


def _next_date_fmt(d, today: date) -> tuple[str, str]:
    """Returns (display_string, color)."""
    if d is None or str(d) in ("NaT", "None", "nan", ""):
        return "—", "#9ca3af"
    try:
        if not isinstance(d, date):
            d = pd.to_datetime(d).date()
        delta = (d - today).days
        label = _fmt_date(d)
        if delta < 0:
            return f"{label} ⚠", _RED
        elif delta <= 7:
            return f"{label} (this week)", _AMBER
        elif delta <= 30:
            return f"{label} (soon)", _BLUE
        else:
            return label, "#374151"
    except Exception:
        return str(d), "#374151"


def _status_badge(status: str) -> str:
    _MAP = {
        "active":      ("#d1fae5", "#065f46"),
        "inactive":    ("#f3f4f6", "#6b7280"),
        "on hold":     ("#fef3c7", "#d97706"),
        "churned":     ("#fee2e2", "#dc2626"),
        "prospect":    ("#ede9fe", "#7c3aed"),
        "committed":   ("#dbeafe", "#1d4ed8"),
    }
    bg, fg = _MAP.get(status.lower(), ("#f3f4f6", "#374151"))
    return (
        f'<span style="background:{bg};color:{fg};padding:2px 8px;'
        f'border-radius:10px;font-size:11px;font-weight:500;">{status}</span>'
    )
