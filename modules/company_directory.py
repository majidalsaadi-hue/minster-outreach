
# Company Directory — strategic card view of all tracked companies.

import hashlib
import pandas as pd
import streamlit as st
from datetime import date


_GREEN  = "#1B5C3F"
_GOLD   = "#C9974A"
_RED    = "#DC2626"
_AMBER  = "#D97706"
_BLUE   = "#1D4ED8"
_PURPLE = "#7C3AED"

# Sector → accent color for avatar ring
_SECTOR_COLORS = {
    "technology":       "#1D4ED8",
    "tech":             "#1D4ED8",
    "energy":           "#D97706",
    "renewable":        "#059669",
    "manufacturing":    "#7C3AED",
    "finance":          "#0891B2",
    "financial":        "#0891B2",
    "healthcare":       "#DC2626",
    "real estate":      "#92400E",
    "logistics":        "#374151",
    "tourism":          "#DB2777",
    "mining":           "#78350F",
    "agriculture":      "#166534",
}


def _sector_color(sector: str) -> str:
    s = sector.lower()
    for k, v in _SECTOR_COLORS.items():
        if k in s:
            return v
    # Deterministic fallback from sector name
    h = int(hashlib.md5(sector.encode()).hexdigest()[:6], 16)
    r = (h >> 16) & 0xFF
    g = (h >> 8)  & 0xFF
    b = h & 0xFF
    # Keep it dark enough for white text
    r = max(40, min(r, 160))
    g = max(40, min(g, 160))
    b = max(40, min(b, 160))
    return f"#{r:02X}{g:02X}{b:02X}"


def _initials(name: str) -> str:
    parts = name.split()
    if len(parts) >= 2:
        return (parts[0][0] + parts[-1][0]).upper()
    return name[:2].upper() if name else "?"


def render(dfs: dict, lang: str = "en"):
    st.markdown(
        f"<h2 style='color:{_GREEN};margin-bottom:2px;'>Investor Directory</h2>"
        f"<p style='color:#6b7280;margin-top:0;font-size:13px;'>"
        f"Strategic overview of all tracked investors — contact, sector & engagement</p>",
        unsafe_allow_html=True,
    )

    investors = dfs.get("Investor Master", pd.DataFrame())
    meetings  = dfs.get("Meeting Log",     pd.DataFrame())

    if investors.empty:
        st.info("No investor data loaded. Upload an Excel tracker to populate this view.")
        return

    # ── Build base table ──────────────────────────────────────────────────────
    keep_inv = [
        "Company Name", "Sector", "Country",
        "Relationship Status", "Journey Stage",
        "Key Contact Name", "Key Contact Title",
        "Next Meeting Date", "Last Meeting Date",
        "Account Manager", "Relationship Manager",
        "Strategic Priority Score",
    ]
    inv = investors[[c for c in keep_inv if c in investors.columns]].copy()

    # ── Pull key contact from Meeting Log attendees if not in master ──────────
    if not meetings.empty and "Company Name" in meetings.columns and "Investor Attendees" in meetings.columns:
        mtg_contact = (
            meetings[meetings["Investor Attendees"].notna()]
            .sort_values("Meeting Date" if "Meeting Date" in meetings.columns else "Company Name",
                         ascending=False)
            .drop_duplicates("Company Name")[["Company Name", "Investor Attendees"]]
            .rename(columns={"Investor Attendees": "_mtg_contact"})
        )
        inv = inv.merge(mtg_contact, on="Company Name", how="left")
    else:
        inv["_mtg_contact"] = None

    # ── Aggregate meeting stats ───────────────────────────────────────────────
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
        inv["# Meetings"]        = 0
        inv["Last Meeting Date"] = None

    for col in ("Next Meeting Date", "Last Meeting Date"):
        if col in inv.columns:
            inv[col] = pd.to_datetime(inv[col], errors="coerce").dt.date

    inv["# Meetings"]          = inv.get("# Meetings", pd.Series()).fillna(0).astype(int)
    inv["Sector"]              = inv.get("Sector",     pd.Series()).fillna("—")
    inv["Country"]             = inv.get("Country",    pd.Series()).fillna("—")
    inv["Relationship Status"] = inv.get("Relationship Status", pd.Series()).fillna("Active")
    inv["Journey Stage"]       = inv.get("Journey Stage", pd.Series()).fillna("—")

    # ── Filters ───────────────────────────────────────────────────────────────
    fc1, fc2, fc3, fc4 = st.columns([2.5, 1.5, 1.5, 1.5])
    with fc1:
        q = st.text_input("Search", placeholder="Search company or contact…",
                          key="cd_search", label_visibility="collapsed")
    with fc2:
        sectors = ["All sectors"] + sorted(s for s in inv["Sector"].unique() if s and s != "—")
        sel_sector = st.selectbox("Sector", sectors, key="cd_sector",
                                  label_visibility="collapsed")
    with fc3:
        countries = ["All countries"] + sorted(c for c in inv["Country"].unique() if c and c != "—")
        sel_country = st.selectbox("Country", countries, key="cd_country",
                                   label_visibility="collapsed")
    with fc4:
        statuses = ["All statuses"] + sorted(
            s for s in inv["Relationship Status"].unique() if s and s != "—"
        )
        sel_status = st.selectbox("Status", statuses, key="cd_status",
                                  label_visibility="collapsed")

    view = inv.copy()
    if q:
        mask = (
            view["Company Name"].str.contains(q, case=False, na=False)
            | view.get("Key Contact Name", pd.Series(dtype=str)).fillna("").str.contains(q, case=False, na=False)
        )
        view = view[mask]
    if sel_sector  != "All sectors":   view = view[view["Sector"] == sel_sector]
    if sel_country != "All countries": view = view[view["Country"] == sel_country]
    if sel_status  != "All statuses":  view = view[view["Relationship Status"] == sel_status]

    view = view.sort_values("Company Name").reset_index(drop=True)

    # ── Summary chips ─────────────────────────────────────────────────────────
    today    = date.today()
    total    = len(view)
    no_mtg   = int((view["# Meetings"] == 0).sum())
    has_next = int(view["Next Meeting Date"].notna().sum()) if "Next Meeting Date" in view.columns else 0
    overdue  = 0
    if "Next Meeting Date" in view.columns:
        overdue = int(
            view["Next Meeting Date"].apply(
                lambda d: isinstance(d, date) and d < today
            ).sum()
        )

    s1, s2, s3, s4 = st.columns(4)
    _chip(s1, str(total),    "Total Investors",    _GREEN)
    _chip(s2, str(no_mtg),   "No meetings yet",    _AMBER if no_mtg else _GREEN)
    _chip(s3, str(has_next), "Next meeting set",   _BLUE)
    _chip(s4, str(overdue),  "Overdue next mtg",   _RED if overdue else _GREEN)

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    if view.empty:
        st.warning("No companies match the current filters.")
        return

    # ── Card grid ─────────────────────────────────────────────────────────────
    cards_per_row = 3
    rows_data = [
        view.iloc[i : i + cards_per_row]
        for i in range(0, len(view), cards_per_row)
    ]

    for chunk in rows_data:
        cols = st.columns(cards_per_row)
        for col_idx, (_, row) in enumerate(chunk.iterrows()):
            with cols[col_idx]:
                _render_card(row, today)

    st.caption(f"{total} investor{'s' if total != 1 else ''} shown")


# ── Card renderer ─────────────────────────────────────────────────────────────

def _render_card(row, today: date):
    company  = str(row.get("Company Name", "—"))
    sector   = str(row.get("Sector",  "—"))
    country  = str(row.get("Country", "—"))
    status   = str(row.get("Relationship Status", "Active"))
    stage    = str(row.get("Journey Stage", "—"))
    n_mtg    = int(row.get("# Meetings", 0))
    last_mtg = row.get("Last Meeting Date")
    next_mtg = row.get("Next Meeting Date")
    am       = str(row.get("Account Manager", "—"))
    rm       = str(row.get("Relationship Manager", "—"))

    # Contact: prefer explicit columns, fall back to meeting attendees
    contact_name  = str(row.get("Key Contact Name",  "") or "").strip()
    contact_title = str(row.get("Key Contact Title", "") or "").strip()
    if not contact_name:
        mtg_att = str(row.get("_mtg_contact", "") or "").strip()
        if mtg_att:
            # First attendee only
            contact_name = mtg_att.split(",")[0].split("،")[0].strip()

    scolor = _sector_color(sector if sector != "—" else company)
    avatar = _initials(company)
    status_bg, status_fg = _status_colors(status)
    next_str, next_color = _next_date_fmt(next_mtg, today)
    last_str = _fmt_date(last_mtg) if last_mtg and str(last_mtg) not in ("NaT", "None", "nan") else "—"

    am_rm = ""
    if am != "—": am_rm = am
    if rm != "—": am_rm = f"{am_rm} / {rm}" if am_rm else rm

    # Priority star
    score = row.get("Strategic Priority Score")
    priority_dot = ""
    try:
        if score and float(score) >= 4:
            priority_dot = f'<span style="color:#C9974A;font-size:13px;margin-left:4px;" title="Strategic Priority">★</span>'
    except Exception:
        pass

    contact_html = ""
    if contact_name:
        contact_html = f"""
        <div style="display:flex;align-items:center;gap:6px;margin-top:6px;padding-top:6px;
                    border-top:1px solid #f0f0f0;">
          <div style="width:28px;height:28px;border-radius:50%;background:#f3f4f6;
                      display:flex;align-items:center;justify-content:center;
                      font-size:11px;font-weight:600;color:#374151;flex-shrink:0;">
            {_initials(contact_name)}
          </div>
          <div>
            <div style="font-size:12px;font-weight:600;color:#111827;line-height:1.2;">{contact_name}</div>
            {"<div style='font-size:10px;color:#6b7280;'>" + contact_title + "</div>" if contact_title else ""}
          </div>
        </div>"""

    next_html = (
        f'<span style="color:{next_color};font-size:11px;font-weight:600;">{next_str}</span>'
        if next_str != "—" else
        '<span style="color:#9ca3af;font-size:11px;">No next meeting</span>'
    )

    am_rm_html = (
        f'<div style="font-size:10px;color:#6b7280;margin-top:4px;">{am_rm}</div>'
        if am_rm else ""
    )

    card_html = f"""
    <div style="border:1px solid #e5e7eb;border-radius:12px;padding:16px;
                background:#fff;height:100%;box-sizing:border-box;
                transition:box-shadow .15s;
                box-shadow:0 1px 3px rgba(0,0,0,.07);">

      <!-- Header: avatar + company -->
      <div style="display:flex;align-items:flex-start;gap:12px;margin-bottom:10px;">
        <div style="width:44px;height:44px;border-radius:10px;
                    background:{scolor};flex-shrink:0;
                    display:flex;align-items:center;justify-content:center;
                    font-size:15px;font-weight:700;color:#fff;
                    letter-spacing:.5px;">
          {avatar}
        </div>
        <div style="flex:1;min-width:0;">
          <div style="font-size:14px;font-weight:700;color:{_GREEN};
                      white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
            {company}{priority_dot}
          </div>
          <div style="display:flex;gap:4px;flex-wrap:wrap;margin-top:4px;">
            <span style="background:#f0fdf4;color:{_GREEN};padding:1px 7px;
                         border-radius:8px;font-size:10px;font-weight:500;">{sector}</span>
            <span style="background:#f8fafc;color:#475569;padding:1px 7px;
                         border-radius:8px;font-size:10px;">🌍 {country}</span>
          </div>
        </div>
        <span style="background:{status_bg};color:{status_fg};padding:2px 8px;
                     border-radius:10px;font-size:10px;font-weight:600;
                     flex-shrink:0;white-space:nowrap;">{status}</span>
      </div>

      <!-- Contact person -->
      {contact_html}

      <!-- Meeting stats -->
      <div style="display:flex;justify-content:space-between;align-items:center;
                  margin-top:10px;padding-top:8px;border-top:1px solid #f0f0f0;">
        <div>
          <div style="font-size:10px;color:#9ca3af;text-transform:uppercase;
                      letter-spacing:.4px;margin-bottom:2px;">Last Meeting</div>
          <div style="font-size:11px;color:#374151;">{last_str}</div>
        </div>
        <div style="text-align:right;">
          <div style="font-size:10px;color:#9ca3af;text-transform:uppercase;
                      letter-spacing:.4px;margin-bottom:2px;">Next Meeting</div>
          {next_html}
        </div>
      </div>

      <!-- Footer: meetings count + AM/RM -->
      <div style="display:flex;justify-content:space-between;align-items:center;
                  margin-top:8px;padding-top:6px;border-top:1px solid #f0f0f0;">
        <span style="background:#f3f4f6;color:#374151;padding:1px 8px;
                     border-radius:8px;font-size:10px;font-weight:500;">
          {n_mtg} meeting{'s' if n_mtg != 1 else ''}
        </span>
        {am_rm_html}
      </div>
    </div>"""

    st.markdown(card_html, unsafe_allow_html=True)
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)


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
        return pd.to_datetime(d).strftime("%d %b %Y")
    except Exception:
        return str(d)


def _next_date_fmt(d, today: date) -> tuple[str, str]:
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
            return f"{label} · this week", _AMBER
        elif delta <= 30:
            return f"{label} · soon", _BLUE
        else:
            return label, "#374151"
    except Exception:
        return str(d), "#374151"


def _status_colors(status: str) -> tuple[str, str]:
    _MAP = {
        "active":    ("#d1fae5", "#065f46"),
        "inactive":  ("#f3f4f6", "#6b7280"),
        "on hold":   ("#fef3c7", "#d97706"),
        "churned":   ("#fee2e2", "#dc2626"),
        "prospect":  ("#ede9fe", "#7c3aed"),
        "committed": ("#dbeafe", "#1d4ed8"),
    }
    return _MAP.get(status.lower(), ("#f3f4f6", "#374151"))
