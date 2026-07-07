
# Contact Management — multiple contacts per investor company.

import io
import re
import pandas as pd
import streamlit as st
from datetime import date

from config.settings import MISA_GREEN, MISA_GOLD
from modules.persistence import save_session

_GREEN = "#1B5C3F"
_GOLD  = "#C9974A"

_CONTACT_TYPES = ["Primary", "Secondary", "Technical", "Executive", "Legal", "Finance"]

_TYPE_COLOR = {
    "Primary":   (_GREEN, "#f0fdf4"),
    "Executive": ("#7C3AED", "#f5f3ff"),
    "Technical": ("#0891B2", "#ecfeff"),
    "Legal":     ("#DC2626", "#fef2f2"),
    "Finance":   ("#D97706", "#fffbeb"),
    "Secondary": ("#6B7280", "#f9fafb"),
}

# Flexible column-name aliases for the Excel upload
_COL_ALIASES = {
    "Full Name":    ["full name", "name", "contact name", "person", "contact person"],
    "Company Name": ["company name", "company", "organization", "organisation", "investor", "fund"],
    "Title":        ["title", "job title", "position", "role"],
    "Department":   ["department", "dept", "division", "team"],
    "Email":        ["email", "email address", "e-mail", "mail"],
    "Phone":        ["phone", "phone number", "mobile", "tel", "telephone", "cell"],
    "LinkedIn":     ["linkedin", "linkedin url", "linkedin profile", "linkedin link"],
    "Contact Type": ["contact type", "type", "relationship", "category"],
    "Notes":        ["notes", "note", "comments", "remarks", "comment"],
}


def render(dfs: dict, lang: str):
    contacts  = dfs.get("Contacts",       pd.DataFrame())
    investors = dfs.get("Investor Master", pd.DataFrame())

    st.markdown(
        f"<h2 style='color:{_GREEN};margin-bottom:2px;'>Contact Directory</h2>"
        f"<p style='color:#6b7280;font-size:13px;margin-top:0;'>"
        f"Manage multiple contacts per investor company</p>",
        unsafe_allow_html=True,
    )

    # ── Action toolbar ─────────────────────────────────────────────────────────
    tb1, tb2, tb3 = st.columns([1, 1, 2])
    with tb1:
        if st.button("⚡ Sync from Investor Master", use_container_width=True,
                     help="Pull Rep / Key Contact fields from each investor record"):
            added = _sync_from_investor_master(dfs)
            if added:
                save_session(dfs)
                st.success(f"✅ {added} new contact{'s' if added != 1 else ''} imported from Investor Master.")
                st.rerun()
            else:
                st.info("All investor contacts are already in the directory.")
    with tb2:
        if st.button("+ Add New Contact", use_container_width=True):
            st.session_state["ct_show_add"] = not st.session_state.get("ct_show_add", False)

    if st.session_state.get("ct_show_add", False):
        with st.container():
            _add_contact_form(dfs, investors)

    # ── Upload Contacts Excel ──────────────────────────────────────────────────
    with st.expander("📤 Upload Contacts Excel", expanded=False):
        _render_upload_panel(dfs, investors)

    contacts = dfs.get("Contacts", pd.DataFrame())
    if contacts.empty:
        st.info("No contacts yet. Use **Sync from Investor Master** or upload an Excel file above.")
        return

    # ── Filters ────────────────────────────────────────────────────────────────
    sf1, sf2, sf3 = st.columns([2.5, 1.5, 1.5])
    with sf1:
        q = st.text_input("Search contacts", placeholder="Name, company, email…",
                          key="ct_search", label_visibility="collapsed")
    with sf2:
        companies = ["All companies"] + sorted(
            c for c in contacts.get("Company Name", pd.Series()).dropna().unique() if c
        )
        sel_co = st.selectbox("Company", companies, key="ct_co", label_visibility="collapsed")
    with sf3:
        types = ["All types"] + _CONTACT_TYPES
        sel_type = st.selectbox("Type", types, key="ct_type", label_visibility="collapsed")

    view = contacts.copy()
    if q:
        search_cols = [c for c in ["Full Name", "Company Name", "Email", "Title"] if c in view.columns]
        if search_cols:
            mask = view[search_cols].apply(
                lambda col: col.fillna("").astype(str).str.lower().str.contains(q.lower(), regex=False)
            ).any(axis=1)
            view = view[mask]
    if sel_co != "All companies":
        view = view[view.get("Company Name", pd.Series()) == sel_co]
    if sel_type != "All types":
        view = view[view.get("Contact Type", pd.Series()) == sel_type]
    view = view.reset_index(drop=True)

    st.caption(f"{len(view)} contact{'s' if len(view) != 1 else ''} · {len(contacts)} total")
    if view.empty:
        st.warning("No contacts match the current filters.")
        return

    # ── Card grid ─────────────────────────────────────────────────────────────
    cols_per_row = 3
    for start in range(0, len(view), cols_per_row):
        chunk = view.iloc[start:start + cols_per_row]
        cols  = st.columns(cols_per_row)
        for ci, (_, row) in enumerate(chunk.iterrows()):
            with cols[ci]:
                _render_contact_card(row, dfs)


def render_for_company(company: str, dfs: dict):
    """Contacts sub-panel used inside the investor profile tab."""
    contacts = dfs.get("Contacts", pd.DataFrame())

    if contacts.empty or "Company Name" not in contacts.columns:
        co_contacts = pd.DataFrame()
    else:
        co_contacts = contacts[contacts["Company Name"] == company].reset_index(drop=True)

    with st.expander(f"+ Add contact for {company}", expanded=False):
        _add_contact_form(dfs, dfs.get("Investor Master", pd.DataFrame()), default_company=company)

    if co_contacts.empty:
        # Check if we can auto-suggest a sync for this company
        inv = dfs.get("Investor Master", pd.DataFrame())
        can_sync = False
        if not inv.empty and "Company Name" in inv.columns:
            row_inv = inv[inv["Company Name"] == company]
            if not row_inv.empty:
                r = row_inv.iloc[0]
                rep = str(r.get("Company Rep", "") or "").strip()
                kc  = str(r.get("Key Contact Name", "") or "").strip()
                if rep or kc:
                    can_sync = True
                    st.markdown(
                        f'<div style="background:#fffbeb;border:1px solid #fde68a;'
                        f'border-radius:6px;padding:8px 12px;font-size:12px;color:#92400e;margin-bottom:8px;">'
                        f'Investor Master has contact data — '
                        f'{("Rep: " + rep) if rep else ""}'
                        f'{(" · Key Contact: " + kc) if kc else ""}'
                        f'. Click below to import.</div>',
                        unsafe_allow_html=True,
                    )
        if can_sync:
            if st.button(f"⚡ Import contacts for {company}", key=f"sync_{company}"):
                added = _sync_from_investor_master(dfs, only_company=company)
                if added:
                    save_session(dfs)
                    st.success(f"✅ {added} contact{'s' if added != 1 else ''} imported.")
                    st.rerun()
        else:
            st.caption("No contacts linked yet — add one above.")
        return

    cols_per_row = 2
    for start in range(0, len(co_contacts), cols_per_row):
        chunk = co_contacts.iloc[start:start + cols_per_row]
        cols  = st.columns(cols_per_row)
        for ci, (_, row) in enumerate(chunk.iterrows()):
            with cols[ci]:
                _render_contact_card(row, dfs)


# ── Upload panel ──────────────────────────────────────────────────────────────

def _render_upload_panel(dfs: dict, investors: pd.DataFrame):
    st.markdown(
        "<p style='font-size:13px;color:#6B7280;margin-bottom:8px;'>"
        "Upload an Excel file (.xlsx) with a <b>Contacts</b> sheet or a flat table. "
        "Required columns: <b>Full Name</b> and <b>Company Name</b>. "
        "Duplicates (same name + company) are skipped automatically.</p>",
        unsafe_allow_html=True,
    )

    # Template download
    template_bytes = _build_template_excel()
    st.download_button(
        "⬇ Download template",
        data=template_bytes,
        file_name="contacts_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )

    uploaded = st.file_uploader(
        "Choose Excel file",
        type=["xlsx", "xls"],
        key="ct_upload",
        label_visibility="collapsed",
    )
    if uploaded is None:
        return

    preview_df, warnings = _parse_contacts_excel(uploaded)
    if preview_df is None:
        return

    if warnings:
        for w in warnings:
            st.warning(w)

    st.markdown(f"**Preview** — {len(preview_df)} row(s) found:")
    st.dataframe(preview_df[["Full Name", "Company Name", "Email", "Title", "Contact Type"]].head(20),
                 use_container_width=True, hide_index=True)

    # Count new vs duplicate
    existing = dfs.get("Contacts", pd.DataFrame())
    new_rows, skipped = _deduplicate(preview_df, existing)
    st.caption(f"{len(new_rows)} new · {skipped} duplicate(s) skipped")

    if new_rows:
        if st.button(f"✅ Import {len(new_rows)} contact(s)", use_container_width=True, key="ct_import_btn"):
            contacts = existing if not existing.empty else pd.DataFrame(columns=list(new_rows[0].keys()))
            next_id_num = _next_id_num(contacts)
            rows_to_add = []
            for i, r in enumerate(new_rows):
                r["Contact ID"] = f"CT-{next_id_num + i:03d}"
                rows_to_add.append(r)
            dfs["Contacts"] = pd.concat(
                [contacts, pd.DataFrame(rows_to_add)],
                ignore_index=True,
            )
            save_session(dfs)
            st.success(f"✅ {len(rows_to_add)} contact(s) imported.")
            st.rerun()
    else:
        st.info("No new contacts to import — all entries already exist in the directory.")


def _parse_contacts_excel(uploaded) -> tuple:
    """Read an uploaded Excel file and return (normalized_df, warnings)."""
    warnings = []
    try:
        xl = pd.ExcelFile(uploaded)
    except Exception as e:
        st.error(f"Could not read file: {e}")
        return None, []

    # Prefer a sheet named "Contacts", otherwise take the first sheet
    sheet = "Contacts" if "Contacts" in xl.sheet_names else xl.sheet_names[0]
    try:
        raw = xl.parse(sheet)
    except Exception as e:
        st.error(f"Could not parse sheet '{sheet}': {e}")
        return None, []

    raw.columns = [str(c).strip() for c in raw.columns]
    raw = raw.dropna(how="all").reset_index(drop=True)

    if raw.empty:
        st.warning("The sheet appears to be empty.")
        return None, []

    # Map columns using aliases
    col_map = _resolve_column_map(raw.columns)

    if "Full Name" not in col_map and "First Name" not in col_map:
        # Try to build Full Name from First Name + Last Name
        fn_col = next((c for c in raw.columns if "first" in c.lower()), None)
        ln_col = next((c for c in raw.columns if "last" in c.lower()), None)
        if fn_col and ln_col:
            raw["__full_name"] = (raw[fn_col].fillna("") + " " + raw[ln_col].fillna("")).str.strip()
            col_map["Full Name"] = "__full_name"
        else:
            st.error("Could not find a 'Full Name' column (or 'First Name' + 'Last Name'). "
                     "Please download the template for the expected format.")
            return None, []

    if "Company Name" not in col_map:
        warnings.append("No 'Company Name' column found — company will be left blank.")

    rows = []
    for _, r in raw.iterrows():
        name = str(r.get(col_map.get("Full Name", ""), "") or "").strip()
        if not name or name.lower() in ("nan", "none"):
            continue
        rows.append({
            "Contact ID":   "",
            "Company Name": str(r.get(col_map.get("Company Name", ""), "") or "").strip(),
            "Full Name":    name,
            "Title":        str(r.get(col_map.get("Title", ""), "") or "").strip(),
            "Department":   str(r.get(col_map.get("Department", ""), "") or "").strip(),
            "Email":        str(r.get(col_map.get("Email", ""), "") or "").strip(),
            "Phone":        str(r.get(col_map.get("Phone", ""), "") or "").strip(),
            "LinkedIn":     str(r.get(col_map.get("LinkedIn", ""), "") or "").strip(),
            "Contact Type": _normalise_type(str(r.get(col_map.get("Contact Type", ""), "") or "")),
            "Notes":        str(r.get(col_map.get("Notes", ""), "") or "").strip(),
            "Last Updated": date.today(),
        })

    if not rows:
        st.warning("No valid rows found (all rows were missing a Full Name).")
        return None, warnings

    return pd.DataFrame(rows), warnings


def _resolve_column_map(columns) -> dict:
    """Return {canonical_field: actual_column} based on alias matching."""
    lower_map = {c.lower().strip(): c for c in columns}
    result = {}
    for canonical, aliases in _COL_ALIASES.items():
        for alias in aliases:
            if alias in lower_map:
                result[canonical] = lower_map[alias]
                break
        # Also try exact match on the canonical name itself
        if canonical not in result and canonical.lower() in lower_map:
            result[canonical] = lower_map[canonical.lower()]
    return result


def _normalise_type(val: str) -> str:
    v = val.strip().title()
    return v if v in _CONTACT_TYPES else "Primary"


def _deduplicate(new_df: pd.DataFrame, existing: pd.DataFrame) -> tuple:
    """Return (list_of_new_row_dicts, skipped_count)."""
    if existing.empty or "Full Name" not in existing.columns:
        return new_df.to_dict("records"), 0

    existing_keys = set(
        (str(r.get("Full Name", "")).strip().lower(),
         str(r.get("Company Name", "")).strip().lower())
        for _, r in existing.iterrows()
    )
    new_rows = []
    skipped  = 0
    for _, r in new_df.iterrows():
        key = (str(r.get("Full Name", "")).strip().lower(),
               str(r.get("Company Name", "")).strip().lower())
        if key in existing_keys:
            skipped += 1
        else:
            existing_keys.add(key)
            new_rows.append(r.to_dict())
    return new_rows, skipped


def _build_template_excel() -> bytes:
    cols = ["Full Name", "Company Name", "Title", "Department",
            "Email", "Phone", "LinkedIn", "Contact Type", "Notes"]
    sample = pd.DataFrame([
        {
            "Full Name": "Ahmed Al-Rashid",
            "Company Name": "Example Capital",
            "Title": "Managing Director",
            "Department": "Investments",
            "Email": "ahmed@example.com",
            "Phone": "+966-50-000-0000",
            "LinkedIn": "https://linkedin.com/in/ahmed",
            "Contact Type": "Primary",
            "Notes": "Key decision maker",
        }
    ], columns=cols)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        sample.to_excel(w, sheet_name="Contacts", index=False)
    return buf.getvalue()


# ── Sync from Investor Master ─────────────────────────────────────────────────

def _sync_from_investor_master(dfs: dict, only_company: str = "") -> int:
    """
    Pull inline contact fields (Company Rep, Key Contact Name, Rep Email, etc.)
    from Investor Master into the Contacts directory.
    Returns the number of new contacts added.
    """
    investors = dfs.get("Investor Master", pd.DataFrame())
    if investors.empty or "Company Name" not in investors.columns:
        return 0

    contacts = dfs.get("Contacts", pd.DataFrame())
    existing_keys = set()
    if not contacts.empty and "Full Name" in contacts.columns:
        existing_keys = {
            (str(r.get("Full Name", "")).strip().lower(),
             str(r.get("Company Name", "")).strip().lower())
            for _, r in contacts.iterrows()
        }

    new_rows = []
    next_num = _next_id_num(contacts)

    for _, inv in investors.iterrows():
        company = str(inv.get("Company Name", "") or "").strip()
        if not company:
            continue
        if only_company and company != only_company:
            continue

        # Field groups to import
        candidates = [
            {
                "Full Name":    str(inv.get("Company Rep", "") or "").strip(),
                "Title":        str(inv.get("Rep Position", "") or "").strip(),
                "Email":        str(inv.get("Rep Email", "") or "").strip(),
                "Phone":        str(inv.get("Rep Phone", "") or "").strip(),
                "Contact Type": "Primary",
            },
            {
                "Full Name":    str(inv.get("Key Contact Name", "") or "").strip(),
                "Title":        str(inv.get("Key Contact Title", "") or "").strip(),
                "Email":        str(inv.get("Key Contact Email", "") or "").strip(),
                "Phone":        str(inv.get("Key Contact Phone", "") or "").strip(),
                "Contact Type": "Secondary",
            },
        ]

        for c in candidates:
            name = c["Full Name"]
            if not name or name.lower() in ("nan", "none", "n/a", "—"):
                continue
            key = (name.lower(), company.lower())
            if key in existing_keys:
                continue
            existing_keys.add(key)
            new_rows.append({
                "Contact ID":   f"CT-{next_num:03d}",
                "Company Name": company,
                "Full Name":    name,
                "Title":        c["Title"],
                "Department":   "",
                "Email":        c["Email"] if c["Email"].lower() not in ("nan", "none") else "",
                "Phone":        c["Phone"] if c["Phone"].lower() not in ("nan", "none") else "",
                "LinkedIn":     "",
                "Contact Type": c["Contact Type"],
                "Notes":        "Imported from Investor Master",
                "Last Updated": date.today(),
            })
            next_num += 1

    if not new_rows:
        return 0

    base = contacts if not contacts.empty else pd.DataFrame(columns=list(new_rows[0].keys()))
    dfs["Contacts"] = pd.concat([base, pd.DataFrame(new_rows)], ignore_index=True)
    return len(new_rows)


# ── Card renderer ─────────────────────────────────────────────────────────────

def _render_contact_card(row, dfs: dict):
    cid     = str(row.get("Contact ID",   "") or "")
    company = str(row.get("Company Name", "") or "")
    name    = str(row.get("Full Name",    "") or "")
    title   = str(row.get("Title",        "") or "")
    dept    = str(row.get("Department",   "") or "")
    email   = str(row.get("Email",        "") or "")
    phone   = str(row.get("Phone",        "") or "")
    linkedin= str(row.get("LinkedIn",     "") or "")
    ctype   = str(row.get("Contact Type", "Primary") or "Primary")
    notes   = str(row.get("Notes",        "") or "")

    color, bg = _TYPE_COLOR.get(ctype, (_GREEN, "#f9fafb"))
    initials  = "".join(p[0].upper() for p in name.split()[:2]) if name else "?"

    def _clean(v):
        return v and v not in ("nan", "—", "none", "None")

    email_html = (
        f'<div style="font-size:11px;margin-top:3px;">'
        f'<a href="mailto:{email}" style="color:#1D4ED8;text-decoration:none;">✉ {email}</a></div>'
    ) if _clean(email) else ""

    phone_html = (
        f'<div style="font-size:11px;color:#6B7280;margin-top:2px;">📞 {phone}</div>'
    ) if _clean(phone) else ""

    li_html = ""
    if _clean(linkedin):
        href = linkedin if linkedin.startswith("http") else f"https://linkedin.com/in/{linkedin}"
        li_html = (f'<div style="font-size:10px;margin-top:2px;">'
                   f'<a href="{href}" target="_blank" style="color:#0A66C2;text-decoration:none;">🔗 LinkedIn</a></div>')

    dept_html  = (f'<span style="font-size:10px;color:#6B7280;"> · {dept}</span>') if _clean(dept) else ""
    notes_html = (
        f'<div style="font-size:10px;color:#6B7280;margin-top:5px;'
        f'border-top:1px solid #f0f0f0;padding-top:5px;">{notes[:100]}</div>'
    ) if _clean(notes) else ""

    card_html = (
        f'<div style="border:1px solid #E5E7EB;border-radius:10px;padding:12px;'
        f'background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.05);margin-bottom:4px;">'
        f'<div style="display:flex;align-items:flex-start;gap:10px;">'
        f'<div style="width:38px;height:38px;border-radius:50%;background:{bg};'
        f'border:2px solid {color};display:flex;align-items:center;justify-content:center;'
        f'font-size:13px;font-weight:700;color:{color};flex-shrink:0;">{initials}</div>'
        f'<div style="flex:1;min-width:0;">'
        f'<div style="font-size:13px;font-weight:700;color:#1F2937;">{name}</div>'
        f'<div style="font-size:11px;color:#6B7280;margin-top:1px;">{title}{dept_html}</div>'
        f'<div style="font-size:10px;color:{color};font-weight:600;margin-top:2px;">{ctype}</div>'
        f'<div style="font-size:10px;color:#9CA3AF;">{company}</div>'
        f'{email_html}{phone_html}{li_html}{notes_html}'
        f'</div>'
        f'</div></div>'
    )
    st.markdown(card_html, unsafe_allow_html=True)


# ── Add contact form ──────────────────────────────────────────────────────────

def _add_contact_form(dfs: dict, investors: pd.DataFrame, default_company: str = ""):
    contacts  = dfs.get("Contacts", pd.DataFrame())

    with st.form(f"add_contact_form_{default_company}", clear_on_submit=True):
        co_options = sorted(investors["Company Name"].dropna().unique().tolist()) if (
            not investors.empty and "Company Name" in investors.columns
        ) else []

        r1c1, r1c2 = st.columns(2)
        if default_company and default_company in co_options:
            company = r1c1.selectbox("Company", co_options,
                                      index=co_options.index(default_company), key=f"ct_co_{default_company}")
        else:
            company = r1c1.selectbox("Company", ["— select —"] + co_options, key=f"ct_co_new_{default_company}")
        name = r1c2.text_input("Full Name *")

        r2c1, r2c2, r2c3 = st.columns(3)
        title   = r2c1.text_input("Title / Position")
        dept    = r2c2.text_input("Department")
        ctype   = r2c3.selectbox("Contact Type", _CONTACT_TYPES, key=f"ct_type_form_{default_company}")

        r3c1, r3c2, r3c3 = st.columns(3)
        email   = r3c1.text_input("Email")
        phone   = r3c2.text_input("Phone")
        linkedin= r3c3.text_input("LinkedIn URL or username")

        notes = st.text_area("Notes", height=60)

        if st.form_submit_button("Add Contact", use_container_width=True):
            if not name:
                st.warning("Full Name is required.")
                return
            if not company or company == "— select —":
                st.warning("Please select a company.")
                return

            existing = contacts if not contacts.empty else pd.DataFrame(columns=["Contact ID"])
            cid = _next_contact_id(existing)
            new_row = {
                "Contact ID":   cid,
                "Company Name": company,
                "Full Name":    name,
                "Title":        title,
                "Department":   dept,
                "Email":        email,
                "Phone":        phone,
                "LinkedIn":     linkedin,
                "Contact Type": ctype,
                "Notes":        notes,
                "Last Updated": date.today(),
            }
            dfs["Contacts"] = pd.concat(
                [existing, pd.DataFrame([new_row])],
                ignore_index=True,
            )
            save_session(dfs)
            st.success(f"✅ Contact **{name}** added for {company}.")
            st.rerun()


# ── ID helpers ────────────────────────────────────────────────────────────────

def _next_id_num(contacts: pd.DataFrame) -> int:
    if contacts.empty or "Contact ID" not in contacts.columns:
        return 1
    nums = []
    for v in contacts["Contact ID"].dropna():
        try:
            nums.append(int(str(v).split("-")[-1]))
        except ValueError:
            pass
    return (max(nums) + 1) if nums else 1


def _next_contact_id(contacts: pd.DataFrame) -> str:
    return f"CT-{_next_id_num(contacts):03d}"
