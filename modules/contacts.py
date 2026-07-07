
# Contact Management — multiple contacts per investor company.

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


def render(dfs: dict, lang: str):
    contacts  = dfs.get("Contacts",       pd.DataFrame())
    investors = dfs.get("Investor Master", pd.DataFrame())

    st.markdown(
        f"<h2 style='color:{_GREEN};margin-bottom:2px;'>Contact Directory</h2>"
        f"<p style='color:#6b7280;font-size:13px;margin-top:0;'>"
        f"Manage multiple contacts per investor company</p>",
        unsafe_allow_html=True,
    )

    with st.expander("+ Add New Contact", expanded=False):
        _add_contact_form(dfs, investors)

    if contacts.empty:
        st.info("No contacts yet. Add one above or open an investor profile.")
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
        mask = view[["Full Name", "Company Name", "Email", "Title"]].apply(
            lambda col: col.fillna("").astype(str).str.lower().str.contains(q.lower(), regex=False)
        ).any(axis=1)
        view = view[mask]
    if sel_co != "All companies":
        view = view[view.get("Company Name", pd.Series()) == sel_co]
    if sel_type != "All types":
        view = view[view.get("Contact Type", pd.Series()) == sel_type]
    view = view.reset_index(drop=True)

    st.caption(f"{len(view)} contact{'s' if len(view) != 1 else ''}")
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
        st.caption("No contacts linked yet — add one above.")
        # Also show any inline contacts from Investor Master for migration hint
        inv = dfs.get("Investor Master", pd.DataFrame())
        if not inv.empty and "Company Name" in inv.columns:
            row = inv[inv["Company Name"] == company]
            if not row.empty:
                r = row.iloc[0]
                rep   = str(r.get("Company Rep",  "") or "").strip()
                kc    = str(r.get("Key Contact Name", "") or "").strip()
                if rep or kc:
                    st.markdown(
                        f'<div style="background:#fffbeb;border:1px solid #fde68a;'
                        f'border-radius:6px;padding:8px 12px;font-size:12px;color:#92400e;">'
                        f'Legacy contact data found in Investor Master — '
                        f'{("Rep: " + rep) if rep else ""}'
                        f'{(" | Key Contact: " + kc) if kc else ""}'
                        f'. Use the form above to migrate to the Contacts directory.</div>',
                        unsafe_allow_html=True,
                    )
        return

    cols_per_row = 2
    for start in range(0, len(co_contacts), cols_per_row):
        chunk = co_contacts.iloc[start:start + cols_per_row]
        cols  = st.columns(cols_per_row)
        for ci, (_, row) in enumerate(chunk.iterrows()):
            with cols[ci]:
                _render_contact_card(row, dfs)


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

    email_html = ""
    if email and email not in ("nan", "—"):
        email_html = (
            f'<div style="font-size:11px;margin-top:3px;">'
            f'<a href="mailto:{email}" style="color:#1D4ED8;text-decoration:none;">✉ {email}</a></div>'
        )
    phone_html = ""
    if phone and phone not in ("nan", "—"):
        phone_html = f'<div style="font-size:11px;color:#6B7280;margin-top:2px;">📞 {phone}</div>'
    li_html = ""
    if linkedin and linkedin not in ("nan", "—"):
        href = linkedin if linkedin.startswith("http") else f"https://linkedin.com/in/{linkedin}"
        li_html = f'<div style="font-size:10px;margin-top:2px;"><a href="{href}" target="_blank" style="color:#0A66C2;text-decoration:none;">🔗 LinkedIn</a></div>'
    dept_html = ""
    if dept and dept not in ("nan", "—"):
        dept_html = f'<span style="font-size:10px;color:#6B7280;"> · {dept}</span>'
    notes_html = ""
    if notes and notes not in ("nan", "—"):
        notes_html = (
            f'<div style="font-size:10px;color:#6B7280;margin-top:5px;'
            f'border-top:1px solid #f0f0f0;padding-top:5px;">{notes[:100]}</div>'
        )

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
        f'<div style="font-size:10px;color:{color};font-weight:600;margin-top:2px;">'
        f'{ctype}</div>'
        f'<div style="font-size:10px;color:#9CA3AF;">{company}</div>'
        f'{email_html}{phone_html}{li_html}{notes_html}'
        f'</div>'
        f'</div></div>'
    )
    st.markdown(card_html, unsafe_allow_html=True)


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


def _next_contact_id(contacts: pd.DataFrame) -> str:
    if contacts.empty or "Contact ID" not in contacts.columns:
        return "CT-001"
    nums = []
    for v in contacts["Contact ID"].dropna():
        try:
            nums.append(int(str(v).split("-")[-1]))
        except ValueError:
            pass
    return f"CT-{(max(nums) + 1 if nums else 1):03d}"
