
# Investor Master — card-based view with live logos, Wikidata enrichment, and today's news.

import hashlib
import html as html_mod
import json
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import numpy as np
import pandas as pd
import streamlit as st
from datetime import date, datetime

from config.settings import (
    SECTORS, COUNTRIES, JOURNEY_STAGES, ESCALATION_FLAGS,
    MISA_GREEN, MISA_GOLD, STATUS_COLORS,
)
from config.translations import t
from modules.persistence import save_session


_GREEN = "#1B5C3F"
_GOLD  = "#C9974A"
_RED   = "#DC2626"

_STAGE_COLOR = {
    "Awareness":             "#6B7280",
    "Initial Contact":       "#0891B2",
    "Engagement":            "#1D4ED8",
    "Opportunity Matching":  "#D97706",
    "Active Negotiation":    "#B45309",
    "Committed":             "#059669",
    "Post-Investment":       _GREEN,
}


# ── Logo & data helpers (cached) ──────────────────────────────────────────────

def _guess_domain(company: str) -> str:
    name = company.lower()
    for w in [" group", " capital", " asset management", " management",
              " holdings", " limited", " ltd", " inc", " corp", " plc",
              " partners", " advisors", " international", " global"]:
        name = name.replace(w, "")
    return name.strip().replace(" ", "").replace(".", "").replace("-", "") + ".com"


def _domain_from_website(website: str) -> str:
    """Extract bare domain from a website URL entered in the Excel tracker."""
    url = website.strip().lower()
    if not url or "." not in url:
        return ""
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    try:
        parsed = urllib.parse.urlparse(url)
        domain = (parsed.netloc or url).replace("www.", "")
        return domain.split("/")[0]
    except Exception:
        return url.replace("www.", "").split("/")[0]


def _logo_url(company: str) -> str:
    return f"https://logo.clearbit.com/{_guess_domain(company)}"


@st.cache_data(ttl=86400, show_spinner=False)
def _logo_html(company: str, size: int = 44, website: str = "") -> str:
    """
    Return an <img> tag with base64-embedded logo, or a coloured initials
    circle if the logo can't be fetched. Uses the website URL when provided
    for a more accurate domain lookup.
    """
    import base64
    color    = _avatar_color(company)
    initials = _initials(company)
    radius   = "8px" if size <= 44 else "10px"

    # Prefer domain from the actual website URL; fall back to guessing
    domain = _domain_from_website(website) if website else ""
    if not domain:
        domain = _guess_domain(company)
    for url in [
        f"https://www.google.com/s2/favicons?domain={domain}&sz=64",
        f"https://logo.clearbit.com/{domain}",
    ]:
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=3) as resp:
                data = resp.read()
            if len(data) < 200:   # tiny placeholder / error image
                continue
            ext = "png"
            b64 = base64.b64encode(data).decode()
            return (
                f'<img src="data:image/{ext};base64,{b64}" '
                f'style="width:{size}px;height:{size}px;object-fit:contain;'
                f'border-radius:{radius};border:1px solid #f0f0f0;background:#fff;" />'
            )
        except Exception:
            continue

    # Fallback: coloured initials circle
    font = max(10, size // 3)
    return (
        f'<div style="width:{size}px;height:{size}px;border-radius:{radius};'
        f'background:{color};display:flex;align-items:center;justify-content:center;'
        f'font-size:{font}px;font-weight:700;color:#fff;flex-shrink:0;">{initials}</div>'
    )


@st.cache_data(ttl=86400, show_spinner=False)
def _fetch_company_data(company: str) -> dict:
    """
    Enrich a company with data from Wikipedia + Wikidata.
    Returns dict: {employees_global, hq_country, industry}
    All values may be None if not found or on network error.
    """
    result = {"employees_global": None, "hq_country": None, "industry": None}
    try:
        q = urllib.parse.quote(company + " company")
        search_url = (
            f"https://en.wikipedia.org/w/api.php?action=query&list=search"
            f"&srsearch={q}&format=json&srlimit=1"
        )
        req = urllib.request.Request(search_url, headers={"User-Agent": "MinsterCRM/1.0"})
        with urllib.request.urlopen(req, timeout=6) as resp:
            search_data = json.loads(resp.read())

        hits = search_data.get("query", {}).get("search", [])
        if not hits:
            return result

        page_title = urllib.parse.quote(hits[0]["title"])
        sum_url = f"https://en.wikipedia.org/api/rest_v1/page/summary/{page_title}"
        req2 = urllib.request.Request(sum_url, headers={"User-Agent": "MinsterCRM/1.0"})
        with urllib.request.urlopen(req2, timeout=6) as resp:
            summary = json.loads(resp.read())

        wd_id = summary.get("wikibase_item")
        if not wd_id:
            return result

        # Fetch Wikidata claims: P1128=employees, P17=country, P452=industry, P159=HQ location
        wd_url = (
            f"https://www.wikidata.org/w/api.php?action=wbgetclaims"
            f"&entity={wd_id}&property=P1128|P17|P452|P159&format=json"
        )
        req3 = urllib.request.Request(wd_url, headers={"User-Agent": "MinsterCRM/1.0"})
        with urllib.request.urlopen(req3, timeout=8) as resp:
            claims_data = json.loads(resp.read())

        claims = claims_data.get("claims", {})
        ids_to_resolve: list[str] = []
        country_id = None
        industry_id = None

        # P1128 — number of employees (numeric, no label needed)
        for claim in claims.get("P1128", []):
            dv = claim.get("mainsnak", {}).get("datavalue", {})
            if dv.get("type") == "quantity":
                amt = float(dv["value"]["amount"].lstrip("+"))
                if amt >= 1_000_000:
                    result["employees_global"] = f"{amt/1_000_000:.1f}M"
                elif amt >= 1_000:
                    result["employees_global"] = f"{int(amt/1_000)}K"
                else:
                    result["employees_global"] = str(int(amt))
                break

        # P17 — country of origin
        for claim in claims.get("P17", []):
            dv = claim.get("mainsnak", {}).get("datavalue", {})
            if dv.get("type") == "wikibase-entityid":
                country_id = dv["value"]["id"]
                ids_to_resolve.append(country_id)
                break

        # P452 — industry
        for claim in claims.get("P452", []):
            dv = claim.get("mainsnak", {}).get("datavalue", {})
            if dv.get("type") == "wikibase-entityid":
                industry_id = dv["value"]["id"]
                ids_to_resolve.append(industry_id)
                break

        # Batch-resolve entity labels in one request
        if ids_to_resolve:
            ids_str = "|".join(ids_to_resolve)
            label_url = (
                f"https://www.wikidata.org/w/api.php?action=wbgetlabels"
                f"&ids={ids_str}&languages=en&format=json"
            )
            req4 = urllib.request.Request(label_url, headers={"User-Agent": "MinsterCRM/1.0"})
            with urllib.request.urlopen(req4, timeout=6) as resp:
                label_data = json.loads(resp.read())
            entities = label_data.get("entities", {})
            if country_id and country_id in entities:
                result["hq_country"] = (
                    entities[country_id].get("labels", {}).get("en", {}).get("value")
                )
            if industry_id and industry_id in entities:
                result["industry"] = (
                    entities[industry_id].get("labels", {}).get("en", {}).get("value")
                )
    except Exception:
        pass
    return result


@st.cache_data(ttl=3600, show_spinner=False)
def _fetch_news(company: str) -> list[dict]:
    """Fetch up to 3 recent news items from Google News RSS."""
    try:
        q   = urllib.parse.quote(company)
        url = f"https://news.google.com/rss/search?q={q}&hl=en-US&gl=US&ceid=US:en"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=6) as resp:
            data = resp.read()
        root  = ET.fromstring(data)
        items = []
        for item in root.findall(".//item")[:3]:
            title = item.findtext("title", "").strip()
            link  = item.findtext("link",  "").strip()
            pub   = item.findtext("pubDate", "").strip()
            if " - " in title:
                title = title.rsplit(" - ", 1)[0].strip()
            elif " — " in title:
                title = title.rsplit(" — ", 1)[0].strip()
            if title:
                items.append({"title": title, "link": link, "pub": pub[:16]})
        return items
    except Exception:
        return []


def _initials(name: str) -> str:
    parts = name.split()
    if len(parts) >= 2:
        return (parts[0][0] + parts[-1][0]).upper()
    return name[:2].upper() if name else "?"


def _avatar_color(name: str) -> str:
    h = int(hashlib.md5(name.encode()).hexdigest()[:6], 16)
    r = max(40, min((h >> 16) & 0xFF, 150))
    g = max(40, min((h >> 8)  & 0xFF, 150))
    b = max(40, min(h & 0xFF,          150))
    return f"#{r:02X}{g:02X}{b:02X}"


def _fmt_sar(val) -> str:
    try:
        v = float(val)
        if v >= 1e9:  return f"SAR {v/1e9:.1f}B"
        if v >= 1e6:  return f"SAR {v/1e6:.0f}M"
        if v > 0:     return f"SAR {v:,.0f}"
    except Exception:
        pass
    return ""


def _blank(val) -> bool:
    return not val or str(val).strip() in ("", "—", "nan", "None", "NaT")


# ── Main render ───────────────────────────────────────────────────────────────

def render(dfs: dict, lang: str):
    investors = dfs.get("Investor Master", pd.DataFrame())

    st.markdown(
        f"<h2 style='color:{_GREEN};margin-bottom:2px;'>Investor Portfolio</h2>"
        f"<p style='color:#6b7280;font-size:13px;margin-top:0;'>"
        f"Live view of all tracked investors — logos, enriched data and today's news</p>",
        unsafe_allow_html=True,
    )

    with st.expander("+ Add New Investor", expanded=False):
        _add_investor_form(dfs, lang)

    if investors.empty:
        st.info("No investor data loaded.")
        return

    # ── Filters ───────────────────────────────────────────────────────────────
    fc1, fc2, fc3, fc4 = st.columns([2.5, 1.5, 1.5, 1.5])
    with fc1:
        q = st.text_input("Search", placeholder="Search investor…",
                          key="inv_search", label_visibility="collapsed")
    with fc2:
        sectors   = ["All sectors"]  + sorted(s for s in investors.get("Sector",  pd.Series()).dropna().unique() if s)
        sel_sec   = st.selectbox("Sector",  sectors,  key="inv_sec",  label_visibility="collapsed")
    with fc3:
        countries = ["All countries"] + sorted(c for c in investors.get("Country", pd.Series()).dropna().unique() if c)
        sel_cty   = st.selectbox("Country", countries, key="inv_cty", label_visibility="collapsed")
    with fc4:
        stages   = ["All stages"] + JOURNEY_STAGES
        sel_stg  = st.selectbox("Stage", stages, key="inv_stg", label_visibility="collapsed")

    view = investors.copy()
    if q:
        view = view[view["Company Name"].fillna("").str.contains(q, case=False, na=False)]
    if sel_sec != "All sectors":
        view = view[view.get("Sector",        pd.Series()) == sel_sec]
    if sel_cty != "All countries":
        view = view[view.get("Country",       pd.Series()) == sel_cty]
    if sel_stg != "All stages":
        view = view[view.get("Journey Stage", pd.Series()) == sel_stg]
    view = view.reset_index(drop=True)

    st.caption(f"{len(view)} investor{'s' if len(view)!=1 else ''}")
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

    if view.empty:
        st.warning("No investors match the current filters.")
        return

    # ── Card grid — 3 per row ─────────────────────────────────────────────────
    cols_per_row = 3
    for chunk_start in range(0, len(view), cols_per_row):
        chunk = view.iloc[chunk_start : chunk_start + cols_per_row]
        cols  = st.columns(cols_per_row)
        for ci, (_, row) in enumerate(chunk.iterrows()):
            with cols[ci]:
                _render_investor_card(row)

    # ── Profile drill-down ────────────────────────────────────────────────────
    st.markdown("---")
    if "Company Name" in view.columns:
        selected = st.selectbox(
            "Open investor profile",
            ["— select to open profile —"] + sorted(view["Company Name"].dropna().unique().tolist()),
            key="inv_profile_select",
        )
        if selected != "— select to open profile —":
            _render_investor_profile(selected, dfs, lang)


# ── Investor card ─────────────────────────────────────────────────────────────

def _render_investor_card(row):
    company  = str(row.get("Company Name", "?"))
    sector   = str(row.get("Sector",   "") or "")
    country  = str(row.get("Country",  "") or "")
    stage    = str(row.get("Journey Stage", "—") or "—")
    rm       = str(row.get("Relationship Manager", "") or "")
    am       = str(row.get("Account Manager", "") or "")
    est_val  = _fmt_sar(row.get("Est. Investment Value (SAR)"))
    priority = row.get("Strategic Priority Score")
    next_mtg = row.get("Next Meeting Date")
    contact_name  = str(row.get("Key Contact Name",  "") or "")
    contact_title = str(row.get("Key Contact Title", "") or "")
    size_global   = str(row.get("Company Size (Global)", "") or "")
    size_ksa      = str(row.get("Company Size (KSA)",    "") or "")
    website   = str(row.get("Website",      "") or "").strip()
    rep_name  = str(row.get("Company Rep",  "") or "").strip()
    rep_pos   = str(row.get("Rep Position", "") or "").strip()
    rep_email = str(row.get("Rep Email",    "") or "").strip()
    rep_phone = str(row.get("Rep Phone",    "") or "").strip()

    color    = _avatar_color(company)
    initials = _initials(company)
    logo     = _logo_url(company)
    sc_color = _STAGE_COLOR.get(stage, "#6B7280")

    # Enrich missing sector/country/size from Wikidata
    enriched = _fetch_company_data(company)
    display_sector  = sector  if not _blank(sector)  else (enriched.get("industry")   or "—")
    display_country = country if not _blank(country) else (enriched.get("hq_country") or "—")
    if _blank(size_global) and enriched.get("employees_global"):
        size_global = enriched["employees_global"] + " employees"

    sector_note  = " *" if _blank(sector)  and not _blank(display_sector)  else ""
    country_note = " *" if _blank(country) and not _blank(display_country) else ""

    priority_star = ""
    try:
        if priority and float(priority) >= 4:
            priority_star = f'<span style="color:{_GOLD};font-size:12px;margin-left:4px;">★</span>'
    except Exception:
        pass

    next_str = ""
    if next_mtg and str(next_mtg) not in ("NaT", "None", "nan", ""):
        try:
            d = pd.to_datetime(next_mtg).date()
            delta = (d - date.today()).days
            if delta < 0:
                next_str = f'<span style="color:{_RED};font-size:10px;">⚠ Next mtg: {d.strftime("%d %b")} (overdue)</span>'
            elif delta <= 7:
                next_str = f'<span style="color:#D97706;font-size:10px;">Next mtg: {d.strftime("%d %b")} (this week)</span>'
            else:
                next_str = f'<span style="color:#6B7280;font-size:10px;">Next mtg: {d.strftime("%d %b %Y")}</span>'
        except Exception:
            pass

    contact_html = ""
    if contact_name:
        contact_html = (
            f'<div style="display:flex;align-items:center;gap:6px;margin:8px 0 4px 0;'
            f'padding-top:8px;border-top:1px solid #f0f0f0;">'
            f'<div style="width:24px;height:24px;border-radius:50%;background:#f3f4f6;'
            f'display:flex;align-items:center;justify-content:center;'
            f'font-size:9px;font-weight:700;color:#374151;flex-shrink:0;">'
            f'{_initials(contact_name)}</div>'
            f'<div><div style="font-size:11px;font-weight:600;color:#111827;">{contact_name}</div>'
            f'{"<div style=font-size:9px;color:#6b7280;>" + contact_title + "</div>" if contact_title else ""}'
            f'</div></div>'
        )

    size_html = ""
    if size_global or size_ksa:
        parts = []
        if size_global:
            parts.append(f'🌐 {size_global}')
        if size_ksa:
            parts.append(f'🇸🇦 {size_ksa}')
        size_html = (
            f'<div style="margin-top:5px;display:flex;gap:8px;flex-wrap:wrap;">'
            + "".join(
                f'<span style="font-size:9px;color:#374151;background:#f3f4f6;'
                f'padding:2px 7px;border-radius:5px;">{p}</span>'
                for p in parts
            )
            + '</div>'
        )

    rm_am = ""
    parts_rm = [p for p in [rm, am] if p and p not in ("—", "TBD", "nan")]
    if parts_rm:
        rm_am = f'<div style="font-size:10px;color:#9CA3AF;margin-top:4px;">{" / ".join(parts_rm)}</div>'

    # Company Rep block
    rep_html = ""
    if rep_name and rep_name not in ("—", "nan"):
        _rep_detail = ""
        if rep_pos:
            _rep_detail += f'<div style="font-size:9px;color:#6B7280;">{rep_pos}</div>'
        if rep_email and rep_email not in ("—", "nan"):
            _rep_detail += (
                f'<div style="font-size:9px;color:#6B7280;">'
                f'<a href="mailto:{rep_email}" style="color:#1D4ED8;text-decoration:none;">'
                f'{rep_email}</a></div>'
            )
        if rep_phone and rep_phone not in ("—", "nan"):
            _rep_detail += f'<div style="font-size:9px;color:#6B7280;">&#x1F4DE; {rep_phone}</div>'
        rep_html = (
            f'<div style="display:flex;align-items:flex-start;gap:6px;margin:6px 0 2px 0;'
            f'padding:5px 8px;background:#f8fafc;border-radius:6px;border:1px solid #f1f5f9;">'
            f'<div style="width:22px;height:22px;border-radius:50%;background:#e8f5ee;'
            f'display:flex;align-items:center;justify-content:center;'
            f'font-size:8px;font-weight:700;color:{_GREEN};flex-shrink:0;margin-top:1px;">'
            f'{_initials(rep_name)}</div>'
            f'<div><div style="font-size:10px;font-weight:600;color:#1F2937;">{rep_name}</div>'
            f'{_rep_detail}'
            f'</div></div>'
        )

    # Website link
    website_html = ""
    _ws = website.lower()
    if website and "." in website and _ws not in ("vvvvv", "n/a", "—", "nan"):
        href = website if website.startswith("http") else f"https://{website}"
        disp = website.replace("https://", "").replace("http://", "").rstrip("/")[:38]
        website_html = (
            f'<div style="margin-top:3px;">'
            f'<a href="{href}" target="_blank" '
            f'style="color:#1D4ED8;font-size:9px;text-decoration:none;">'
            f'&#x1F517; {disp}</a></div>'
        )

    news_items = _fetch_news(company)
    news_html  = ""
    if news_items:
        news_html = (
            '<div style="margin-top:10px;padding-top:8px;border-top:1px solid #f0f0f0;">'
            '<div style="font-size:9px;font-weight:700;color:#9CA3AF;letter-spacing:.5px;'
            'text-transform:uppercase;margin-bottom:5px;">Latest News</div>'
        )
        for n in news_items[:2]:
            title = html_mod.escape(n["title"])[:90]
            pub   = n.get("pub", "")[:11]
            link  = n.get("link", "#")
            news_html += (
                f'<div style="margin-bottom:5px;">'
                f'<a href="{link}" target="_blank" style="color:#1D4ED8;font-size:10px;'
                f'font-weight:500;text-decoration:none;line-height:1.3;display:block;">{title}</a>'
                f'<span style="color:#9CA3AF;font-size:9px;">{pub}</span>'
                f'</div>'
            )
        news_html += '</div>'
    else:
        news_html = '<div style="margin-top:8px;font-size:10px;color:#D1D5DB;">No recent news found</div>'

    card = (
        f'<div style="border:1px solid #E5E7EB;border-radius:12px;padding:14px;'
        f'background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.06);margin-bottom:4px;">'

        # Logo + name row
        f'<div style="display:flex;align-items:flex-start;gap:10px;margin-bottom:10px;">'
        f'<div style="flex-shrink:0;">'
        f'{_logo_html(company, 44, website)}'
        f'</div>'
        f'<div style="flex:1;min-width:0;">'
        f'<div style="font-size:14px;font-weight:700;color:{_GREEN};line-height:1.2;">'
        f'{company}{priority_star}</div>'
        f'<div style="display:flex;gap:4px;flex-wrap:wrap;margin-top:4px;">'
        f'<span style="background:#f0fdf4;color:{_GREEN};padding:1px 7px;border-radius:6px;'
        f'font-size:10px;" title="{"Enriched from web" if sector_note else ""}">'
        f'{display_sector}{sector_note}</span>'
        f'<span style="background:#f8fafc;color:#475569;padding:1px 7px;border-radius:6px;'
        f'font-size:10px;" title="{"Enriched from web" if country_note else ""}">'
        f'🌍 {display_country}{country_note}</span>'
        f'</div></div>'
        f'</div>'

        # Stage bar
        f'<div style="background:#f3f4f6;border-radius:4px;height:4px;margin-bottom:4px;overflow:hidden;">'
        f'<div style="background:{sc_color};height:100%;width:100%;border-radius:4px;opacity:.7;"></div>'
        f'</div>'
        f'<div style="font-size:10px;color:{sc_color};font-weight:600;margin-bottom:4px;">{stage}</div>'

        # Company size
        f'{size_html}'

        # Contact
        f'{contact_html}'

        # Value + next meeting
        f'<div style="display:flex;justify-content:space-between;align-items:center;margin-top:6px;">'
        f'{"<span style=font-size:12px;font-weight:700;color:" + _GOLD + ";>" + est_val + "</span>" if est_val else "<span></span>"}'
        f'{next_str}'
        f'</div>'

        # AM/RM
        f'{rm_am}'

        # Company Rep + website
        f'{rep_html}'
        f'{website_html}'

        # News
        f'{news_html}'
        f'</div>'
    )

    st.markdown(card, unsafe_allow_html=True)
    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)


# ── Add investor form ─────────────────────────────────────────────────────────

def _add_investor_form(dfs: dict, lang: str):
    with st.form("add_investor_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        company = c1.text_input("Company Name")
        country = c2.selectbox("Country", COUNTRIES)

        c3, c4 = st.columns(2)
        sector  = c3.selectbox("Sector", SECTORS)
        stage   = c4.selectbox("Journey Stage", JOURNEY_STAGES)

        c5, c6 = st.columns(2)
        rm = c5.text_input("Relationship Manager")
        am = c6.text_input("Account Manager")

        c7, c8 = st.columns(2)
        contact_name  = c7.text_input("Key Contact Name")
        contact_title = c8.text_input("Key Contact Title / Position")

        c_web, c_email, c_phone = st.columns(3)
        website     = c_web.text_input("Company Website", placeholder="e.g. blackrock.com")
        rep_email   = c_email.text_input("Contact Email", placeholder="e.g. name@company.com")
        rep_phone   = c_phone.text_input("Contact Phone", placeholder="e.g. +966 11 000 0000")

        c9, c10 = st.columns(2)
        size_global = c9.text_input("Company Size (Global)", placeholder="e.g. 50,000 employees")
        size_ksa    = c10.text_input("Company Size (KSA)",    placeholder="e.g. 2,000 employees")

        c11, c12 = st.columns(2)
        est_val  = c11.number_input("Est. Investment Value (SAR)", min_value=0.0, step=1_000_000.0)
        next_mtg = c12.date_input("Next Meeting Date", value=None)

        notes = st.text_area("Notes")

        if st.form_submit_button("Add Investor", use_container_width=True):
            if not company:
                st.warning("Company name is required.")
                return
            investors = dfs.get("Investor Master", pd.DataFrame())
            new_id    = _next_id(investors, "Investor ID", "INV")
            new_row   = {
                "Investor ID":                new_id,
                "Company Name":               company,
                "Country":                    country,
                "Sector":                     sector,
                "Relationship Manager":       rm,
                "Account Manager":            am or "TBD",
                "Journey Stage":              stage,
                "Key Contact Name":           contact_name,
                "Key Contact Title":          contact_title,
                "Website":                    website,
                "Rep Email":                  rep_email,
                "Rep Phone":                  rep_phone,
                "Company Size (Global)":      size_global,
                "Company Size (KSA)":         size_ksa,
                "Est. Investment Value (SAR)": est_val if est_val > 0 else None,
                "Next Meeting Date":           next_mtg,
                "Last Updated":               np.datetime64(datetime.now()),
                "Escalation Flag":            "None",
                "Notes":                      notes,
            }
            dfs["Investor Master"] = pd.concat(
                [investors, pd.DataFrame([new_row])], ignore_index=True
            )
            save_session(dfs)
            st.success(f"{company} added ({new_id})")
            st.rerun()


# ── Investor profile (detail view) ────────────────────────────────────────────

def _render_investor_profile(company: str, dfs: dict, lang: str):
    investors     = dfs.get("Investor Master",      pd.DataFrame())
    actions       = dfs.get("Action Items",         pd.DataFrame())
    meetings      = dfs.get("Meeting Log",          pd.DataFrame())
    opportunities = dfs.get("Opportunity Pipeline", pd.DataFrame())
    tasks         = dfs.get("RM Tasks",             pd.DataFrame())
    deals         = dfs.get("Deal Progress",        pd.DataFrame())

    row = investors[investors["Company Name"] == company].iloc[0]

    stage    = str(row.get("Journey Stage", "—"))
    country  = str(row.get("Country", "—"))
    sector   = str(row.get("Sector",  "—"))
    rm       = str(row.get("Relationship Manager", "—"))
    am       = str(row.get("Account Manager", "TBD"))
    contact  = str(row.get("Key Contact Name",  "") or "")
    ctitle   = str(row.get("Key Contact Title", "") or "")
    size_global = str(row.get("Company Size (Global)", "") or "")
    size_ksa    = str(row.get("Company Size (KSA)",    "") or "")
    website   = str(row.get("Website",      "") or "").strip()
    rep_name  = str(row.get("Company Rep",  "") or "").strip()
    rep_pos   = str(row.get("Rep Position", "") or "").strip()
    rep_email = str(row.get("Rep Email",    "") or "").strip()
    rep_phone = str(row.get("Rep Phone",    "") or "").strip()

    # Wikidata enrichment for profile header
    enriched = _fetch_company_data(company)
    if _blank(sector):
        sector = enriched.get("industry") or "—"
    if _blank(country):
        country = enriched.get("hq_country") or "—"
    if _blank(size_global) and enriched.get("employees_global"):
        size_global = enriched["employees_global"] + " employees"

    logo         = _logo_url(company)
    color        = _avatar_color(company)
    initials     = _initials(company)
    sc_color     = _STAGE_COLOR.get(stage, _GREEN)
    est_val      = _fmt_sar(row.get("Est. Investment Value (SAR)"))
    commitment   = _fmt_sar(row.get("Actual Commitment (SAR)"))

    contact_block = ""
    if contact:
        contact_block = (
            f'<div style="margin-top:8px;font-size:12px;color:rgba(255,255,255,0.8);">'
            f'Contact: <strong style="color:#fff;">{contact}</strong>'
            f'{" · " + ctitle if ctitle else ""}</div>'
        )

    size_block = ""
    size_parts = []
    if size_global: size_parts.append(f"🌐 {size_global}")
    if size_ksa:    size_parts.append(f"🇸🇦 {size_ksa} (KSA)")
    if size_parts:
        size_block = (
            f'<div style="margin-top:6px;display:flex;gap:10px;flex-wrap:wrap;">'
            + "".join(
                f'<span style="font-size:11px;color:rgba(255,255,255,0.75);">{p}</span>'
                for p in size_parts
            )
            + '</div>'
        )

    # Rep block for profile header
    _ws_prof = website.lower()
    prof_website_html = ""
    if website and "." in website and _ws_prof not in ("vvvvv", "n/a", "—", "nan"):
        href = website if website.startswith("http") else f"https://{website}"
        disp = website.replace("https://", "").replace("http://", "").rstrip("/")[:40]
        prof_website_html = (
            f'<span style="color:rgba(255,255,255,0.6);font-size:11px;margin-left:8px;">'
            f'<a href="{href}" target="_blank" style="color:{_GOLD};text-decoration:none;">'
            f'&#x1F517; {disp}</a></span>'
        )
    prof_rep_html = ""
    if rep_name and rep_name not in ("—", "nan"):
        _rep_parts = [f'&#x1F91D; {rep_name}']
        if rep_pos:
            _rep_parts.append(rep_pos)
        _rep_line2 = ""
        if rep_email and rep_email not in ("—", "nan"):
            _rep_line2 += (
                f'<a href="mailto:{rep_email}" '
                f'style="color:{_GOLD};text-decoration:none;font-size:10px;">'
                f'{rep_email}</a>'
            )
        if rep_phone and rep_phone not in ("—", "nan"):
            sep = " &nbsp;·&nbsp; " if _rep_line2 else ""
            _rep_line2 += f'{sep}<span style="font-size:10px;">&#x1F4DE; {rep_phone}</span>'
        prof_rep_html = (
            f'<div style="margin-top:6px;font-size:11px;color:rgba(255,255,255,0.75);">'
            f'{" · ".join(_rep_parts)}'
            f'{"<br>" + _rep_line2 if _rep_line2 else ""}'
            f'</div>'
        )

    st.markdown(
        f'<div style="background:linear-gradient(135deg,#0f2d1e,{_GREEN});'
        f'border-radius:12px;padding:20px 24px;margin:12px 0;">'
        f'<div style="display:flex;align-items:center;gap:14px;">'
        f'{_logo_html(company, 56, website)}'
        f'<div style="flex:1;">'
        f'<div style="color:#fff;font-size:22px;font-weight:700;">{company}'
        f'{prof_website_html}</div>'
        f'<div style="color:rgba(255,255,255,0.7);font-size:13px;margin-top:3px;">'
        f'{sector} &nbsp;·&nbsp; {country} &nbsp;·&nbsp; RM: {rm} &nbsp;·&nbsp; AM: {am}'
        f'</div>'
        f'{contact_block}'
        f'{prof_rep_html}'
        f'{size_block}'
        f'</div>'
        f'<div style="text-align:right;">'
        f'<div style="color:{sc_color};font-size:11px;font-weight:600;'
        f'background:rgba(255,255,255,0.15);padding:2px 10px;border-radius:8px;">{stage}</div>'
        f'{"<div style=color:" + _GOLD + ";font-size:13px;font-weight:700;margin-top:6px;>" + est_val + "</div>" if est_val else ""}'
        f'{"<div style=color:rgba(255,255,255,0.6);font-size:11px;>Committed: " + commitment + "</div>" if commitment else ""}'
        f'</div></div></div>',
        unsafe_allow_html=True,
    )

    # Key metrics row
    est   = row.get("Est. Investment Value (SAR)")
    cmmt  = row.get("Actual Commitment (SAR)")
    jobs  = row.get("Est. Jobs Created")
    prio  = row.get("Strategic Priority Score", "—")
    min_a = row.get("Minister Action Required", "None Required")

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Est. Investment", f"SAR {est:,.0f}"  if pd.notna(est)  and est  else "—")
    m2.metric("Commitment",      f"SAR {cmmt:,.0f}" if pd.notna(cmmt) and cmmt else "—")
    m3.metric("Est. Jobs",       f"{int(jobs):,}"   if pd.notna(jobs) and jobs else "—")
    m4.metric("Priority Score",  str(prio) if str(prio) not in ("—", "nan", "") else "—")
    m5.metric("Minister Action", str(min_a) if str(min_a) not in ("None Required", "nan", "") else "None")

    def _link(df, col, val):
        return df[df[col] == val] if not df.empty and col in df.columns else pd.DataFrame()

    linked_meetings = _link(meetings,      "Company Name", company)
    linked_opps     = _link(opportunities, "Company Name", company)
    linked_actions  = _link(actions,       "Company Name", company)
    linked_tasks    = _link(tasks,         "Linked Investor", company)
    linked_deals    = _link(deals,         "Company Name", company)

    if not linked_actions.empty and "Type of Engagement" in linked_actions.columns:
        linked_challenges   = linked_actions[linked_actions["Type of Engagement"] == "Challenge"]
        linked_actions_only = linked_actions[linked_actions["Type of Engagement"] != "Challenge"]
    else:
        linked_challenges   = pd.DataFrame()
        linked_actions_only = linked_actions

    tabs = st.tabs([
        f"Opportunities ({len(linked_opps)})",
        f"Action Items ({len(linked_actions_only)})",
        f"Challenges ({len(linked_challenges)})",
        f"Meetings ({len(linked_meetings)})",
        f"RM Tasks ({len(linked_tasks)})",
        f"Deals ({len(linked_deals)})",
    ])

    with tabs[0]:
        if linked_opps.empty:
            st.info("No opportunities linked yet.")
        else:
            cols = [c for c in ["Opportunity Name", "Opportunity Stage", "Opportunity Status",
                                 "Est. Value (SAR)", "Confidence Level", "Target Closure Date",
                                 "Blockers", "Escalation Required"] if c in linked_opps.columns]
            st.dataframe(linked_opps[cols], use_container_width=True, hide_index=True)

    with tabs[1]:
        if linked_actions_only.empty:
            st.info("No action items yet.")
        else:
            cols = [c for c in ["Action Description", "Type of Engagement", "Status", "Priority",
                                 "Progress", "Due Date", "Assigned To", "Remarks"] if c in linked_actions_only.columns]
            df_s = linked_actions_only[cols].sort_values("Due Date") if "Due Date" in linked_actions_only.columns else linked_actions_only[cols]
            st.dataframe(df_s, use_container_width=True, hide_index=True)

    with tabs[2]:
        if linked_challenges.empty:
            st.info("No challenges logged.")
        else:
            cols = [c for c in ["Action Description", "Status", "Priority", "Due Date",
                                 "Assigned To", "Escalation Flag", "Remarks"] if c in linked_challenges.columns]
            st.dataframe(linked_challenges[cols], use_container_width=True, hide_index=True)

    with tabs[3]:
        if linked_meetings.empty:
            st.info("No meetings logged.")
        else:
            cols = [c for c in ["Meeting Date", "Meeting Type", "Meeting Status", "Meeting Objective",
                                 "Key Discussion Points", "Decisions Made", "Next Steps",
                                 "Follow-Up Owner", "Follow-Up Due Date"] if c in linked_meetings.columns]
            df_s = linked_meetings[cols].sort_values("Meeting Date", ascending=False) if "Meeting Date" in linked_meetings.columns else linked_meetings[cols]
            st.dataframe(df_s, use_container_width=True, hide_index=True)

    with tabs[4]:
        if linked_tasks.empty:
            st.info("No RM tasks linked.")
        else:
            cols = [c for c in ["Task ID", "Task Title", "Priority", "Status", "Due Date", "Notes"] if c in linked_tasks.columns]
            st.dataframe(linked_tasks[cols], use_container_width=True, hide_index=True)

    with tabs[5]:
        if linked_deals.empty:
            st.info("No deals in progress.")
        else:
            cols = [c for c in ["Deal ID", "Deal Name", "Deal Stage", "Deal Status",
                                 "Challenge Severity", "Est. Value (SAR)", "Escalation Required",
                                 "Assigned Owner", "Target Resolution Date"] if c in linked_deals.columns]
            st.dataframe(linked_deals[cols], use_container_width=True, hide_index=True)

    notes = str(row.get("Notes", "") or "")
    if notes and notes not in ("nan", ""):
        st.markdown(f"**Notes:** {notes}")


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
