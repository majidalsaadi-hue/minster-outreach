
# Opportunity Pipeline — Excel import with approval flow + minister-grade active board.

import hashlib
import io as _io
import numpy as np
import pandas as pd
import streamlit as st
from collections import defaultdict
from datetime import date, datetime

from config.settings import (
    SECTORS, OPPORTUNITY_TYPES, OPPORTUNITY_SOURCES, OPPORTUNITY_STAGES,
    CONFIDENCE_LEVELS, MISA_GREEN, MISA_GOLD,
    DEAL_STAGES, DEAL_STATUSES, CHALLENGE_CLASSIFICATIONS,
    CHALLENGE_SEVERITIES, ESCALATION_LEVELS,
)
from config.translations import t
from modules.persistence import save_session

_GREEN = "#1B5C3F"
_GOLD  = "#C9974A"
_RED   = "#DC2626"
_AMBER = "#D97706"
_BLUE  = "#1D4ED8"

OPP_STATUSES = ["Active", "Under Review", "Blocked", "Converted to Deal", "Dropped"]

_STAGES = [
    "Exploration",
    "Due Diligence",
    "Active Negotiation",
    "Committed",
    "Post-Investment",
]
_STAGE_COLORS = {
    "Exploration":        "#6B7280",
    "Due Diligence":      "#1D4ED8",
    "Active Negotiation": "#D97706",
    "Committed":          "#059669",
    "Post-Investment":    _GREEN,
}
_CONF_COLORS = {
    "High":        ("#D1FAE5", "#065F46"),
    "Medium":      ("#FEF3C7", "#92400E"),
    "Low":         ("#FEE2E2", "#991B1B"),
    "Suggested":   ("#F3F4F6", "#6B7280"),
    "Speculative": ("#F3F4F6", "#374151"),
}
_CONF_BADGE = {
    "High":      ("#059669", "#fff"),
    "Medium":    ("#D97706", "#fff"),
    "Suggested": ("#6B7280", "#fff"),
}

# Legacy Excel parsing
_LEGACY_PREFIXES = ["Action Items"]
_OPP_SKIP_VALS   = {"opportunity", "opportunities", "type", "n/a", "none", "", "nan",
                     "sector", "engagement type"}
_OPP_KEYWORDS    = [
    "invest", "partnership", "collaborat", "joint venture", "co-invest",
    "fintech", "infrastruc", "technology transfer", "fund", "capital",
    "equity", "roadshow", "pipeline", "strategic partner",
    "ecosystem", "platform", "market entry", "market access",
]


# ── Main render ───────────────────────────────────────────────────────────────

def render(dfs: dict, lang: str):
    opps      = dfs.get("Opportunity Pipeline", pd.DataFrame())
    investors = dfs.get("Investor Master",      pd.DataFrame())
    actions   = dfs.get("Action Items",         pd.DataFrame())

    # Supplement pipeline with action items (Type of Engagement = Opportunity)
    # for any company that has no formal pipeline entry yet.
    if not actions.empty and "Type of Engagement" in actions.columns:
        opp_acts = actions[
            actions["Type of Engagement"].str.strip().str.lower() == "opportunity"
        ].copy()
        if not opp_acts.empty:
            pipeline_cos = (
                set(opps["Company Name"].dropna().unique())
                if not opps.empty and "Company Name" in opps.columns
                else set()
            )
            act_co_col = "Company Name" if "Company Name" in opp_acts.columns else None
            if act_co_col:
                missing = opp_acts[~opp_acts[act_co_col].isin(pipeline_cos)].copy()
            else:
                missing = opp_acts.copy() if opps.empty else pd.DataFrame()
            if not missing.empty:
                missing = missing.rename(columns={"Action Description": "Opportunity Name"})
                missing["Opportunity Status"] = "Active"
                missing["Opportunity Stage"]  = "Exploration"
                missing["Confidence Level"]   = "Suggested"
                if "Opportunity ID" not in missing.columns:
                    missing.insert(0, "Opportunity ID",
                                   [f"ACT-{i+1:03d}" for i in range(len(missing))])
                opps = pd.concat([opps, missing], ignore_index=True) if not opps.empty else missing

    st.markdown(
        f"<h2 style='color:{_GREEN};margin-bottom:2px;'>Opportunity Pipeline</h2>"
        f"<p style='color:#6b7280;font-size:13px;margin-top:0;'>"
        f"Upload your Excel tracker to extract, review and approve investment opportunities</p>",
        unsafe_allow_html=True,
    )

    # ── Import from Excel ──────────────────────────────────────────────────────
    with st.expander("📥 Import Opportunities from Excel Tracker", expanded=opps.empty):
        _render_excel_import(dfs, lang)

    # ── Pending review ─────────────────────────────────────────────────────────
    candidates = st.session_state.get("opp_candidates", [])
    if candidates:
        _render_opp_review(candidates, dfs)
        st.markdown("---")

    # ── Add manually ───────────────────────────────────────────────────────────
    with st.expander("+ Add Opportunity Manually", expanded=False):
        _add_form(dfs, investors, lang)

    # ── Normalise stage column ─────────────────────────────────────────────────
    if not opps.empty and "Opportunity Stage" in opps.columns:
        opps["Opportunity Stage"] = opps["Opportunity Stage"].apply(
            lambda s: s if str(s) in _STAGES else "Exploration"
        )

    if opps.empty:
        _render_empty_state(actions)
    else:
        # ── Minister board — active opportunities ──────────────────────────────
        _render_minister_board(opps)

        st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

        # ── Filters ───────────────────────────────────────────────────────────
        fc1, fc2, fc3, fc4 = st.columns([2.5, 1.5, 1.5, 1.5])
        with fc1:
            q = st.text_input("Search", placeholder="Search company or opportunity…",
                              key="opp_search", label_visibility="collapsed")
        with fc2:
            sel_stage = st.selectbox("Stage", ["All stages"] + _STAGES,
                                     key="opp_stage", label_visibility="collapsed")
        with fc3:
            conf_opts = ["All confidence"] + CONFIDENCE_LEVELS
            sel_conf  = st.selectbox("Confidence", conf_opts,
                                     key="opp_conf", label_visibility="collapsed")
        with fc4:
            sel_status = st.selectbox("Status", ["All statuses"] + OPP_STATUSES,
                                      key="opp_status", label_visibility="collapsed")

        view = opps.copy()
        if q:
            mask = (
                view.get("Company Name",     pd.Series(dtype=str)).fillna("").str.contains(q, case=False, na=False)
                | view.get("Opportunity Name", pd.Series(dtype=str)).fillna("").str.contains(q, case=False, na=False)
            )
            view = view[mask]
        if sel_stage  != "All stages":     view = view[view.get("Opportunity Stage",  pd.Series()) == sel_stage]
        if sel_conf   != "All confidence": view = view[view.get("Confidence Level",   pd.Series()) == sel_conf]
        if sel_status != "All statuses":   view = view[view.get("Opportunity Status", pd.Series()) == sel_status]

        if not view.empty:
            # ── KPI strip ─────────────────────────────────────────────────────
            total_val  = _sum_col(view, "Est. Value (SAR)")
            committed  = int((view.get("Opportunity Stage",  pd.Series()) == "Committed").sum()) if "Opportunity Stage" in view.columns else 0
            high_conf  = int((view.get("Confidence Level",   pd.Series()) == "High").sum())      if "Confidence Level"  in view.columns else 0

            k1, k2, k3, k4 = st.columns(4)
            _kpi(k1, str(len(view)),       "Total Opportunities", _GREEN)
            _kpi(k2, _fmt_sar(total_val),  "Pipeline Value",      _GOLD)
            _kpi(k3, str(committed),       "Committed",           "#059669")
            _kpi(k4, str(high_conf),       "High Confidence",     _BLUE)

            st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

            _render_stage_bar(view)
            st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
            _render_board(view)
        else:
            st.warning("No opportunities match the current filters.")

    # ── Deal Progress (merged section) ─────────────────────────────────────────
    st.markdown("---")
    _render_deal_section(dfs)


# ── Minister board ────────────────────────────────────────────────────────────

def _render_minister_board(opps: pd.DataFrame):
    """Top-of-page minister view: total value + per-company opportunity cards."""
    active_mask = (
        opps.get("Opportunity Status", pd.Series()).str.lower().isin(["active", "under review"])
        if "Opportunity Status" in opps.columns
        else pd.Series([True] * len(opps))
    )
    active = opps[active_mask].copy() if "Opportunity Status" in opps.columns else opps.copy()
    if active.empty:
        active = opps.copy()

    total_val  = _sum_col(active, "Est. Value (SAR)")
    n_active   = len(active)
    n_committed = int((active.get("Opportunity Stage", pd.Series()) == "Committed").sum()) if "Opportunity Stage" in active.columns else 0
    n_companies = active["Company Name"].dropna().nunique() if "Company Name" in active.columns else 0

    # ── Header banner ─────────────────────────────────────────────────────────
    st.markdown(
        f'<div style="background:linear-gradient(135deg,#071a0f 0%,{_GREEN} 100%);'
        f'border-radius:14px;padding:22px 28px;margin-bottom:14px;">'
        f'<div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;">'
        f'<div>'
        f'<div style="color:{_GOLD};font-size:10px;font-weight:700;letter-spacing:1px;'
        f'text-transform:uppercase;margin-bottom:6px;">Active Investment Opportunities</div>'
        f'<div style="color:#fff;font-size:28px;font-weight:700;line-height:1.1;">'
        f'{n_active} Opportunities</div>'
        f'<div style="color:rgba(255,255,255,0.65);font-size:13px;margin-top:5px;">'
        f'{n_companies} companies &nbsp;·&nbsp; {n_committed} committed &nbsp;·&nbsp; {n_active - n_committed} in progress'
        f'</div>'
        f'</div>'
        f'<div style="text-align:right;">'
        f'<div style="color:rgba(255,255,255,0.55);font-size:11px;margin-bottom:4px;">Total Pipeline Value</div>'
        f'<div style="color:{_GOLD};font-size:34px;font-weight:700;line-height:1.1;">'
        f'{_fmt_sar(total_val) if total_val else "—"}</div>'
        f'</div>'
        f'</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ── Per-company cards ─────────────────────────────────────────────────────
    if "Company Name" not in active.columns:
        return

    companies = active["Company Name"].dropna().unique().tolist()
    cols = st.columns(2)
    for i, co in enumerate(companies):
        co_opps = active[active["Company Name"] == co]
        with cols[i % 2]:
            _render_company_opp_card(co, co_opps)


def _render_company_opp_card(company: str, co_opps: pd.DataFrame):
    total_val   = _sum_col(co_opps, "Est. Value (SAR)")
    n           = len(co_opps)
    n_committed = int((co_opps.get("Opportunity Stage", pd.Series()) == "Committed").sum()) if "Opportunity Stage" in co_opps.columns else 0

    opp_rows_html = ""
    for _, row in co_opps.iterrows():
        stage    = str(row.get("Opportunity Stage", "Exploration"))
        name     = str(row.get("Opportunity Name",  "—"))
        val      = row.get("Est. Value (SAR)")
        conf     = str(row.get("Confidence Level",  ""))
        sc       = _STAGE_COLORS.get(stage, "#6B7280")
        val_str  = _fmt_sar(float(val)) if val and str(val) not in ("nan", "None", "") else ""
        conf_bg, conf_fg = _CONF_COLORS.get(conf, ("#F3F4F6", "#6B7280"))

        opp_rows_html += (
            f'<div style="display:flex;align-items:center;gap:8px;'
            f'padding:6px 8px;border-radius:6px;margin-bottom:4px;'
            f'background:#f9fafb;border-left:3px solid {sc};">'
            f'<div style="flex:1;min-width:0;">'
            f'<div style="font-size:11px;font-weight:600;color:#111827;'
            f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">{name}</div>'
            f'<div style="font-size:9px;color:{sc};font-weight:600;">{stage}</div>'
            f'</div>'
            f'<div style="display:flex;gap:4px;align-items:center;flex-shrink:0;">'
            f'{"<span style=font-size:10px;font-weight:700;color:" + _GOLD + ";>" + val_str + "</span>" if val_str else ""}'
            f'{"<span style=background:" + conf_bg + ";color:" + conf_fg + ";padding:1px 5px;border-radius:3px;font-size:8px;font-weight:600;>" + conf + "</span>" if conf and conf not in ("nan","—","") else ""}'
            f'</div>'
            f'</div>'
        )

    st.markdown(
        f'<div style="border:1px solid #E5E7EB;border-radius:10px;padding:14px;'
        f'margin-bottom:10px;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.05);">'
        f'<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px;">'
        f'<div>'
        f'<div style="font-size:14px;font-weight:700;color:{_GREEN};">{company}</div>'
        f'<div style="font-size:10px;color:#9CA3AF;margin-top:2px;">'
        f'{n} opportunity{"s" if n!=1 else ""}'
        f'{" · " + str(n_committed) + " committed" if n_committed else ""}'
        f'</div>'
        f'</div>'
        f'{"<div style=font-size:13px;font-weight:700;color:" + _GOLD + ";>" + _fmt_sar(total_val) + "</div>" if total_val else ""}'
        f'</div>'
        f'{opp_rows_html}'
        f'</div>',
        unsafe_allow_html=True,
    )


# ── Excel import ──────────────────────────────────────────────────────────────

def _render_excel_import(dfs: dict, lang: str):
    st.markdown(
        f'<div style="background:#f0fdf4;border:1px solid #86efac;border-radius:8px;'
        f'padding:12px 16px;margin-bottom:12px;font-size:13px;color:#065F46;">'
        f'Upload your <strong>Excel Action Tracker</strong> (the file with <em>Action Items [Company]</em> sheets). '
        f'The CRM will scan each sheet and extract opportunities from the header block, '
        f'action items tagged as "Opportunity", and action descriptions that suggest investment opportunities. '
        f'You review and approve before anything is saved.'
        f'</div>',
        unsafe_allow_html=True,
    )

    uploaded = st.file_uploader(
        "Upload Action Tracker Excel",
        type=["xlsx"],
        key="opp_import_uploader",
        label_visibility="collapsed",
    )

    if uploaded is not None:
        with st.spinner("Scanning for opportunities…"):
            try:
                file_bytes = uploaded.read()
                candidates = _parse_opps_from_excel(_io.BytesIO(file_bytes), uploaded.name, dfs)
            except Exception as e:
                st.error(f"Failed to parse file: {e}")
                return

        # Also load full Action Items from the tracker so the synthesis
        # (and rest of the CRM) can see them without a separate main-upload.
        try:
            from modules.data_loader import load_excel
            _full = load_excel(_io.BytesIO(file_bytes))
            if _full:
                _new_acts = _full.get("Action Items", pd.DataFrame())
                if not _new_acts.empty:
                    _existing = dfs.get("Action Items", pd.DataFrame())
                    if _existing.empty:
                        dfs["Action Items"] = _new_acts
                    else:
                        dfs["Action Items"] = pd.concat(
                            [_existing, _new_acts], ignore_index=True
                        ).drop_duplicates(
                            subset=["Company Name", "Action Description"],
                            keep="first"
                        ) if "Action Description" in _existing.columns else pd.concat(
                            [_existing, _new_acts], ignore_index=True
                        )
                    from modules.persistence import save_session
                    save_session(dfs)
        except Exception:
            pass

        if not candidates:
            st.warning("No opportunity candidates found in this file. Check that sheets start with 'Action Items'.")
            return

        # Deduplicate against already-saved opportunities
        existing_opps = dfs.get("Opportunity Pipeline", pd.DataFrame())
        if not existing_opps.empty and "Opportunity Name" in existing_opps.columns:
            existing_names = {
                (str(r.get("Company Name","")).lower().strip(),
                 str(r.get("Opportunity Name","")).lower().strip())
                for _, r in existing_opps.iterrows()
            }
            candidates = [
                c for c in candidates
                if (c["company"].lower().strip(), c["opp_name"].lower().strip()) not in existing_names
            ]

        if not candidates:
            st.success("✓ All opportunities from this file are already in the pipeline.")
            return

        n_high  = sum(1 for c in candidates if c["confidence"] == "High")
        n_med   = sum(1 for c in candidates if c["confidence"] == "Medium")
        n_sug   = sum(1 for c in candidates if c["confidence"] == "Suggested")
        st.success(
            f"Found **{len(candidates)}** new candidates: "
            f"{n_high} from header block · {n_med} from action items · {n_sug} suggested"
        )
        st.session_state["opp_candidates"] = candidates
        st.rerun()


def _parse_opps_from_excel(fileobj, filename: str, dfs: dict) -> list[dict]:
    """
    Parse a legacy Excel tracker file and return candidate opportunities.
    Sources (in priority order):
      1. Header block col H rows 13-18  → confidence "High"
      2. Action items Type=Opportunity  → confidence "Medium"
      3. Action items with OPP keywords → confidence "Suggested"
    """
    candidates: list[dict] = []
    seen: set[tuple] = set()

    try:
        raw = pd.ExcelFile(fileobj)
    except Exception:
        return []

    for sheet_name in raw.sheet_names:
        prefix = next((p for p in _LEGACY_PREFIXES if sheet_name.startswith(p)), None)
        if not prefix:
            continue
        company = sheet_name[len(prefix):].strip()
        if not company:
            continue

        try:
            df_raw = raw.parse(sheet_name, header=None)
        except Exception:
            continue

        # ── 1. Header block: rows 13-18, col H (index 7) ─────────────────────
        for row_i in range(13, 19):
            try:
                cell_val = df_raw.iloc[row_i, 7]
            except (IndexError, KeyError):
                continue
            if cell_val is None or (isinstance(cell_val, float) and pd.isna(cell_val)):
                continue
            sval = str(cell_val).strip()
            if not sval or sval.lower() in _OPP_SKIP_VALS:
                continue
            if len(sval) > 1 and sval[0] in "¦¹4" and sval[1] == " ":
                continue
            if sval[0:1] in "¦¹":
                continue
            key = (company.lower(), sval.lower())
            if key not in seen:
                seen.add(key)
                candidates.append({
                    "company":    company,
                    "opp_name":  sval,
                    "sector":    "",
                    "source":    "header_block",
                    "confidence":"High",
                    "context":   "Explicitly listed in tracker header",
                })

        # ── Find action items header row ──────────────────────────────────────
        header_row_idx = None
        for i, hrow in df_raw.iterrows():
            vals = [str(v).strip().lower() for v in hrow if not (isinstance(v, float) and pd.isna(v))]
            if any(m in vals for m in ["action item", "id", "status"]):
                header_row_idx = i
                break
        if header_row_idx is None:
            continue

        try:
            df_data = raw.parse(sheet_name, header=header_row_idx)
        except Exception:
            continue
        df_data.columns = [str(c).strip() for c in df_data.columns]
        df_data = df_data.dropna(how="all")

        action_col = next((c for c in df_data.columns if "action item" in c.lower()), None)
        if not action_col:
            continue
        eng_col  = next((c for c in df_data.columns if "engagement" in c.lower()), None)
        sec_col  = next((c for c in df_data.columns if c.lower() == "sector"), None)
        asgn_col = next((c for c in df_data.columns if "assigned" in c.lower()), None)

        for _, row in df_data.iterrows():
            desc = str(row.get(action_col, "") or "").strip()
            if not desc or desc.lower() in _OPP_SKIP_VALS or len(desc) < 8:
                continue
            eng  = str(row.get(eng_col,  "") or "").strip().lower() if eng_col  else ""
            sec  = str(row.get(sec_col,  "") or "").strip()          if sec_col  else ""
            asgn = str(row.get(asgn_col, "") or "").strip()          if asgn_col else ""
            for bad in ("nan", "None"):
                sec  = sec.replace(bad, "").strip()
                asgn = asgn.replace(bad, "").strip()

            key = (company.lower(), desc.lower())

            # ── 2. Explicit "Opportunity" engagement type ─────────────────────
            if "opportunity" in eng:
                if key not in seen:
                    seen.add(key)
                    candidates.append({
                        "company":    company,
                        "opp_name":  desc[:160],
                        "sector":    sec,
                        "source":    "action_item",
                        "confidence":"Medium",
                        "context":   f"Action item · Type=Opportunity{' · ' + asgn if asgn else ''}",
                    })

            # ── 3. Keyword match — investment-related wording ─────────────────
            elif any(kw in desc.lower() for kw in _OPP_KEYWORDS):
                if key not in seen:
                    seen.add(key)
                    candidates.append({
                        "company":    company,
                        "opp_name":  desc[:160],
                        "sector":    sec,
                        "source":    "suggested",
                        "confidence":"Suggested",
                        "context":   f"Inferred from action description{' · ' + asgn if asgn else ''}",
                    })

    return candidates


# ── Opportunity review / approval UI ─────────────────────────────────────────

def _render_opp_review(candidates: list[dict], dfs: dict):
    n_high = sum(1 for c in candidates if c["confidence"] == "High")
    n_med  = sum(1 for c in candidates if c["confidence"] == "Medium")
    n_sug  = sum(1 for c in candidates if c["confidence"] == "Suggested")

    st.markdown(
        f'<div style="background:linear-gradient(135deg,#071a0f,{_GREEN});'
        f'border-radius:12px 12px 0 0;padding:16px 22px;">'
        f'<div style="color:{_GOLD};font-size:10px;font-weight:700;letter-spacing:.8px;'
        f'text-transform:uppercase;margin-bottom:4px;">Review Extracted Opportunities</div>'
        f'<div style="color:#fff;font-size:17px;font-weight:700;">'
        f'{len(candidates)} candidates found across {len({c["company"] for c in candidates})} companies</div>'
        f'<div style="display:flex;gap:12px;margin-top:8px;">'
        f'<span style="background:rgba(5,150,105,0.3);color:#6EE7B7;padding:2px 10px;'
        f'border-radius:8px;font-size:11px;font-weight:600;">✓ {n_high} from header</span>'
        f'<span style="background:rgba(217,119,6,0.3);color:#FCD34D;padding:2px 10px;'
        f'border-radius:8px;font-size:11px;font-weight:600;">⚡ {n_med} from action items</span>'
        f'<span style="background:rgba(107,114,128,0.3);color:#D1D5DB;padding:2px 10px;'
        f'border-radius:8px;font-size:11px;font-weight:600;">? {n_sug} suggested</span>'
        f'</div></div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div style="background:#0a1f14;border-radius:0 0 12px 12px;'
        'padding:14px 16px;margin-bottom:16px;">',
        unsafe_allow_html=True,
    )

    # ── Column headers ────────────────────────────────────────────────────────
    hc0, hc1, hc2, hc3, hc4, hc5 = st.columns([0.5, 1.2, 3.5, 2, 2, 1.8])
    for col, label in zip(
        [hc0, hc1, hc2, hc3, hc4, hc5],
        ["✓", "Confidence", "Opportunity Name", "Sector", "Est. Value (SAR)", "Stage"],
    ):
        col.markdown(
            f'<div style="font-size:9px;font-weight:700;color:rgba(255,255,255,0.4);'
            f'letter-spacing:.4px;text-transform:uppercase;padding-bottom:4px;">{label}</div>',
            unsafe_allow_html=True,
        )

    # ── Group by company ──────────────────────────────────────────────────────
    by_company: dict[str, list] = defaultdict(list)
    for i, c in enumerate(candidates):
        by_company[c["company"]].append((i, c))

    for company, items in by_company.items():
        st.markdown(
            f'<div style="color:{_GOLD};font-size:12px;font-weight:700;'
            f'margin:10px 0 6px 0;padding-left:4px;border-left:3px solid {_GOLD};">'
            f'{company}</div>',
            unsafe_allow_html=True,
        )

        for global_idx, cand in items:
            conf       = cand["confidence"]
            badge_bg, badge_fg = _CONF_BADGE.get(conf, ("#6B7280", "#fff"))
            default_checked = conf != "Suggested"

            c0, c1, c2, c3, c4, c5 = st.columns([0.5, 1.2, 3.5, 2, 2, 1.8])
            c0.checkbox(
                "",
                value=default_checked,
                key=f"opp_ck_{global_idx}",
                label_visibility="collapsed",
            )
            c1.markdown(
                f'<span style="background:{badge_bg};color:{badge_fg};padding:2px 8px;'
                f'border-radius:6px;font-size:10px;font-weight:600;">{conf}</span>'
                f'<div style="font-size:9px;color:rgba(255,255,255,0.35);margin-top:2px;'
                f'line-height:1.3;">{cand["context"][:45]}</div>',
                unsafe_allow_html=True,
            )
            c2.text_input(
                "Name",
                value=cand["opp_name"],
                key=f"opp_nm_{global_idx}",
                label_visibility="collapsed",
            )
            c3.text_input(
                "Sector",
                value=cand.get("sector", ""),
                placeholder="e.g. Fintech",
                key=f"opp_sc_{global_idx}",
                label_visibility="collapsed",
            )
            c4.number_input(
                "Value",
                min_value=0.0,
                step=1_000_000.0,
                value=0.0,
                format="%g",
                key=f"opp_val_{global_idx}",
                label_visibility="collapsed",
            )
            c5.selectbox(
                "Stage",
                _STAGES,
                index=0,
                key=f"opp_stg_{global_idx}",
                label_visibility="collapsed",
            )

    st.markdown("</div>", unsafe_allow_html=True)

    # ── Action buttons ────────────────────────────────────────────────────────
    n_checked = sum(
        1 for i in range(len(candidates))
        if st.session_state.get(f"opp_ck_{i}", False)
    )

    btn1, btn2 = st.columns([2, 1])
    if btn1.button(
        f"✓ Approve {n_checked} Selected Opportunit{'y' if n_checked==1 else 'ies'}",
        type="primary",
        use_container_width=True,
        disabled=(n_checked == 0),
    ):
        approved = []
        for i, cand in enumerate(candidates):
            if st.session_state.get(f"opp_ck_{i}", False):
                approved.append({
                    **cand,
                    "opp_name": st.session_state.get(f"opp_nm_{i}",  cand["opp_name"]),
                    "sector":   st.session_state.get(f"opp_sc_{i}",  cand.get("sector", "")),
                    "est_val":  st.session_state.get(f"opp_val_{i}", 0.0),
                    "stage":    st.session_state.get(f"opp_stg_{i}", "Exploration"),
                })
        _bulk_add_to_pipeline(approved, dfs)
        del st.session_state["opp_candidates"]
        # Clean up widget state
        for i in range(len(candidates)):
            for suffix in ("ck", "nm", "sc", "val", "stg"):
                st.session_state.pop(f"opp_{suffix}_{i}", None)
        st.success(f"✓ Added {len(approved)} opportunities to the pipeline!")
        st.rerun()

    if btn2.button("✗ Cancel / Clear", use_container_width=True):
        del st.session_state["opp_candidates"]
        st.rerun()


def _bulk_add_to_pipeline(approved: list[dict], dfs: dict):
    opps      = dfs.get("Opportunity Pipeline", pd.DataFrame())
    investors = dfs.get("Investor Master",      pd.DataFrame())

    ids = _gen_opp_ids(opps, len(approved))
    new_rows = []
    for opp_id, cand in zip(ids, approved):
        company  = cand["company"]
        inv_id   = _get_investor_id(investors, company)
        conf_map = {"High": "High", "Medium": "Medium", "Suggested": "Low"}
        new_rows.append({
            "Opportunity ID":     opp_id,
            "Investor ID":        inv_id,
            "Company Name":       company,
            "Opportunity Name":   cand["opp_name"],
            "Sector":             cand.get("sector", ""),
            "Opportunity Stage":  cand.get("stage", "Exploration"),
            "Opportunity Status": "Active",
            "Opportunity Type":   "Opportunity",
            "Opportunity Source": f"Excel Import — {cand['source']}",
            "Confidence Level":   conf_map.get(cand["confidence"], "Medium"),
            "Est. Value (SAR)":   cand.get("est_val") or None,
            "Last Updated":       np.datetime64(datetime.now()),
        })

    if new_rows:
        new_df = pd.DataFrame(new_rows)
        dfs["Opportunity Pipeline"] = pd.concat([opps, new_df], ignore_index=True)
        save_session(dfs)


def _gen_opp_ids(existing: pd.DataFrame, n: int) -> list[str]:
    nums = []
    if not existing.empty and "Opportunity ID" in existing.columns:
        for v in existing["Opportunity ID"].dropna():
            try:
                nums.append(int(str(v).split("-")[-1]))
            except ValueError:
                pass
    start = (max(nums) + 1) if nums else 1
    return [f"OPP-{start+i:03d}" for i in range(n)]


# ── Empty state ───────────────────────────────────────────────────────────────

def _render_empty_state(actions: pd.DataFrame):
    n_act_opps = 0
    if not actions.empty and "Type of Engagement" in actions.columns:
        n_act_opps = int(
            (actions["Type of Engagement"].str.strip().str.lower() == "opportunity").sum()
        )

    suggestion_text = ""
    if n_act_opps:
        suggestion_text = (
            f"<br>Your Action Items already contain <strong>{n_act_opps} Opportunity-type actions</strong> "
            f"across your investor portfolio. Upload the Excel Tracker above to extract them."
        )

    st.markdown(
        f'<div style="background:#f0fdf4;border:2px dashed #86efac;border-radius:12px;'
        f'padding:40px;text-align:center;margin-top:16px;">'
        f'<div style="font-size:40px;margin-bottom:12px;">🎯</div>'
        f'<h3 style="color:{_GREEN};margin-bottom:8px;">No opportunities yet</h3>'
        f'<p style="color:#6b7280;font-size:13px;max-width:480px;margin:0 auto;">'
        f'Upload your Excel Action Tracker above and the CRM will automatically scan '
        f'each sheet for opportunity names and investment-related actions.'
        f'{suggestion_text}</p>'
        f'</div>',
        unsafe_allow_html=True,
    )


# ── Stage progress bar ────────────────────────────────────────────────────────

def _render_stage_bar(df: pd.DataFrame):
    if "Opportunity Stage" not in df.columns:
        return
    counts = df["Opportunity Stage"].value_counts().to_dict()
    total  = sum(counts.values()) or 1

    segs_html = ""
    for stage in _STAGES:
        cnt = counts.get(stage, 0)
        pct = cnt / total * 100
        if pct < 1:
            continue
        color = _STAGE_COLORS.get(stage, "#6B7280")
        segs_html += (
            f'<div style="flex:{pct};background:{color};height:100%;'
            f'display:flex;align-items:center;justify-content:center;'
            f'font-size:10px;color:#fff;font-weight:600;white-space:nowrap;'
            f'overflow:hidden;padding:0 4px;min-width:0;" title="{stage}: {cnt}">'
            f'{cnt}</div>'
        )

    legend_html = ""
    for stage in _STAGES:
        cnt   = counts.get(stage, 0)
        color = _STAGE_COLORS.get(stage, "#6B7280")
        legend_html += (
            f'<span style="display:inline-flex;align-items:center;gap:4px;'
            f'margin-right:14px;font-size:11px;color:#374151;">'
            f'<span style="width:10px;height:10px;border-radius:2px;'
            f'background:{color};display:inline-block;flex-shrink:0;"></span>'
            f'{stage} ({cnt})</span>'
        )

    st.markdown(
        f'<div style="border-radius:6px;overflow:hidden;height:28px;display:flex;margin-bottom:8px;">'
        f'{segs_html}</div>'
        f'<div style="margin-bottom:4px;">{legend_html}</div>',
        unsafe_allow_html=True,
    )


# ── Kanban board ──────────────────────────────────────────────────────────────

def _render_board(df: pd.DataFrame):
    stage_col = "Opportunity Stage" if "Opportunity Stage" in df.columns else None
    if stage_col is None:
        _render_list(df)
        return

    active_stages = [s for s in _STAGES if s in df[stage_col].values]
    if not active_stages:
        _render_list(df)
        return

    cols = st.columns(len(active_stages))
    for col_idx, stage in enumerate(active_stages):
        stage_df = df[df[stage_col] == stage]
        color    = _STAGE_COLORS.get(stage, "#6B7280")
        with cols[col_idx]:
            val = _sum_col(stage_df, "Est. Value (SAR)")
            st.markdown(
                f'<div style="background:{color}18;border:1px solid {color}40;'
                f'border-radius:8px;padding:8px 10px;margin-bottom:8px;">'
                f'<div style="color:{color};font-size:11px;font-weight:700;'
                f'letter-spacing:.4px;text-transform:uppercase;">{stage}</div>'
                f'<div style="display:flex;justify-content:space-between;align-items:center;margin-top:3px;">'
                f'<span style="color:#374151;font-size:12px;font-weight:600;">'
                f'{len(stage_df)} opp{"s" if len(stage_df)!=1 else ""}</span>'
                f'<span style="color:{color};font-size:11px;font-weight:600;">'
                f'{_fmt_sar(val)}</span>'
                f'</div></div>',
                unsafe_allow_html=True,
            )
            for _, row in stage_df.iterrows():
                _opp_card(row)


def _render_list(df: pd.DataFrame):
    for _, row in df.iterrows():
        _opp_card(row)


def _opp_card(row):
    company  = str(row.get("Company Name",       "?"))
    opp_name = str(row.get("Opportunity Name",   "—"))
    stage    = str(row.get("Opportunity Stage",  "—"))
    conf     = str(row.get("Confidence Level",   "—"))
    status   = str(row.get("Opportunity Status", "Active"))
    val      = row.get("Est. Value (SAR)", None)
    am       = str(row.get("Assigned AM", "") or row.get("Account Manager", "") or "—")
    opp_type = str(row.get("Opportunity Type", ""))
    sector   = str(row.get("Sector", ""))

    color         = _STAGE_COLORS.get(stage, "#6B7280")
    conf_bg, conf_fg = _CONF_COLORS.get(conf, ("#F3F4F6", "#374151"))
    val_str       = _fmt_sar(float(val)) if val and str(val) not in ("nan", "None", "") else ""
    avatar        = _initials(company)
    avatar_color  = _company_color(company)

    status_badge = ""
    if status.lower() in ("blocked", "dropped"):
        status_badge = f'<span style="background:#FEE2E2;color:#991B1B;padding:1px 6px;border-radius:4px;font-size:9px;font-weight:600;margin-left:4px;">{status.upper()}</span>'
    elif status.lower() == "converted to deal":
        status_badge = f'<span style="background:#D1FAE5;color:#065F46;padding:1px 6px;border-radius:4px;font-size:9px;font-weight:600;margin-left:4px;">DEAL</span>'

    sector_tag = f'<span style="background:#f8fafc;color:#6b7280;padding:1px 5px;border-radius:4px;font-size:9px;">{sector}</span>' if sector and sector != "nan" else ""
    type_tag   = f'<span style="background:#ede9fe;color:#7c3aed;padding:1px 5px;border-radius:4px;font-size:9px;">{opp_type}</span>' if opp_type and opp_type not in ("nan", "Opportunity") else ""

    st.markdown(
        f'<div style="border:1px solid #e5e7eb;border-left:3px solid {color};'
        f'border-radius:8px;padding:10px;margin-bottom:6px;background:#fff;">'
        f'<div style="display:flex;align-items:center;gap:7px;margin-bottom:6px;">'
        f'<div style="width:28px;height:28px;border-radius:6px;background:{avatar_color};'
        f'display:flex;align-items:center;justify-content:center;'
        f'font-size:10px;font-weight:700;color:#fff;flex-shrink:0;">{avatar}</div>'
        f'<div style="flex:1;min-width:0;">'
        f'<div style="font-size:11px;font-weight:700;color:{_GREEN};'
        f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">{company}</div>'
        f'</div>'
        f'{"<span style=font-size:12px;font-weight:700;color:" + _GOLD + ";>" + val_str + "</span>" if val_str else ""}'
        f'</div>'
        f'<div style="font-size:12px;font-weight:600;color:#111827;margin-bottom:5px;'
        f'line-height:1.3;">{opp_name}{status_badge}</div>'
        f'<div style="display:flex;flex-wrap:wrap;gap:4px;margin-bottom:5px;">'
        f'<span style="background:{conf_bg};color:{conf_fg};padding:1px 6px;'
        f'border-radius:4px;font-size:9px;font-weight:600;">{conf}</span>'
        f'{sector_tag}{type_tag}'
        f'</div>'
        f'{"<div style=font-size:10px;color:#9ca3af;>AM: " + am + "</div>" if am != "—" else ""}'
        f'</div>',
        unsafe_allow_html=True,
    )


# ── Add opportunity form ──────────────────────────────────────────────────────

def _add_form(dfs: dict, investors: pd.DataFrame, lang: str):
    company_options = [""] + (
        sorted(investors["Company Name"].dropna().unique().tolist())
        if not investors.empty and "Company Name" in investors.columns else []
    )
    with st.form("add_opp_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        company  = c1.selectbox("Company", company_options)
        opp_name = c2.text_input("Opportunity Name")
        c3, c4 = st.columns(2)
        opp_type = c3.selectbox("Type", OPPORTUNITY_TYPES)
        source   = c4.selectbox("Source", OPPORTUNITY_SOURCES)
        c5, c6 = st.columns(2)
        sector   = c5.selectbox("Sector", [""] + SECTORS)
        stage    = c6.selectbox("Stage", _STAGES)
        c7, c8 = st.columns(2)
        conf     = c7.selectbox("Confidence Level", CONFIDENCE_LEVELS)
        status   = c8.selectbox("Status", OPP_STATUSES)
        c9, c10 = st.columns(2)
        est_val  = c9.number_input("Est. Value (SAR)", min_value=0.0, step=1_000_000.0)
        target_d = c10.date_input("Target Closure", value=None)
        c11, c12 = st.columns(2)
        assigned = c11.text_input("Assigned AM")
        escalate = c12.selectbox("Escalation Required", ["No", "Yes"])
        blockers = st.text_area("Blockers / Notes")

        if st.form_submit_button("Add Opportunity", use_container_width=True):
            if not company or not opp_name:
                st.warning("Company and opportunity name are required.")
                return
            opps   = dfs.get("Opportunity Pipeline", pd.DataFrame())
            inv_id = _get_investor_id(investors, company)
            new_id = _next_id(opps, "Opportunity ID", "OPP")
            new_row = {
                "Opportunity ID":     new_id,
                "Investor ID":        inv_id,
                "Company Name":       company,
                "Opportunity Name":   opp_name,
                "Opportunity Type":   opp_type,
                "Opportunity Source": source,
                "Sector":             sector,
                "Opportunity Stage":  stage,
                "Confidence Level":   conf,
                "Opportunity Status": status,
                "Est. Value (SAR)":   est_val if est_val > 0 else None,
                "Assigned AM":        assigned,
                "Start Date":         date.today(),
                "Target Closure Date":target_d,
                "Blockers":           blockers,
                "Escalation Required":escalate,
                "Last Updated":       np.datetime64(datetime.now()),
            }
            dfs["Opportunity Pipeline"] = pd.concat(
                [opps, pd.DataFrame([new_row])], ignore_index=True
            )
            save_session(dfs)
            st.success(f"Opportunity {new_id} added for {company}")
            st.rerun()


# ── Deal Progress section ─────────────────────────────────────────────────────

_SEV_COLORS = {
    "Critical": (_RED,      "#fff"),
    "High":     (_AMBER,    "#fff"),
    "Medium":   ("#6B7280", "#fff"),
    "Low":      (_GREEN,    "#fff"),
}
_DEAL_STAT_COLOR = {
    "Active":        _GREEN,
    "On Hold":       _AMBER,
    "Blocked":       _RED,
    "Closed — Won":  "#059669",
    "Closed — Lost": "#6B7280",
}


def _render_deal_section(dfs: dict):
    deals     = dfs.get("Deal Progress",   pd.DataFrame())
    investors = dfs.get("Investor Master", pd.DataFrame())

    n_deals    = len(deals)
    n_blocked  = int((deals["Deal Status"]        == "Blocked").sum())  if not deals.empty and "Deal Status"        in deals.columns else 0
    n_critical = int((deals["Challenge Severity"] == "Critical").sum()) if not deals.empty and "Challenge Severity" in deals.columns else 0
    n_active   = int(deals["Deal Status"].isin(["Active", "On Hold"]).sum()) if not deals.empty and "Deal Status" in deals.columns else 0
    total_val  = _sum_col(deals, "Est. Value (SAR)")

    header_bg = "linear-gradient(135deg,#1a0505 0%,#7f1d1d 100%)" if n_blocked else f"linear-gradient(135deg,#071a0f 0%,{_GREEN} 80%)"
    st.markdown(
        f'<div style="background:{header_bg};border-radius:14px;padding:20px 28px;margin-bottom:14px;">'
        f'<div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;">'
        f'<div>'
        f'<div style="color:{_GOLD};font-size:10px;font-weight:700;letter-spacing:1px;'
        f'text-transform:uppercase;margin-bottom:6px;">Deal Progress Tracker</div>'
        f'<div style="color:#fff;font-size:26px;font-weight:700;line-height:1.1;">'
        f'{n_deals} Deal{"s" if n_deals != 1 else ""}</div>'
        f'<div style="color:rgba(255,255,255,0.65);font-size:13px;margin-top:5px;">'
        f'{n_blocked} blocked &nbsp;·&nbsp; {n_critical} critical &nbsp;·&nbsp; {n_active} active'
        f'</div>'
        f'</div>'
        f'<div style="text-align:right;">'
        f'<div style="color:rgba(255,255,255,0.55);font-size:11px;margin-bottom:4px;">Total Deal Value</div>'
        f'<div style="color:{_GOLD};font-size:30px;font-weight:700;line-height:1.1;">'
        f'{_fmt_sar(total_val) if total_val else "—"}'
        f'</div></div></div></div>',
        unsafe_allow_html=True,
    )

    k1, k2, k3, k4 = st.columns(4)
    _kpi(k1, str(n_deals),    "Total Deals",      "#1F2937")
    _kpi(k2, str(n_active),   "Active",            _GREEN)
    _kpi(k3, str(n_blocked),  "Blocked",           _RED    if n_blocked  else "#6B7280")
    _kpi(k4, str(n_critical), "Critical Severity", _RED    if n_critical else "#6B7280")

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    if not deals.empty:
        companies = deals["Company Name"].dropna().unique().tolist() if "Company Name" in deals.columns else []
        if companies:
            cols = st.columns(2)
            for i, co in enumerate(companies):
                co_deals = deals[deals["Company Name"] == co]
                with cols[i % 2]:
                    _render_deal_company_card(co, co_deals)
        else:
            _render_deal_list(deals)
    else:
        st.markdown(
            f'<div style="background:#fafafa;border:2px dashed #e5e7eb;border-radius:10px;'
            f'padding:30px;text-align:center;margin-bottom:12px;">'
            f'<div style="font-size:32px;margin-bottom:8px;">📋</div>'
            f'<div style="color:#6B7280;font-size:13px;">'
            f'No deals tracked yet. Use the form below to add one.</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    with st.expander("+ Add Deal", expanded=False):
        _add_deal_form(dfs, investors)


def _render_deal_company_card(company: str, co_deals: pd.DataFrame):
    n          = len(co_deals)
    n_blocked  = int((co_deals.get("Deal Status", pd.Series()) == "Blocked").sum()) if "Deal Status" in co_deals.columns else 0
    total_val  = _sum_col(co_deals, "Est. Value (SAR)")

    rows_html = ""
    for _, row in co_deals.iterrows():
        dname  = str(row.get("Deal Name",          "—"))
        dstg   = str(row.get("Deal Stage",         "—"))
        dstat  = str(row.get("Deal Status",        "—"))
        dsev   = str(row.get("Challenge Severity", ""))
        sev_bg, sev_fg = _SEV_COLORS.get(dsev, ("#F3F4F6", "#6B7280"))
        stat_c = _DEAL_STAT_COLOR.get(dstat, "#6B7280")
        val    = row.get("Est. Value (SAR)")
        val_str = _fmt_sar(float(val)) if val and str(val) not in ("nan", "None", "") else ""

        rows_html += (
            f'<div style="display:flex;align-items:center;gap:8px;'
            f'padding:6px 8px;border-radius:6px;margin-bottom:4px;'
            f'background:#f9fafb;border-left:3px solid {stat_c};">'
            f'<div style="flex:1;min-width:0;">'
            f'<div style="font-size:11px;font-weight:600;color:#111827;'
            f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">{dname}</div>'
            f'<div style="font-size:9px;color:{stat_c};font-weight:600;">{dstg} · {dstat}</div>'
            f'</div>'
            f'<div style="display:flex;gap:4px;align-items:center;flex-shrink:0;">'
            f'{"<span style=font-size:10px;font-weight:700;color:" + _GOLD + ";>" + val_str + "</span>" if val_str else ""}'
            f'{"<span style=background:" + sev_bg + ";color:" + sev_fg + ";padding:1px 5px;border-radius:3px;font-size:8px;font-weight:600;>" + dsev + "</span>" if dsev and dsev not in ("nan","—","") else ""}'
            f'</div></div>'
        )

    border_col = _RED if n_blocked else "#E5E7EB"
    st.markdown(
        f'<div style="border:1px solid {border_col};border-radius:10px;padding:14px;'
        f'margin-bottom:10px;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.05);">'
        f'<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px;">'
        f'<div>'
        f'<div style="font-size:14px;font-weight:700;color:{_GREEN};">{company}</div>'
        f'<div style="font-size:10px;color:#9CA3AF;margin-top:2px;">'
        f'{n} deal{"s" if n!=1 else ""}'
        f'{" · " + str(n_blocked) + " blocked" if n_blocked else ""}'
        f'</div></div>'
        f'{"<div style=font-size:13px;font-weight:700;color:" + _GOLD + ";>" + _fmt_sar(total_val) + "</div>" if total_val else ""}'
        f'</div>'
        f'{rows_html}'
        f'</div>',
        unsafe_allow_html=True,
    )


def _render_deal_list(deals: pd.DataFrame):
    for _, row in deals.iterrows():
        dname  = str(row.get("Deal Name",          "—"))
        dstg   = str(row.get("Deal Stage",         "—"))
        dstat  = str(row.get("Deal Status",        "Active"))
        dsev   = str(row.get("Challenge Severity", ""))
        co     = str(row.get("Company Name",       ""))
        val    = row.get("Est. Value (SAR)")
        sev_bg, sev_fg = _SEV_COLORS.get(dsev, ("#F3F4F6", "#6B7280"))
        stat_c = _DEAL_STAT_COLOR.get(dstat, "#6B7280")
        val_str = _fmt_sar(float(val)) if val and str(val) not in ("nan", "None", "") else ""

        st.markdown(
            f'<div style="border:1px solid #e5e7eb;border-left:3px solid {stat_c};'
            f'border-radius:8px;padding:10px;margin-bottom:6px;background:#fff;">'
            f'<div style="display:flex;justify-content:space-between;align-items:center;">'
            f'<div>'
            f'<div style="font-size:12px;font-weight:700;color:#111827;">{dname}</div>'
            f'<div style="font-size:10px;color:{stat_c};margin-top:2px;">'
            f'{co}  ·  {dstg}  ·  {dstat}</div>'
            f'</div>'
            f'<div style="display:flex;gap:6px;align-items:center;">'
            f'{"<span style=font-size:11px;font-weight:700;color:" + _GOLD + ";>" + val_str + "</span>" if val_str else ""}'
            f'{"<span style=background:" + sev_bg + ";color:" + sev_fg + ";padding:2px 7px;border-radius:4px;font-size:9px;font-weight:600;>" + dsev + "</span>" if dsev and dsev not in ("nan","—","") else ""}'
            f'</div></div></div>',
            unsafe_allow_html=True,
        )


def _add_deal_form(dfs: dict, investors: pd.DataFrame):
    company_options = [""] + (
        sorted(investors["Company Name"].dropna().unique().tolist())
        if not investors.empty and "Company Name" in investors.columns else []
    )
    with st.form("add_deal_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        company   = c1.selectbox("Company",    company_options, key="deal_co")
        deal_name = c2.text_input("Deal Name")
        c3, c4 = st.columns(2)
        stage  = c3.selectbox("Deal Stage",          DEAL_STAGES)
        status = c4.selectbox("Deal Status",         DEAL_STATUSES)
        c5, c6 = st.columns(2)
        severity  = c5.selectbox("Challenge Severity", CHALLENGE_SEVERITIES)
        esc_req   = c6.selectbox("Escalation Required", ["No", "Yes"])
        c7, c8 = st.columns(2)
        est_val   = c7.number_input("Est. Value (SAR)", min_value=0.0, step=1_000_000.0)
        target_d  = c8.date_input("Target Resolution Date", value=None)
        c9, c10 = st.columns(2)
        challenge_class = c9.selectbox("Challenge Classification", [""] + CHALLENGE_CLASSIFICATIONS)
        esc_level       = c10.selectbox("Escalation Level",         ESCALATION_LEVELS)
        challenge_desc = st.text_area("Challenge Description")
        solution       = st.text_area("Proposed Solution")

        if st.form_submit_button("Add Deal", use_container_width=True):
            if not company or not deal_name:
                st.warning("Company and deal name are required.")
                return
            deals  = dfs.get("Deal Progress", pd.DataFrame())
            inv_id = _get_investor_id(investors, company)
            new_id = _next_id(deals, "Deal ID", "DL")
            new_row = {
                "Deal ID":                   new_id,
                "Investor ID":               inv_id,
                "Company Name":              company,
                "Deal Name":                 deal_name,
                "Deal Stage":                stage,
                "Deal Status":               status,
                "Challenge Severity":        severity,
                "Challenge Classification":  challenge_class,
                "Challenge Description":     challenge_desc,
                "Proposed Solution":         solution,
                "Escalation Required":       esc_req,
                "Escalation Level":          esc_level,
                "Est. Value (SAR)":          est_val if est_val > 0 else None,
                "Target Resolution Date":    target_d,
                "Last Updated":              date.today(),
            }
            dfs["Deal Progress"] = pd.concat(
                [deals, pd.DataFrame([new_row])], ignore_index=True
            )
            save_session(dfs)
            st.success(f"Deal {new_id} added for {company}")
            st.rerun()


# ── Helpers ───────────────────────────────────────────────────────────────────

def _kpi(col, value: str, label: str, color: str):
    col.markdown(
        f'<div style="background:{color};padding:14px 10px;border-radius:8px;text-align:center;">'
        f'<div style="color:rgba(255,255,255,0.75);font-size:9px;font-weight:600;'
        f'letter-spacing:.5px;text-transform:uppercase;margin-bottom:4px;">{label}</div>'
        f'<div style="color:#fff;font-size:22px;font-weight:700;line-height:1.1;">{value}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


def _initials(name: str) -> str:
    parts = name.split()
    if len(parts) >= 2:
        return (parts[0][0] + parts[-1][0]).upper()
    return name[:2].upper() if name else "?"


def _company_color(name: str) -> str:
    h = int(hashlib.md5(name.encode()).hexdigest()[:6], 16)
    r = max(40, min(int((h >> 16) & 0xFF), 160))
    g = max(40, min(int((h >> 8)  & 0xFF), 160))
    b = max(40, min(int(h & 0xFF),          160))
    return f"#{r:02X}{g:02X}{b:02X}"


def _sum_col(df: pd.DataFrame, col: str) -> float:
    if df.empty or col not in df.columns:
        return 0.0
    return pd.to_numeric(df[col], errors="coerce").fillna(0).sum()


def _fmt_sar(val: float) -> str:
    if not val or (isinstance(val, float) and pd.isna(val)):
        return "—"
    if val >= 1e9:
        return f"SAR {val/1e9:.1f}B"
    if val >= 1e6:
        return f"SAR {val/1e6:.0f}M"
    if val > 0:
        return f"SAR {val:,.0f}"
    return "—"


def _get_investor_id(investors: pd.DataFrame, company: str) -> str:
    if investors.empty or "Company Name" not in investors.columns:
        return ""
    m = investors[investors["Company Name"] == company]
    return str(m.iloc[0].get("Investor ID", "")) if not m.empty else ""


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
