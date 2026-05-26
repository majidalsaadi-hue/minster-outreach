
# RM Task Management — create, track, and complete RM-level tasks.

import pandas as pd
import streamlit as st
from datetime import date, timedelta

from config.settings import TASK_PRIORITIES, MISA_GREEN, MISA_GOLD
from config.translations import t

TASK_STATUSES = ["Not Started", "In Progress", "Completed"]


def render(dfs: dict, lang: str):
    tasks     = dfs.get("RM Tasks",        pd.DataFrame())
    investors = dfs.get("Investor Master", pd.DataFrame())

    st.markdown(f"### {t('nav_tasks', lang)}")

    # ── Daily summary ─────────────────────────────────────────────────────────
    _render_daily_summary(tasks, lang)

    # ── Add task ──────────────────────────────────────────────────────────────
    with st.expander(f"➕ {t('add_task', lang)}", expanded=False):
        _add_task_form(dfs, investors, lang)

    # ── Pending tasks ─────────────────────────────────────────────────────────
    if tasks.empty:
        st.info("No tasks. Add one above.")
        return

    today = date.today()
    pending   = tasks[tasks.get("Status", pd.Series()) != "Completed"] if "Status" in tasks.columns else tasks
    completed = tasks[tasks.get("Status", pd.Series()) == "Completed"] if "Status" in tasks.columns else pd.DataFrame()

    tab1, tab2 = st.tabs([
        f"🔲 {t('pending_tasks', lang)} ({len(pending)})",
        f"✅ {t('completed_tasks', lang)} ({len(completed)})",
    ])

    with tab1:
        if pending.empty:
            st.success("All tasks complete! 🎉")
        else:
            _render_tasks_table(pending, dfs, lang, today, show_complete_btn=True)

    with tab2:
        if completed.empty:
            st.info("No completed tasks yet.")
        else:
            _render_tasks_table(completed, dfs, lang, today, show_complete_btn=False)


def _render_daily_summary(tasks: pd.DataFrame, lang: str):
    if tasks.empty:
        return
    today = date.today()
    soon  = today + timedelta(days=7)

    pending_count  = 0
    overdue_count  = 0
    due_soon_count = 0

    if "Status" in tasks.columns and "Due Date" in tasks.columns:
        active = tasks[~tasks["Status"].isin(["Completed"])]
        pending_count = len(active)
        due_dates = pd.to_datetime(active["Due Date"], errors="coerce")
        overdue_count  = int((due_dates.dt.date < today).sum())
        due_soon_count = int(((due_dates.dt.date >= today) & (due_dates.dt.date <= soon)).sum())

    c1, c2, c3 = st.columns(3)
    c1.metric(t("pending_tasks", lang), f"{pending_count}")
    c2.metric("Overdue",               f"{overdue_count}",   delta=f"-{overdue_count}" if overdue_count else None, delta_color="inverse")
    c3.metric("Due This Week",         f"{due_soon_count}")


def _render_tasks_table(df: pd.DataFrame, dfs: dict, lang: str, today: date, show_complete_btn: bool):
    display_cols = [
        c for c in [
            "Task ID", "Task Title", "Linked Investor", "Priority",
            "Due Date", "Status", "Notes",
        ] if c in df.columns
    ]

    for _, row in df.iterrows():
        due    = _safe_date(row.get("Due Date"))
        overdue = due and due < today and row.get("Status") != "Completed"
        border  = "2px solid #C0392B" if overdue else "1px solid #E0E0E0"
        color   = "#FFF5F5" if overdue else "#FFFFFF"

        with st.container():
            st.markdown(
                f"<div style='border:{border};background:{color};"
                f"padding:10px;border-radius:8px;margin-bottom:8px;'>",
                unsafe_allow_html=True,
            )
            c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 1, 1])
            c1.markdown(f"**{row.get('Task Title','—')}**")
            c2.markdown(f"*{row.get('Linked Investor','—')}*")

            pri = row.get("Priority", "Medium")
            pri_color = {"High": "#C0392B", "Medium": "#C9974A", "Low": MISA_GREEN}.get(pri, "#9B9B9B")
            c3.markdown(f"<span style='color:{pri_color};font-weight:700'>{pri}</span>",
                        unsafe_allow_html=True)
            c4.markdown(str(due) if due else "—")
            if show_complete_btn and row.get("Status") != "Completed":
                if c5.button("✅", key=f"complete_{row.get('Task ID',id(row))}",
                             help="Mark complete"):
                    _mark_complete(dfs, row.get("Task ID"))
                    st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)


def _add_task_form(dfs: dict, investors: pd.DataFrame, lang: str):
    company_options = [""] + (
        sorted(investors["Company Name"].dropna().unique().tolist())
        if not investors.empty and "Company Name" in investors.columns else []
    )

    with st.form("add_task_form", clear_on_submit=True):
        title     = st.text_input(t("task_title", lang))
        c1, c2 = st.columns(2)
        linked_inv = c1.selectbox(t("linked_investor", lang), company_options)
        priority   = c2.selectbox(t("task_priority", lang), TASK_PRIORITIES)
        c3, c4 = st.columns(2)
        due_date   = c3.date_input(t("task_due_date", lang), value=None)
        status     = c4.selectbox(t("task_status", lang), TASK_STATUSES)
        notes      = st.text_area(t("task_notes", lang))

        if st.form_submit_button(t("add_task", lang)):
            if not title:
                st.warning("Task title is required.")
                return
            tasks  = dfs.get("RM Tasks", pd.DataFrame())
            new_id = _next_id(tasks, "Task ID", "TASK")
            new_row = {
                "Task ID":        new_id,
                "Task Title":     title,
                "Linked Investor":linked_inv or "",
                "Priority":       priority,
                "Due Date":       due_date,
                "Status":         status,
                "Notes":          notes,
                "Created Date":   date.today(),
            }
            dfs["RM Tasks"] = pd.concat(
                [tasks, pd.DataFrame([new_row])], ignore_index=True
            )
            st.success(f"✅ Task '{title}' added.")
            st.rerun()


def _mark_complete(dfs: dict, task_id):
    tasks = dfs.get("RM Tasks", pd.DataFrame())
    if tasks.empty or "Task ID" not in tasks.columns:
        return
    tasks.loc[tasks["Task ID"] == task_id, "Status"] = "Completed"
    dfs["RM Tasks"] = tasks


def _safe_date(val):
    if val is None:
        return None
    if isinstance(val, date):
        return val
    try:
        return pd.to_datetime(val).date()
    except Exception:
        return None


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
