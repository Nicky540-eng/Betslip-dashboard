"""Cashier Slip Dashboard — branch → cashier → bet slips per game.

Run locally:   streamlit run app.py
Neon:          put DATABASE_URL in .streamlit/secrets.toml (or Streamlit Cloud secrets)
"""
from __future__ import annotations

import os

import pandas as pd
import streamlit as st

import slipdb as db

st.set_page_config(page_title="Cashier Bet Slips", layout="wide")

# ---------------------------------------------------------------------------
# Connection
# ---------------------------------------------------------------------------
def _database_url() -> str | None:
    try:
        url = st.secrets.get("DATABASE_URL")
        if url:
            return url
    except Exception:
        pass
    return os.environ.get("DATABASE_URL")


@st.cache_resource
def get_engine():
    engine = db.make_engine(_database_url())
    db.init_db(engine)
    # First start: load the bundled Jan–Jul 2026 files into Neon (skipped once they're there).
    db.seed_from_folder(engine, os.path.join(os.path.dirname(os.path.abspath(__file__)), "seed_data"))
    return engine


@st.cache_data(ttl=900, show_spinner="Loading data…")
def cached_slips(version: str) -> pd.DataFrame:
    return db.load_slips(get_engine())


engine = get_engine()
page = st.sidebar.radio("Page", ["Dashboard", "Upload CSVs"], label_visibility="collapsed")


# ---------------------------------------------------------------------------
# Upload
# ---------------------------------------------------------------------------
def upload_page() -> None:
    st.title("Upload CSVs")
    files = st.file_uploader("Slip Summary CSV files", type=["csv"], accept_multiple_files=True)
    replace = st.checkbox("Replace months that are already loaded")
    if not files:
        return

    plan = []
    for f in files:
        raw = f.getvalue()
        md5 = db.md5_of(raw)
        row = {"File": f.name, "Month": "", "Status": "", "_md5": md5, "_ok": False}
        try:
            prev = db.find_md5(engine, md5)
            kind = db.detect_report_type(raw)
            if prev:
                row["Status"] = "Already uploaded"
            elif kind != "slip_summary":
                row["Status"] = "Not a Slip Summary file"
            else:
                df, month = db.parse_slip_summary(raw)
                row.update(Month=f"{month:%b %Y}", _df=df, _month=month)
                if any(p.get("_month") == month for p in plan):
                    row["Status"] = "Same month as another file"
                elif db.month_loaded(engine, month) and not replace:
                    row["Status"] = "Month already loaded"
                else:
                    row["Status"] = "Ready"
                    row["_ok"] = True
        except Exception as e:  # noqa: BLE001
            row["Status"] = str(e)
        plan.append(row)

    st.dataframe(pd.DataFrame(plan)[["File", "Month", "Status"]], hide_index=True, width="stretch")
    ready = [p for p in plan if p["_ok"]]
    if st.button("Save", type="primary", disabled=not ready):
        for p in ready:
            try:
                db.save_slip_summary(engine, p["File"], p["_md5"], p["_df"], p["_month"], replace)
                st.success(f"{p['Month']} saved")
            except Exception as e:  # noqa: BLE001
                st.error(f"{p['File']}: {e}")
        st.cache_data.clear()


# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------
ALL_BRANCHES = "All branches"
ALL_CASHIERS = "All cashiers"


def dashboard_page() -> None:
    data = cached_slips(db.data_version(engine))
    st.title("Cashier Bet Slips")
    if data.empty:
        st.info("No data yet — upload CSVs first.")
        return

    months = sorted(data["month"].unique())
    first, last = pd.Timestamp(months[0]), pd.Timestamp(months[-1])
    total_label = f"Total ({len(months)} months)"

    # Branch (or all branches)
    options = [ALL_BRANCHES] + sorted(data["shop"].unique())
    if st.session_state.get("branch") not in options:
        st.session_state["branch"] = ALL_BRANCHES
    for col, b in zip(st.columns(len(options)), options):
        if col.button(b, key=f"branch_{b}", width="stretch",
                      type="primary" if st.session_state["branch"] == b else "secondary"):
            st.session_state["branch"] = b
            st.rerun()
    branch = st.session_state["branch"]
    all_branches = branch == ALL_BRANCHES
    bdf = data if all_branches else data[data["shop"] == branch]

    # Month
    month = st.selectbox("Month", months, format_func=lambda m: pd.Timestamp(m).strftime("%B %Y"))
    month_label = pd.Timestamp(month).strftime("%b %Y")
    num = {month_label: st.column_config.NumberColumn(format="%d"), total_label: st.column_config.NumberColumn(format="%d")}
    names = {"shop": "Branch", "cashier": "Cashier", "game": "Game"}

    # 1. Cashiers: selected month + all-months total
    keys = ["shop", "cashier"] if all_branches else ["cashier"]
    st.subheader(f"Cashiers — {branch} — {month_label}")
    cashiers = db.month_and_total(bdf, keys, month, month_label, total_label)
    if all_branches:  # keep each branch together
        cashiers = cashiers.sort_values(["shop", month_label], ascending=[True, False])
    cashiers = cashiers.rename(columns=names)
    st.dataframe(cashiers, hide_index=True, width="stretch", column_config=num)

    # 2. Bet slips per game for the chosen cashier (or all cashiers)
    st.subheader(f"Bet slips per game — {month_label}")
    if all_branches:
        people = [ALL_CASHIERS] + [f"{c} ({b})" for b, c in zip(cashiers["Branch"], cashiers["Cashier"])]
    else:
        people = [ALL_CASHIERS] + list(cashiers["Cashier"])
    who = st.selectbox("Cashier", people, key=f"who_{branch}")
    if who == ALL_CASHIERS:
        gdf = bdf
    else:
        idx = people.index(who) - 1
        gdf = bdf[bdf["cashier"] == cashiers["Cashier"].iloc[idx]]
        if all_branches:
            gdf = gdf[gdf["shop"] == cashiers["Branch"].iloc[idx]]
    games = db.month_and_total(gdf, ["game"], month, month_label, total_label).rename(columns=names)
    st.dataframe(games, hide_index=True, width="stretch", column_config=num)

    # Download
    st.divider()
    st.download_button(
        "Download Excel",
        data=db.build_report(bdf, month, month_label, total_label, branch, f"{first:%b %Y} – {last:%b %Y}"),
        file_name=f"{branch.replace(' ', '_')}_{pd.Timestamp(month):%b%Y}_bet_slips.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


if page == "Dashboard":
    dashboard_page()
else:
    upload_page()
