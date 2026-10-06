"""
Proposal Submission Pipeline - team tracker (Streamlit + SQLite)

Run:              streamlit run app.py
Requirements:     pip install "streamlit>=1.50" openpyxl pandas

Data is stored in a local SQLite file (proposals.db, next to this script).
Everyone who can reach the app can VIEW the tracker. Only people who enter the
shared password can add / edit / delete proposals. Set the password in
.streamlit/secrets.toml  (edit_password = "...")  or the EDIT_PASSWORD env var.
"""
from __future__ import annotations

import hmac
import io
import os
import sqlite3
import time
from contextlib import closing
import uuid
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# Settings - edit these lists to add a person / status / submission mode
# ─────────────────────────────────────────────────────────────────────────────
APP_TITLE = "Proposal Submission Pipeline"
TIMEZONE = "Asia/Karachi"          # used for "today" / overdue calculations
SOON_DAYS = 3                      # deadlines within this many days turn amber

PEOPLE = [
    "Irfan", "Ali Waheed", "Sajida", "Ayesha", "Zunaira",
    "Shahid Minhas", "Maryam", "Amina", "Sajjad Haider", "Daniyal", "Rana Nazir"
]
MODES = [
    "Email only",
    "Email + Hard copy",
    "Online/Portal only",
    "Online/Portal + Hard copy",
    "Hard copy only",
    "Email + Online/Portal",
    "Email + Online/Portal + Hard copy",
]
STATUSES = ["Not Started Yet", "In Progress", "Submitted", "Not Submitted", "Declined", "Accepted"]
CLOSED_STATUSES = {"Submitted", "Not Submitted"}   # no longer "open" work

COLUMNS = ["ID", "Proposal", "Deadline", "Submission Mode", "Assigned to", "Status"]
TEXT_COLUMNS = ["ID", "Proposal", "Submission Mode", "Assigned to", "Status"]

# The database file. Override with the PROPOSALS_DB environment variable if you
# want it somewhere else (e.g. a OneDrive/Dropbox folder for automatic backup).
DB_FILE = Path(os.environ.get("PROPOSALS_DB") or Path(__file__).with_name("proposals.db"))
# If present, imported once when the database is first created.
SEED_FILE = Path(__file__).with_name("seed_proposals.csv")

# Row colour by Status. Semi-transparent so they look right in both light and dark themes.
# A status that isn't listed here (e.g. "Accepted") gets no colour.
ROW_COLORS = {
    "Submitted": "rgba(34, 197, 94, 0.16)",       # green
    "In Progress": "rgba(245, 158, 11, 0.24)",    # amber
    "Not Started Yet": "rgba(56, 189, 248, 0.22)",  # light blue
    "Declined": "rgba(220, 38, 38, 0.20)",        # red
    "Not Submitted": "rgba(220, 38, 38, 0.20)",   # red
}
# Same colours for the Excel export (hex, no #).
EXCEL_FILLS = {
    "Submitted": "D5F0DC",
    "In Progress": "FCE5B0",
    "Not Started Yet": "D6EEFB",
    "Declined": "F8CFCF",
    "Not Submitted": "F8CFCF",
}


# ─────────────────────────────────────────────────────────────────────────────
# Small helpers
# ─────────────────────────────────────────────────────────────────────────────
def today_local() -> date:
    try:
        return datetime.now(ZoneInfo(TIMEZONE)).date()
    except Exception:  # tzdata missing - fall back to the machine's date
        return date.today()


def _secret(name: str, default=None):
    """Read from Streamlit secrets, falling back to an environment variable."""
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return os.environ.get(name.upper(), default)


def _width_kwargs() -> dict:
    """st.dataframe 'stretch' syntax differs between Streamlit versions."""
    try:
        major, minor = (int(x) for x in st.__version__.split(".")[:2])
        if (major, minor) >= (1, 50):
            return {"width": "stretch"}
    except Exception:
        pass
    return {"use_container_width": True}


def new_id() -> str:
    return "p" + uuid.uuid4().hex[:8]     # leading letter so Sheets never reads it as a number


# ─────────────────────────────────────────────────────────────────────────────
# Storage  (SQLite)
# ─────────────────────────────────────────────────────────────────────────────
def normalise(raw: pd.DataFrame) -> pd.DataFrame:
    """Coerce whatever came back from storage into a clean, typed frame."""
    df = raw.copy() if raw is not None else pd.DataFrame()
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUMNS].dropna(how="all")
    for col in TEXT_COLUMNS:
        df[col] = df[col].fillna("").astype(str).str.strip()
    df["Deadline"] = pd.to_datetime(df["Deadline"], errors="coerce", format="mixed")
    df = df[df["Proposal"] != ""].copy()
    missing = df["ID"] == ""
    df.loc[missing, "ID"] = [new_id() for _ in range(int(missing.sum()))]
    return df.reset_index(drop=True)


def _connect() -> sqlite3.Connection:
    con = sqlite3.connect(DB_FILE, timeout=15)   # wait if a teammate is mid-write
    con.execute("PRAGMA journal_mode=WAL")       # readers don't block the writer
    return con


def _write(sql: str, params: tuple) -> None:
    with closing(_connect()) as con:
        with con:                                 # commits on success
            con.execute(sql, params)


def init_db() -> None:
    """Create the table; on the very first run, import seed_proposals.csv if it exists."""
    first_run = not DB_FILE.exists()
    with closing(_connect()) as con:
        with con:
            con.execute(
                """CREATE TABLE IF NOT EXISTS proposals (
                       id TEXT PRIMARY KEY,
                       proposal TEXT NOT NULL,
                       deadline TEXT,
                       submission_mode TEXT,
                       assigned_to TEXT,
                       status TEXT
                   )"""
            )
    if first_run and SEED_FILE.exists():
        for rec in normalise(pd.read_csv(SEED_FILE, dtype=str)).to_dict("records"):
            upsert_proposal(rec)


def load_data() -> pd.DataFrame:
    init_db()
    with closing(_connect()) as con:
        raw = pd.read_sql_query(
            'SELECT id AS ID, proposal AS Proposal, deadline AS Deadline, '
            'submission_mode AS "Submission Mode", assigned_to AS "Assigned to", '
            "status AS Status FROM proposals",
            con,
        )
    return normalise(raw)


def upsert_proposal(record: dict) -> None:
    """Insert, or update if the ID already exists. Only touches that one row."""
    deadline = pd.Timestamp(record["Deadline"])
    deadline = "" if pd.isna(deadline) else deadline.strftime("%Y-%m-%d")
    _write(
        """INSERT INTO proposals (id, proposal, deadline, submission_mode, assigned_to, status)
           VALUES (?, ?, ?, ?, ?, ?)
           ON CONFLICT(id) DO UPDATE SET
               proposal = excluded.proposal, deadline = excluded.deadline,
               submission_mode = excluded.submission_mode,
               assigned_to = excluded.assigned_to, status = excluded.status""",
        (
            record.get("ID") or new_id(), record["Proposal"], deadline,
            record["Submission Mode"], record["Assigned to"], record["Status"],
        ),
    )


def delete_proposal(pid: str) -> None:
    _write("DELETE FROM proposals WHERE id = ?", (pid,))


# ─────────────────────────────────────────────────────────────────────────────
# Presentation logic
# ─────────────────────────────────────────────────────────────────────────────
def decorate(df: pd.DataFrame, today: date) -> pd.DataFrame:
    """Add 'Time left' text and an internal urgency key, and sort open work first."""
    d = df.copy()
    days = (d["Deadline"].dt.normalize() - pd.Timestamp(today)).dt.days

    labels, urgency = [], []
    for status, n in zip(d["Status"], days):
        if status == "Submitted":
            labels.append("")
            urgency.append("submitted")
        elif status == "Not Submitted":
            labels.append("")
            urgency.append("missed")
        elif pd.isna(n):
            labels.append("")
            urgency.append("ok")
        elif n < 0:
            labels.append(f"{-int(n)} day{'s' if n != -1 else ''} overdue")
            urgency.append("overdue")
        elif n == 0:
            labels.append("Due today")
            urgency.append("soon")
        elif n == 1:
            labels.append("Tomorrow")
            urgency.append("soon")
        else:
            labels.append(f"{int(n)} days")
            urgency.append("soon" if n <= SOON_DAYS else "ok")

    d["Time left"] = labels
    d["_urgency"] = urgency
    d["_days"] = days

    closed = d["Status"].isin(CLOSED_STATUSES)
    open_rows = d[~closed].sort_values("Deadline", na_position="last")
    closed_rows = d[closed].sort_values("Deadline", ascending=False, na_position="last")
    return pd.concat([open_rows, closed_rows], ignore_index=True)


def split_people(value: str) -> list[str]:
    return [p.strip() for p in str(value).split(",") if p.strip()]


def apply_filters(d: pd.DataFrame, search, statuses, people, open_only, window) -> pd.DataFrame:
    out = d
    if search:
        out = out[out["Proposal"].str.contains(search, case=False, regex=False)]
    if statuses:
        out = out[out["Status"].isin(statuses)]
    if people:
        out = out[out["Assigned to"].map(lambda v: any(p in split_people(v) for p in people))]
    if open_only:
        out = out[~out["Status"].isin(CLOSED_STATUSES)]
    if window == "Overdue":
        out = out[out["_urgency"] == "overdue"]
    elif window == "Next 7 days":
        out = out[(~out["Status"].isin(CLOSED_STATUSES)) & out["_days"].between(0, 7)]
    elif window == "This month":
        t = pd.Timestamp(today_local())
        out = out[(out["Deadline"].dt.year == t.year) & (out["Deadline"].dt.month == t.month)]
    return out.reset_index(drop=True)


def to_excel_bytes(d: pd.DataFrame) -> bytes:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    cols = ["Proposal", "Deadline", "Submission Mode", "Assigned to", "Status"]
    out = d[cols].copy()
    out["Deadline"] = out["Deadline"].dt.date

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        out.to_excel(xw, sheet_name="Proposal Pipeline", index=False)
        ws = xw.sheets["Proposal Pipeline"]

        head_fill = PatternFill("solid", start_color="1F3A5F")
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = head_fill
            cell.alignment = Alignment(vertical="center")
        widths = {"A": 52, "B": 12, "C": 32, "D": 34, "E": 16}
        for letter, w in widths.items():
            ws.column_dimensions[letter].width = w
        for i, status in enumerate(d["Status"], start=2):
            ws.cell(row=i, column=2).number_format = "d-mmm-yy"
            fill = EXCEL_FILLS.get(status)
            if fill:
                for c in range(1, len(cols) + 1):
                    ws.cell(row=i, column=c).fill = PatternFill("solid", start_color=fill)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(len(out) + 1, 2)}"
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# Access control
# ─────────────────────────────────────────────────────────────────────────────
def is_editor() -> bool:
    return bool(st.session_state.get("is_editor"))


def password_ok(attempt: str) -> bool:
    expected = _secret("edit_password")
    if not expected:
        return False
    return hmac.compare_digest(attempt.encode(), str(expected).encode())


def sidebar_access() -> None:
    with st.sidebar:
        st.subheader("Access")
        if is_editor():
            st.success("Editing unlocked")
            if st.button("Lock"):
                st.session_state["is_editor"] = False
                st.rerun()
            return

        st.caption("View-only. Enter the team password to add or edit proposals.")
        if not _secret("edit_password"):
            st.warning("No `edit_password` secret is set, so editing is disabled.")
            return
        with st.form("unlock"):
            attempt = st.text_input("Password", type="password")
            go = st.form_submit_button("Unlock editing")
        if go:
            if password_ok(attempt):
                st.session_state["is_editor"] = True
                st.rerun()
            else:
                time.sleep(1)      # slow down guessing
                st.error("Wrong password.")


# ─────────────────────────────────────────────────────────────────────────────
# Dialogs (add / edit / delete)
# ─────────────────────────────────────────────────────────────────────────────
def _options(base: list[str], current: list[str]) -> list[str]:
    """Keep legacy values (e.g. someone no longer in PEOPLE) selectable."""
    return base + [c for c in current if c and c not in base]


def _bump_table() -> None:
    st.session_state["table_version"] = st.session_state.get("table_version", 0) + 1


def _proposal_form(record: dict | None, form_key: str) -> None:
    rec = record or {}
    current_people = split_people(rec.get("Assigned to", ""))
    deadline_default = rec.get("Deadline")
    if deadline_default is None or pd.isna(deadline_default):
        deadline_default = today_local()
    deadline_default = pd.Timestamp(deadline_default).date()

    mode_opts = _options(MODES, [rec.get("Submission Mode", "")])
    status_opts = _options(STATUSES, [rec.get("Status", "")])
    mode_idx = mode_opts.index(rec["Submission Mode"]) if rec.get("Submission Mode") in mode_opts else None
    status_val = rec.get("Status") or "Not Started Yet"

    with st.form(form_key):
        name = st.text_input("Proposal", value=rec.get("Proposal", ""))
        deadline = st.date_input("Deadline", value=deadline_default, format="DD/MM/YYYY")
        mode = st.selectbox("Submission mode", mode_opts, index=mode_idx, placeholder="Choose a mode")
        people = st.multiselect(
            "Assigned to", _options(PEOPLE, current_people), default=current_people,
            placeholder="Pick one or more people",
        )
        status = st.selectbox("Status", status_opts, index=status_opts.index(status_val))
        saved = st.form_submit_button("Save", type="primary")

    if not saved:
        return
    if not name.strip():
        st.error("Please enter the proposal name.")
        return
    if not mode:
        st.error("Please choose a submission mode.")
        return
    try:
        upsert_proposal({
            "ID": rec.get("ID", ""),
            "Proposal": name.strip(),
            "Deadline": deadline,
            "Submission Mode": mode,
            "Assigned to": ", ".join(people),
            "Status": status,
        })
    except Exception as exc:
        st.error(f"Could not save: {exc}")
        return
    _bump_table()
    st.rerun()


@st.dialog("Add proposal")
def add_dialog() -> None:
    _proposal_form(None, "add_form")


@st.dialog("Edit proposal")
def edit_dialog(record: dict) -> None:
    _proposal_form(record, "edit_form")


@st.dialog("Delete proposal")
def delete_dialog(pid: str, name: str) -> None:
    st.write(f"Delete **{name}**? This can't be undone.")
    c1, c2 = st.columns(2)
    if c1.button("Yes, delete", type="primary"):
        try:
            delete_proposal(pid)
        except Exception as exc:
            st.error(f"Could not delete: {exc}")
            return
        _bump_table()
        st.rerun()
    if c2.button("Cancel"):
        st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# Page
# ─────────────────────────────────────────────────────────────────────────────
def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📋", layout="wide")
    sidebar_access()

    st.title(APP_TITLE)
    try:
        df = load_data()
    except Exception as exc:
        st.error(f"Couldn't open the database ({DB_FILE}): {exc}")
        st.stop()

    today = today_local()
    data = decorate(df, today)
    open_mask = ~data["Status"].isin(CLOSED_STATUSES)

    # Summary tiles
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Overdue", int((data["_urgency"] == "overdue").sum()))
    m2.metric("Due in next 7 days", int((open_mask & data["_days"].between(0, 7)).sum()))
    m3.metric("In progress", int((data["Status"] == "In Progress").sum()))
    m4.metric("Submitted", int((data["Status"] == "Submitted").sum()))

    # Filters
    f1, f2, f3, f4, f5 = st.columns([2.2, 1.6, 1.6, 1.4, 1.1])
    search = f1.text_input("Search proposals", placeholder="Type part of a name…")
    statuses = f2.multiselect("Status", _options(STATUSES, sorted(set(data["Status"]) - set(STATUSES))))
    all_people = _options(PEOPLE, sorted({p for v in data["Assigned to"] for p in split_people(v)}))
    people = f3.multiselect("Assigned to", all_people)
    window = f4.selectbox("Deadline", ["Any time", "Overdue", "Next 7 days", "This month"])
    open_only = f5.toggle("Open only", value=False, help="Hide Submitted / Not Submitted")

    view = apply_filters(data, search, statuses, people, open_only, window)

    # Table
    display_cols = ["Proposal", "Deadline", "Time left", "Submission Mode", "Assigned to", "Status"]
    styled = view.style.apply(
        lambda row: [f"background-color: {ROW_COLORS.get(row['Status'], '')}"] * len(row), axis=1
    )
    event = st.dataframe(
        styled,
        hide_index=True,
        column_order=display_cols,
        column_config={
            "Proposal": st.column_config.TextColumn("Proposal", width="large"),
            "Deadline": st.column_config.DateColumn("Deadline", format="DD MMM YYYY"),
            "Time left": st.column_config.TextColumn("Time left", width="small"),
        },
        on_select="rerun" if is_editor() else "ignore",
        selection_mode="single-row",
        key=f"table_{st.session_state.get('table_version', 0)}",
        **_width_kwargs(),
    )
    st.caption(
        f"{len(view)} of {len(data)} proposals · green = submitted · amber = in progress "
        "· light blue = not started yet · red = declined / not submitted"
    )

    # Actions
    selected = None
    if is_editor():
        rows = event.selection.rows if event is not None else []
        if rows and rows[0] < len(view):
            selected = view.iloc[rows[0]].to_dict()

    a1, a2, a3, _, a5 = st.columns([1.1, 1.1, 1.1, 2.2, 1.3])
    if is_editor():
        if a1.button("➕ Add proposal", type="primary"):
            add_dialog()
        if a2.button("✏️ Edit selected", disabled=selected is None):
            edit_dialog(selected)
        if a3.button("🗑️ Delete", disabled=selected is None):
            delete_dialog(selected["ID"], selected["Proposal"])
        if selected is None:
            st.caption("Tip: click a row's checkbox (left edge of the table) to edit or delete it.")
    else:
        st.caption("🔒 View-only — unlock editing from the sidebar with the team password.")

    a5.download_button(
        "⬇️ Export to Excel",
        data=to_excel_bytes(view),
        file_name=f"Proposal_Submission_Pipeline_{today:%Y-%m-%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Exports the rows currently shown (respecting your filters).",
    )


if __name__ == "__main__":
    main()
