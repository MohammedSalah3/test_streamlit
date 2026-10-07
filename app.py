import html as _html
import re
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Hassan Allam Construction | Projects Dashboard",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
DATA_FOLDER = BASE_DIR / "data"

OVERVIEW = "__overview__"
APP_VERSION = "v11 - new UI"

SHEET_ICONS = {
    "unit_prices": "💰",
    "pr": "📝",
    "pr_empty": "⚪",
    "total": "📊",
    "generic": "📄",
    "error": "⚠️",
}

SHEET_TYPE_LABELS = {
    "unit_prices": "Unit prices list",
    "pr": "PR / PO sheet",
    "pr_empty": "Empty PR template",
    "total": "Totals sheet",
    "generic": "Generic sheet",
    "error": "Error",
}

st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&display=swap');

        :root {
            --ink: #1F2933;
            --muted: #667085;
            --line: #E4E7EC;
            --surface: #FFFFFF;
            --canvas: #F5F7FA;
            --primary: #1D4E89;
            --good: #2E7D5B;
            --warn: #B7791F;
            --bad: #B83232;
        }

        .stApp { background: var(--canvas); }

        .stApp, .stApp p, .stApp label, .stApp button, .stApp input,
        .stApp textarea, .stApp h1, .stApp h2, .stApp h3, .stApp h4,
        .stApp h5, .stMarkdown {
            font-family: 'IBM Plex Sans', 'Segoe UI', system-ui, sans-serif;
        }

        .block-container {
            padding-top: 2.4rem;
            padding-bottom: 3rem;
            max-width: 1400px;
        }

        /* ---------- header ---------- */
        .app-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
            gap: 16px;
            flex-wrap: wrap;
            padding-bottom: 16px;
            margin-bottom: 18px;
            border-bottom: 1px solid var(--line);
        }
        .brand { font-size: 13px; color: var(--muted); font-weight: 500; }
        .file-title {
            font-size: 28px; font-weight: 600; color: var(--ink);
            line-height: 1.2; margin-top: 4px;
        }
        .chips { display: flex; gap: 8px; flex-wrap: wrap; }
        .chip {
            font-size: 12.5px; padding: 4px 11px; border-radius: 999px;
            background: #EEF2F7; color: #344054; border: 1px solid var(--line);
            white-space: nowrap;
        }
        .chip.accent { background: #E6EEF8; color: var(--primary); border-color: #C9D9EE; }

        /* ---------- section titles ---------- */
        .section-title {
            font-size: 15px; font-weight: 600; color: var(--ink);
            margin: 18px 0 8px;
        }

        /* ---------- KPI cards ---------- */
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(175px, 1fr));
            gap: 12px;
            margin: 4px 0 18px;
        }
        .kpi {
            background: var(--surface); border: 1px solid var(--line);
            border-radius: 10px; padding: 14px 16px;
        }
        .kpi .label { font-size: 12.5px; color: var(--muted); }
        .kpi .value {
            font-size: 26px; font-weight: 600; color: var(--ink);
            font-variant-numeric: tabular-nums; line-height: 1.25;
        }
        .kpi .sub { font-size: 12px; color: var(--muted); margin-top: 2px; }
        .kpi.good .value { color: var(--good); }
        .kpi.warn .value { color: var(--warn); }
        .kpi.bad .value { color: var(--bad); }

        /* ---------- progress bar ---------- */
        .bar {
            height: 8px; background: #E8ECF1; border-radius: 999px;
            overflow: hidden; margin-top: 10px;
        }
        .bar > span {
            display: block; height: 100%; border-radius: 999px;
            background: var(--primary);
        }
        .bar.done > span { background: var(--good); }
        .bar.over > span { background: var(--bad); }

        /* ---------- PO cards (overview) ---------- */
        div[class*="st-key-pocard_"] {
            background: var(--surface); border: 1px solid var(--line);
            border-radius: 10px; padding: 16px 16px 12px;
        }
        .po-head { display: flex; justify-content: space-between;
                   align-items: flex-start; gap: 8px; }
        .po-title { font-size: 17px; font-weight: 600; color: var(--ink); }
        .po-sub { font-size: 12.5px; color: var(--muted); margin-top: 2px; }
        .po-nums {
            display: grid; grid-template-columns: repeat(3, 1fr);
            gap: 8px; margin: 12px 0 4px;
        }
        .po-nums span { display: block; font-size: 11.5px; color: var(--muted); }
        .po-nums b {
            font-size: 16px; font-weight: 600; color: var(--ink);
            font-variant-numeric: tabular-nums;
        }
        .po-pct { font-size: 12px; color: var(--muted); margin-top: 6px; }

        /* ---------- empty state ---------- */
        .empty-state {
            background: var(--surface); border: 1px dashed #C5CCD6;
            border-radius: 10px; padding: 28px 22px; color: var(--muted);
        }
        .empty-state b { color: var(--ink); font-size: 16px; }

        /* ---------- sidebar ---------- */
        section[data-testid="stSidebar"] {
            background: var(--surface); border-right: 1px solid var(--line);
        }
        .side-brand { font-size: 17px; font-weight: 600; color: var(--ink);
                      line-height: 1.25; }
        .side-brand span { display: block; font-size: 12.5px;
                           font-weight: 400; color: var(--muted); margin-top: 2px; }
        .side-label { font-size: 12.5px; font-weight: 600; color: var(--muted);
                      margin: 6px 0 4px; }

        /* ---------- widgets ---------- */
        button[data-baseweb="tab"] { font-weight: 500; }
        div[data-testid="stExpander"] {
            background: var(--surface); border-radius: 10px;
        }

        /* Hide GitHub icon, Deploy button, main menu and footer.
           NOTE: stToolbar itself must stay visible - it holds the
           button that re-opens the sidebar menu. */
        [data-testid="stToolbarActions"],
        [data-testid="stAppDeployButton"],
        [data-testid="stMainMenu"],
        [data-testid="stDecoration"],
        [data-testid="stStatusWidget"],
        .stAppDeployButton,
        #MainMenu,
        footer {
            visibility: hidden !important;
            display: none !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# Dark mode: inverts the light theme (works for tables and charts too).
# Requires the base theme to be "light" (see .streamlit/config.toml).
DARK_CSS = """
<style>
    .stApp {
        filter: invert(0.9) hue-rotate(180deg);
    }
</style>
"""


# =========================================================
# SMALL HELPERS
# =========================================================

EXCEL_ERRORS = {
    "#N/A", "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#NUM!", "#NULL!",
}


def clean_value(value):
    """Strip text and turn Excel error values (#N/A, #REF! ...) into None."""
    if isinstance(value, str):
        text = value.strip()
        if text == "" or text in EXCEL_ERRORS:
            return None
        return text
    return value


def fmt(value):
    """Convert any cell value to a display-friendly string."""
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        if pd.isna(value):
            return ""
        if value.is_integer():
            return str(int(value))
    return str(value)


def is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def g(grid, r, c):
    """Safe cell access (0-based)."""
    if r is None or c is None or r < 0 or r >= len(grid):
        return None
    row = grid[r]
    return row[c] if 0 <= c < len(row) else None


def unique_names(names):
    seen = {}
    result = []
    for name in names:
        if name in seen:
            seen[name] += 1
            result.append(f"{name}_{seen[name]}")
        else:
            seen[name] = 1
            result.append(name)
    return result


def to_num(series):
    return pd.to_numeric(series, errors="coerce")


def arrow_safe(df):
    """Make mixed-type columns safe for st.dataframe."""
    df = df.copy()
    for col in df.columns:
        if df[col].dtype == object:
            converted = to_num(df[col])
            if converted.notna().sum() == df[col].notna().sum():
                df[col] = converted
            else:
                df[col] = df[col].map(fmt)
    return df


def find_row(grid, text, limit=30):
    """First row (0-based) that has a cell equal to `text` (case-insensitive)."""
    target = text.upper()
    for r in range(min(limit, len(grid))):
        for value in grid[r]:
            if isinstance(value, str) and value.strip().upper() == target:
                return r
    return None


def find_label_row(grid, prefixes, limit):
    for r in range(limit):
        for value in grid[r]:
            if (
                isinstance(value, str)
                and value.strip().upper().startswith(prefixes)
            ):
                return r
    return None


# =========================================================
# READ WORKSHEET INTO A GRID (ONE FAST PASS)
# =========================================================

def read_grid(ws):
    grid = [
        [clean_value(v) for v in row]
        for row in ws.iter_rows(values_only=True)
    ]

    while grid and all(v is None for v in grid[-1]):
        grid.pop()

    width = max((len(r) for r in grid), default=0)

    return [row + [None] * (width - len(row)) for row in grid]


# =========================================================
# SHEET TYPE DETECTION
# =========================================================

def detect_sheet_type(grid):
    """
    Decide how a sheet should be parsed by looking at its content.
    To support a new structure: add a detector here + a parser below.
    """
    head = {
        str(v).strip().upper()
        for row in grid[:15]
        for v in row
        if v is not None
    }

    a1 = str(g(grid, 0, 0) or "").upper()
    a3 = str(g(grid, 2, 0) or "").upper()

    if "INTERNAL HIRE RATE" in head:
        return "unit_prices"

    if a3.startswith("PR# NO") or a1.startswith("PROJECT NAME"):
        return "pr"

    if {"ITEMS NAME", "REQ", "SEND", "REMAIN"} <= head:
        return "total"

    return "generic"


# =========================================================
# PARSER 1: UNIT PRICES
# =========================================================

def parse_unit_prices(grid):
    h = find_row(grid, "INTERNAL HIRE RATE")

    if h is None:
        return None

    header = unique_names(
        [
            str(v) if v is not None else f"Column {i + 1}"
            for i, v in enumerate(grid[h])
        ]
    )

    df = pd.DataFrame(grid[h + 1:], columns=header)
    df = df.dropna(how="all").dropna(axis=1, how="all")

    name_col = next(
        (c for c in df.columns if c.upper() == "ITEMS NAME"), None
    )
    if name_col:
        df = df[df[name_col].notna()]

    title_lines = [
        " ".join(fmt(v) for v in row if v is not None)
        for row in grid[:h]
        if any(v is not None for v in row)
    ]

    return {
        "type": "unit_prices",
        "table": df.reset_index(drop=True),
        "title_lines": title_lines,
    }


# =========================================================
# PARSER 2: PR / PO SHEET
# =========================================================

def parse_pr(grid):
    h = find_row(grid, "ITEMS NAME")

    if h is None:
        return None

    up = [
        str(v).strip().upper() if v is not None else ""
        for v in grid[h]
    ]

    def col_of(*names):
        for name in names:
            if name in up:
                return up.index(name)
        return None

    c_no = col_of("NO") if col_of("NO") is not None else 0
    c_code = col_of("ASSET CODE")
    c_serial = col_of("FULL SERIAL")
    c_name = col_of("ITEMS NAME")
    c_req = col_of("REQ QTY")
    c_send = col_of("SEND QTY")
    c_remain = col_of("REMAIN QTY")
    c_remarks = col_of("REMARKS")
    c_return = col_of("T.RETURN", "T. RETURN")

    if c_req is None or c_send is None:
        return None

    # ---------- header information (rows above the table) ----------

    info = {}
    categories = []

    for r in range(h):
        label = g(grid, r, 0)

        if isinstance(label, str) and label.endswith(":"):
            value = next(
                (v for v in grid[r][1:3] if v is not None), None
            )
            info[label.rstrip(":").strip()] = value

        cat_label = g(grid, r, 3)
        cat_flag = g(grid, r, 4)

        if (
            isinstance(cat_label, str)
            and is_number(cat_flag)
            and cat_flag
        ):
            categories.append(cat_label)

    # ---------- shipment columns ----------

    date_row = find_label_row(grid, ("RECIVED", "RECEIVED"), h)
    voucher_row = find_label_row(grid, ("STORE VOUCHER",), h)
    site_row = h - 1 if (h - 1) not in (date_row, voucher_row) else None

    ship_cols = [
        c
        for c in range(c_req + 1, c_send)
        if g(grid, h, c) is not None or g(grid, date_row, c) is not None
    ]

    # ---------- items ----------

    items = []
    matrix_rows = []

    for r in range(h + 1, len(grid)):
        name = g(grid, r, c_name)

        if name is None:
            continue

        items.append(
            {
                "No": g(grid, r, c_no),
                "Asset Code": g(grid, r, c_code),
                "Full Serial": g(grid, r, c_serial),
                "Item": name,
                "Req Qty": g(grid, r, c_req),
                "Send Qty": g(grid, r, c_send),
                "Remain Qty": g(grid, r, c_remain),
                "T.Return": g(grid, r, c_return),
                "Remarks": g(grid, r, c_remarks),
            }
        )
        matrix_rows.append([g(grid, r, c) for c in ship_cols])

    items_df = pd.DataFrame(items)

    ship_names = unique_names(
        [
            f"S{i + 1} · {fmt(g(grid, date_row, c)) or 'no date'}"
            for i, c in enumerate(ship_cols)
        ]
    )

    matrix_df = pd.DataFrame(matrix_rows, columns=ship_names)

    empty_result = {
        "type": "pr_empty",
        "info": info,
        "categories": categories,
    }

    if items_df.empty:
        return empty_result

    for col in ["Req Qty", "Send Qty", "Remain Qty", "T.Return"]:
        items_df[col] = to_num(items_df[col])

    items_df["Remain Qty"] = items_df["Remain Qty"].fillna(
        items_df["Req Qty"].fillna(0) - items_df["Send Qty"].fillna(0)
    )

    for col in ["Req Qty", "Send Qty", "T.Return"]:
        items_df[col] = items_df[col].fillna(0)

    for col in matrix_df.columns:
        matrix_df[col] = to_num(matrix_df[col]).fillna(0)

    active = (
        (items_df["Req Qty"] != 0)
        | (items_df["Send Qty"] != 0)
        | (matrix_df.abs().sum(axis=1) != 0 if len(matrix_df.columns) else False)
    )

    if not active.any():
        return empty_result

    def status(row):
        if row["Remain Qty"] < 0:
            return "Over-sent"
        if row["Req Qty"] > 0 and row["Remain Qty"] == 0:
            return "Completed"
        if row["Send Qty"] == 0:
            return "Not started"
        return "Partial"

    items_df["Status"] = items_df.apply(status, axis=1)
    items_df["Active"] = active.values

    items_df = items_df[
        [
            "No", "Asset Code", "Item", "Req Qty", "Send Qty",
            "Remain Qty", "T.Return", "Status", "Remarks",
            "Full Serial", "Active",
        ]
    ]

    # ---------- shipments table ----------

    shipments = []

    for i, c in enumerate(ship_cols):
        shipments.append(
            {
                "#": f"S{i + 1}",
                "Supplier": g(grid, h, c),
                "Date": g(grid, date_row, c),
                "Store Voucher": g(grid, voucher_row, c),
                "Site / Ref": g(grid, site_row, c),
                "Total Qty": matrix_df[ship_names[i]].sum(),
                "Items": int((matrix_df[ship_names[i]] != 0).sum()),
            }
        )

    shipments_df = pd.DataFrame(shipments)

    supplier = next(
        (s["Supplier"] for s in shipments if s["Supplier"]), None
    )

    act = items_df[items_df["Active"]]

    req_total = float(act["Req Qty"].sum())
    send_total = float(act["Send Qty"].sum())

    summary = {
        "PR No": info.get("PR# No"),
        "Date": info.get("PR# Date"),
        "Category": ", ".join(categories) or None,
        "Supplier": supplier,
        "Items": int(len(act)),
        "Required": req_total,
        "Sent": send_total,
        "Remaining": float(act["Remain Qty"].sum()),
        "Sent %": (send_total / req_total) if req_total else None,
    }

    matrix_full = pd.concat(
        [
            items_df[["No", "Asset Code", "Item"]].reset_index(drop=True),
            matrix_df.reset_index(drop=True),
        ],
        axis=1,
    )[items_df["Active"].values]

    return {
        "type": "pr",
        "info": info,
        "categories": categories,
        "items": items_df.reset_index(drop=True),
        "shipments": shipments_df,
        "matrix": matrix_full.reset_index(drop=True),
        "summary": summary,
    }


# =========================================================
# PARSER 3: TOTAL SHEET
# =========================================================

def parse_total(grid):
    h = find_row(grid, "ITEMS NAME")

    if h is None:
        return None

    header = grid[h]

    # group titles (SLAB / WALLS & COLUMNS ...) sit above the header,
    # skipping the "Area=" row
    group_row = None

    for r in range(h - 1, max(h - 4, -1), -1):
        if any(
            isinstance(v, str) and not v.endswith("=")
            for v in grid[r][4:]
        ):
            group_row = r
            break

    labels = {}
    current = ""

    for c in range(len(header)):
        value = g(grid, group_row, c)

        if value:
            current = str(value).strip()
            current = {"SCAFFOLDIN": "SCAFFOLDING"}.get(current.upper(), current)

        labels[c] = current if header[c] is not None else None

    names = {}

    for c, title in enumerate(header):
        if title is None:
            continue
        label = labels[c]
        names[c] = f"{label} | {title}" if (label and c >= 4) else str(title)

    # Area per group (from the "Area=" row)
    areas = {}
    for c in range(len(header)):
        v = g(grid, h - 1, c)
        if isinstance(v, str) and v.upper().startswith("AREA") and labels.get(c):
            areas[labels[c]] = g(grid, h - 1, c + 1)

    cols = list(names.keys())
    item_rows = []
    footer_rows = []

    for r in range(h + 1, len(grid)):
        row = {names[c]: g(grid, r, c) for c in cols}

        if is_number(g(grid, r, 0)):
            item_rows.append(row)
        elif isinstance(g(grid, r, 3), str):
            footer_rows.append(row)

    items = pd.DataFrame(item_rows)
    footer = pd.DataFrame(footer_rows)

    if items.empty:
        return None

    item_col = "ITEMS NAME"
    code_col = "ASSET CODE"

    keep = items[item_col].notna() | (
        items[code_col].notna() & (items[code_col] != 0)
    )
    items = items[keep].copy()
    items = items.rename(columns={item_col: "Item"})

    numeric_cols = [c for c in items.columns if " | " in c]

    for col in numeric_cols:
        items[col] = to_num(items[col]).fillna(0)

    groups = []
    for c in cols:
        if labels[c] and c >= 4 and labels[c] not in groups:
            groups.append(labels[c])

    # totals per group for the chart
    chart_rows = []

    for group in groups:
        def pick(*keys):
            for key in keys:
                name = f"{group} | {key}"
                if name in items.columns:
                    return float(items[name].sum())
            return None

        req = pick("REQ", "T. REQ")
        send = pick("SEND", "T. SEND")
        remain = pick("REMAIN", "T. REMAIN")

        chart_rows.append(
            {
                "Group": group,
                "Required": req or 0,
                "Sent": send or 0,
                "Remaining": remain or 0,
            }
        )

    group_totals = pd.DataFrame(chart_rows)

    # footer (cost totals)
    footer_clean = pd.DataFrame()

    if not footer.empty:
        footer_clean = footer.rename(columns={item_col: "Label"})
        keep_cols = ["Label"] + [
            c for c in footer_clean.columns if " | " in c
        ]
        footer_clean = footer_clean[keep_cols].copy()

        for col in keep_cols[1:]:
            footer_clean[col] = to_num(footer_clean[col])

        footer_clean = footer_clean.dropna(axis=1, how="all")

    return {
        "type": "total",
        "items": items.reset_index(drop=True),
        "footer": footer_clean,
        "groups": groups,
        "areas": areas,
        "group_totals": group_totals,
    }


# =========================================================
# GENERIC FALLBACK (UNKNOWN STRUCTURE)
# =========================================================

def parse_generic(grid):
    df = pd.DataFrame(grid)
    df = df.dropna(how="all").dropna(axis=1, how="all")
    return {"type": "generic", "table": df.reset_index(drop=True)}


PARSERS = {
    "unit_prices": parse_unit_prices,
    "pr": parse_pr,
    "total": parse_total,
}


def parse_sheet(ws):
    try:
        grid = read_grid(ws)

        if not grid:
            return {"type": "generic", "table": pd.DataFrame()}

        sheet_type = detect_sheet_type(grid)
        parser = PARSERS.get(sheet_type)

        result = parser(grid) if parser else None

        # PO number lives in the sheet name (e.g. "PO 94000171374 MASTER")
        if result and result["type"] == "pr":
            match = re.search(r"\d{8,}", ws.title)
            result["summary"] = {
                "PO No": match.group(0) if match else None,
                **result["summary"],
            }

        # parser could not understand the sheet -> safe fallback
        return result or parse_generic(grid)

    except Exception as e:
        return {"type": "error", "message": str(e)}


# =========================================================
# LOAD EXCEL FILE (CACHED, ONE PASS)
# =========================================================

@st.cache_data(show_spinner="Reading Excel file...")
def load_excel_file(file_path, modified_time):
    workbook = openpyxl.load_workbook(file_path, data_only=True)

    return {ws.title: parse_sheet(ws) for ws in workbook.worksheets}


# =========================================================
# FIND EXCEL FILES
# =========================================================

def get_excel_files():
    if not DATA_FOLDER.exists():
        return []

    files = [
        f
        for f in DATA_FOLDER.iterdir()
        if f.is_file()
        and f.suffix.lower() in [".xlsx", ".xlsm"]
        and not f.name.startswith("~$")
    ]

    return sorted(files, key=lambda x: x.name.lower())


# =========================================================
# UI HELPERS (HTML pieces)
# =========================================================

STATUS_ICONS = {
    "Completed": "🟢",
    "Partial": "🟡",
    "Over-sent": "🔴",
    "Not started": "⚪",
}

CHART_COLORS = ["#1D4E89", "#2E7D5B", "#B7791F"]


def esc(value):
    return _html.escape(value if isinstance(value, str) else fmt(value))


def bar_html(fraction):
    if fraction is None:
        return ""

    pct = max(0.0, min(fraction, 1.0)) * 100

    if fraction > 1.0001:
        cls = "over"
    elif fraction >= 0.9999:
        cls = "done"
    else:
        cls = ""

    return f'<div class="bar {cls}"><span style="width:{pct:.1f}%"></span></div>'


def kpi(label, value, sub="", tone="", bar=None):
    sub_html = f'<div class="sub">{esc(sub)}</div>' if sub else ""
    return (
        f'<div class="kpi {tone}">'
        f'<div class="label">{esc(label)}</div>'
        f'<div class="value">{esc(value)}</div>'
        f"{sub_html}{bar_html(bar)}"
        "</div>"
    )


def kpi_grid(cards):
    st.markdown(
        '<div class="kpi-grid">' + "".join(cards) + "</div>",
        unsafe_allow_html=True,
    )


def chips_html(items, accent_first=False):
    parts = []
    for i, item in enumerate(items):
        if item in (None, ""):
            continue
        cls = "chip accent" if (accent_first and i == 0) else "chip"
        parts.append(f'<span class="{cls}">{esc(item)}</span>')
    return f'<div class="chips">{"".join(parts)}</div>'


def section_title(text):
    st.markdown(
        f'<div class="section-title">{esc(text)}</div>', unsafe_allow_html=True
    )


def num(value):
    return f"{value:,.0f}"


def pct_text(value):
    return f"{value:.0%}" if value is not None and pd.notna(value) else "—"


# =========================================================
# VIEWS
# =========================================================

def show_table(df, height=600, column_config=None):
    st.dataframe(
        arrow_safe(df),
        width="stretch",
        hide_index=True,
        height=height,
        column_config=column_config,
    )


def text_filter(df, columns, key):
    query = st.text_input(
        "🔎 Search", key=key, placeholder="Item name or asset code..."
    )

    if not query:
        return df

    mask = pd.Series(False, index=df.index)

    for col in columns:
        if col in df.columns:
            mask |= (
                df[col]
                .astype(str)
                .str.contains(query, case=False, na=False, regex=False)
            )

    return df[mask]


def show_unit_prices(data, sheet):
    table = data["table"]

    st.markdown(
        chips_html([f"{len(table):,} items"] + data["title_lines"][-1:]),
        unsafe_allow_html=True,
    )

    df = text_filter(table, ["ITEMS NAME", "ASSET CODE"], f"s_{sheet}")

    st.caption(f"Showing {len(df):,} of {len(table):,} items")
    show_table(df)


def show_pr(data, sheet):
    info = data["info"]
    summary = data["summary"]
    items = data["items"]
    active = items[items["Active"]]

    po = summary.get("PO No")
    title = f"PO {po}" if po else sheet

    st.markdown(
        '<div class="po-head" style="margin-bottom:8px">'
        f'<div class="po-title" style="font-size:22px">{esc(title)}</div>'
        + chips_html(
            [
                summary["Category"],
                summary["Supplier"],
                f"PR {fmt(summary['PR No'])}" if summary["PR No"] else None,
                fmt(summary["Date"]) if summary["Date"] else None,
            ],
            accent_first=True,
        )
        + "</div>",
        unsafe_allow_html=True,
    )

    extra = [
        f"{k}: {fmt(v)}"
        for k, v in info.items()
        if k not in ("PR# No", "PR# Date") and v is not None
    ]
    if extra:
        st.markdown(chips_html(extra), unsafe_allow_html=True)

    over = int((active["Status"] == "Over-sent").sum())

    kpi_grid(
        [
            kpi("Required", num(summary["Required"]), f"{summary['Items']} items"),
            kpi("Sent", num(summary["Sent"]), tone="good"),
            kpi("Remaining", num(summary["Remaining"]), tone="warn"),
            kpi(
                "Sent / Required",
                pct_text(summary["Sent %"]),
                bar=summary["Sent %"],
            ),
            kpi(
                "Over-sent items",
                over,
                "Sent more than required" if over else "None",
                tone="bad" if over else "",
            ),
        ]
    )

    tab_items, tab_ship, tab_matrix = st.tabs(
        ["📦 Items", "🚚 Shipments", "🧮 Items × Shipments"]
    )

    with tab_items:
        f1, f2, f3 = st.columns([2, 2, 1.2])

        with f1:
            query = st.text_input(
                "🔎 Search",
                key=f"s_{sheet}",
                placeholder="Item name or asset code...",
            )
        with f2:
            statuses = st.multiselect(
                "Status",
                list(STATUS_ICONS),
                key=f"st_{sheet}",
                format_func=lambda x: f"{STATUS_ICONS[x]} {x}",
            )
        with f3:
            st.write("")
            only_active = st.toggle(
                "With quantities only", value=True, key=f"a_{sheet}"
            )

        df = active if only_active else items

        if statuses:
            df = df[df["Status"].isin(statuses)]

        if query:
            mask = pd.Series(False, index=df.index)
            for col in ("Item", "Asset Code"):
                mask |= (
                    df[col]
                    .astype(str)
                    .str.contains(query, case=False, na=False, regex=False)
                )
            df = df[mask]

        view = df.drop(columns=["Active", "Full Serial"]).copy()

        view["Sent %"] = [
            min(send / req, 1.0) * 100 if req > 0 else None
            for req, send in zip(view["Req Qty"], view["Send Qty"])
        ]
        view["Status"] = view["Status"].map(
            lambda x: f"{STATUS_ICONS.get(x, '')} {x}"
        )

        ordered = [
            "No", "Asset Code", "Item", "Req Qty", "Send Qty",
            "Remain Qty", "Sent %", "T.Return", "Status", "Remarks",
        ]
        if view["Remarks"].isna().all():
            ordered.remove("Remarks")
        else:
            view["Remarks"] = view["Remarks"].fillna("").astype(str)

        view = view[ordered]

        st.caption(f"{len(view):,} items")

        show_table(
            view,
            column_config={
                "Req Qty": st.column_config.NumberColumn("Required", format="localized"),
                "Send Qty": st.column_config.NumberColumn("Sent", format="localized"),
                "Remain Qty": st.column_config.NumberColumn("Remaining", format="localized"),
                "T.Return": st.column_config.NumberColumn("Returned", format="localized"),
                "Sent %": st.column_config.ProgressColumn(
                    "Progress", min_value=0, max_value=100, format="%d%%"
                ),
            },
        )

    with tab_ship:
        if data["shipments"].empty:
            st.info("No shipments recorded in this sheet.")
        else:
            show_table(data["shipments"], height=450)

    with tab_matrix:
        if data["matrix"].shape[1] <= 3:
            st.info("No shipments recorded in this sheet.")
        else:
            show_table(data["matrix"])


def show_total(data, sheet):
    footer = data["footer"]

    def footer_value(label_part, column_suffix):
        if footer.empty:
            return None

        rows = footer[footer["Label"].astype(str).str.upper().str.contains(label_part)]

        if rows.empty:
            return None

        for col in footer.columns:
            if col.endswith(column_suffix):
                value = rows.iloc[0][col]
                return None if pd.isna(value) else value

        return None

    purch = [footer_value("PURCH", s) for s in ("T. REQ", "T. SEND", "T. REMAIN")]
    ext = [footer_value("EXT HIRE", s) for s in ("T. REQ", "T. SEND", "T. REMAIN")]
    pct = footer_value("PERCENTAGE", "T. REQ")

    if any(v is not None for v in purch):
        section_title("Value summary (L.E)")

        cards = [
            kpi("Purchase value – Required", num(purch[0]) if purch[0] is not None else "—"),
            kpi("Purchase value – Sent", num(purch[1]) if purch[1] is not None else "—", tone="good"),
            kpi("Purchase value – Remaining", num(purch[2]) if purch[2] is not None else "—", tone="warn"),
            kpi(
                "Sent (by value)",
                f"{pct:.1%}" if pct is not None else "—",
                bar=pct,
            ),
        ]
        kpi_grid(cards)

        if any(v is not None for v in ext):
            kpi_grid(
                [
                    kpi(label, num(v) if v is not None else "—")
                    for label, v in zip(
                        ["Ext. hire – Required", "Ext. hire – Sent", "Ext. hire – Remaining"],
                        ext,
                    )
                ]
            )

    groups_df = data["group_totals"]
    groups_df = groups_df[groups_df["Group"] != "Total Summary"]

    if not groups_df.empty:
        section_title("Quantities by work group")
        st.bar_chart(
            groups_df.set_index("Group")[["Required", "Sent", "Remaining"]],
            color=CHART_COLORS,
            stack=False,
        )

    section_title("Items")

    options = ["All groups"] + data["groups"]
    group = st.selectbox("Work group", options, key=f"g_{sheet}")

    items = data["items"]
    base_cols = [c for c in ["No", "ASSET CODE", "Item"] if c in items.columns]

    if group == "All groups":
        value_cols = [c for c in items.columns if " | " in c]
    else:
        value_cols = [c for c in items.columns if c.startswith(f"{group} | ")]
        area = data["areas"].get(group)
        if area not in (None, 0):
            st.caption(f"Area = {fmt(area)}")

    df = items[base_cols + value_cols]

    only_active = st.checkbox(
        "Only items with quantities", value=True, key=f"ta_{sheet}"
    )
    if only_active and value_cols:
        df = df[(df[value_cols].abs().sum(axis=1) != 0)]

    df = text_filter(df, ["Item", "ASSET CODE"], f"s_{sheet}")

    st.caption(f"{len(df):,} items")
    show_table(df)

    if not footer.empty:
        with st.expander("Footer totals (as in the Excel sheet)"):
            show_table(footer, height=250)


def show_generic(data, sheet):
    st.caption(
        "Unknown sheet structure – showing the raw content. "
        "A dedicated parser can be added for this layout."
    )

    df = data["table"]

    if df.empty:
        st.info("This sheet does not contain any data.")
        return

    show_table(df)


def open_sheet(name):
    """Callback used by the 'Open sheet' buttons on the overview cards."""
    st.session_state.selected_sheet = name


def show_overview(project_data, project_name):
    rows = []

    for sheet, data in project_data.items():
        if data["type"] == "pr":
            rows.append({"Sheet": sheet, **data["summary"]})

    if not rows:
        st.info("No PR / PO sheet with quantities was found in this file.")
        return

    df = pd.DataFrame(rows)

    total_req = df["Required"].sum()
    total_send = df["Sent"].sum()
    ratio = (total_send / total_req) if total_req else None

    kpi_grid(
        [
            kpi("PR / PO with data", len(rows)),
            kpi("Total required", num(total_req)),
            kpi("Total sent", num(total_send), tone="good"),
            kpi("Total remaining", num(df["Remaining"].sum()), tone="warn"),
            kpi("Sent / Required", pct_text(ratio), bar=ratio),
        ]
    )

    section_title("PR / PO sheets")

    per_row = 3

    for start in range(0, len(rows), per_row):
        cols = st.columns(per_row)

        for col, row in zip(cols, rows[start:start + per_row]):
            with col:
                with st.container(key=f"pocard_{row['Sheet']}"):
                    po = row.get("PO No")
                    title = f"PO {po}" if po else row["Sheet"]

                    st.markdown(
                        '<div class="po-head">'
                        f'<div class="po-title">{esc(title)}</div>'
                        + chips_html([row["Category"]], accent_first=True)
                        + "</div>"
                        f'<div class="po-sub">{esc(row["Supplier"] or "No supplier")}</div>'
                        f'<div class="po-sub">PR {esc(row["PR No"]) if row["PR No"] else "—"}'
                        f' &nbsp;|&nbsp; {esc(row["Date"]) if row["Date"] else "no date"}</div>'
                        + bar_html(row["Sent %"])
                        + f'<div class="po-pct">{pct_text(row["Sent %"])} sent'
                        f' &nbsp;|&nbsp; {row["Items"]} items</div>'
                        '<div class="po-nums">'
                        f'<div><span>Required</span><b>{num(row["Required"])}</b></div>'
                        f'<div><span>Sent</span><b>{num(row["Sent"])}</b></div>'
                        f'<div><span>Remaining</span><b>{num(row["Remaining"])}</b></div>'
                        "</div>",
                        unsafe_allow_html=True,
                    )

                    st.button(
                        "Open sheet",
                        key=f"open_{row['Sheet']}",
                        on_click=open_sheet,
                        args=(row["Sheet"],),
                        width="stretch",
                    )

    with st.expander("Compare all PR / PO sheets"):
        display = df.copy()
        display["Sent %"] = display["Sent %"].map(
            lambda v: f"{v:.0%}" if pd.notna(v) else ""
        )
        show_table(display, height=250)

        st.bar_chart(
            df.set_index("Sheet")[["Required", "Sent", "Remaining"]],
            color=CHART_COLORS,
            stack=False,
        )


def show_sheet(sheet, data):
    kind = data["type"]

    if kind == "unit_prices":
        show_unit_prices(data, sheet)
    elif kind == "pr":
        show_pr(data, sheet)
    elif kind == "pr_empty":
        st.markdown(
            '<div class="empty-state"><b>Empty PR template</b><br>'
            "No quantities have been entered in this sheet yet. "
            "It will show up here with its KPIs and tables as soon as "
            "quantities are added in Excel.</div>",
            unsafe_allow_html=True,
        )
    elif kind == "total":
        show_total(data, sheet)
    elif kind == "error":
        st.error(f"Could not read this sheet: {data['message']}")
    else:
        show_generic(data, sheet)


# =========================================================
# CHECK DATA FOLDER / FILES
# =========================================================

excel_files = get_excel_files()

if not DATA_FOLDER.exists():
    st.error("❌ The data folder does not exist.")
    st.write("Expected location:")
    st.code(str(DATA_FOLDER))
    st.info(
        """
        GitHub structure:

        test_streamlit/
        ├── app.py
        ├── requirements.txt
        └── data/
            └── Project.xlsx
        """
    )
    st.stop()

if not excel_files:
    st.error("❌ No Excel files found inside the data folder.")
    st.code(str(DATA_FOLDER))

    items = list(DATA_FOLDER.iterdir())

    if items:
        st.write("Files/folders currently found:")
        for item in items:
            st.write(f"- {item.name}")
    else:
        st.warning("The data folder is completely empty.")

    st.stop()


# =========================================================
# SIDEBAR
# =========================================================

project_files = {f.stem: f for f in excel_files}
project_names = list(project_files)

# A shared link like  ?project=<file name>&sheet=<sheet name>
# opens directly on that file / sheet.
query_project = st.query_params.get("project")
start_index = (
    project_names.index(query_project) if query_project in project_names else 0
)

with st.sidebar:
    st.markdown(
        '<div class="side-brand">🏗️ Hassan Allam Construction'
        "<span>Projects dashboard</span></div>",
        unsafe_allow_html=True,
    )
    st.divider()

    # ---------------- files ----------------
    st.markdown(
        f'<div class="side-label">Files ({len(project_names)})</div>',
        unsafe_allow_html=True,
    )

    file_query = st.text_input(
        "🔎 Search files",
        key="file_search",
        placeholder="Type part of a file name...",
    )

    matching = [
        n for n in project_names if file_query.strip().lower() in n.lower()
    ]

    if not matching:
        st.warning("No file matches your search.")
        matching = project_names

    if file_query.strip():
        st.caption(f"Showing {len(matching)} of {len(project_names)} files")

    selected_project = st.selectbox(
        "Select file",
        matching,
        index=matching.index(project_names[start_index])
        if project_names[start_index] in matching
        else 0,
        key="project_select",
    )

    with st.expander(f"📄 All files ({len(project_names)})"):
        for f in excel_files:
            st.caption(f.name)

    st.divider()

    # ---------------- options ----------------
    st.markdown('<div class="side-label">Options</div>', unsafe_allow_html=True)

    hide_empty = st.toggle(
        "Hide empty PR templates", value=True, key="hide_empty"
    )

    st.toggle("🌙 Dark mode", key="dark_mode")

    st.divider()
    st.caption(f"App version: {APP_VERSION}")


if st.session_state.get("dark_mode"):
    st.markdown(DARK_CSS, unsafe_allow_html=True)


# =========================================================
# LOAD SELECTED PROJECT
# =========================================================

selected_file = project_files[selected_project]

project_data = load_excel_file(
    str(selected_file), selected_file.stat().st_mtime
)

if not project_data:
    st.error("❌ No sheets were found in this Excel file.")
    st.stop()

visible_sheets = [
    name
    for name, data in project_data.items()
    if not (hide_empty and data["type"] == "pr_empty")
]

if hide_empty and len(visible_sheets) < len(project_data):
    hidden_count = len(project_data) - len(visible_sheets)
else:
    hidden_count = 0

pr_with_data = sum(1 for d in project_data.values() if d["type"] == "pr")
updated = datetime.fromtimestamp(selected_file.stat().st_mtime).strftime("%Y-%m-%d %H:%M")

header_chips = [
    f"{len(project_data)} sheets",
    f"{pr_with_data} PR / PO with data",
    f"{hidden_count} empty templates hidden" if hidden_count else None,
    f"Updated {updated}",
]

st.markdown(
    '<div class="app-header"><div>'
    '<div class="brand">Hassan Allam Construction &nbsp;|&nbsp; Projects dashboard</div>'
    f'<div class="file-title">{esc(selected_project)}</div>'
    "</div>" + chips_html(header_chips) + "</div>",
    unsafe_allow_html=True,
)


# =========================================================
# SELECTED SHEET STATE
# =========================================================

nav_items = [OVERVIEW] + visible_sheets

first_load = "project" not in st.session_state

if st.session_state.get("project") != selected_project:
    st.session_state.project = selected_project

    query_sheet = st.query_params.get("sheet")

    st.session_state.selected_sheet = (
        query_sheet if (first_load and query_sheet in nav_items) else OVERVIEW
    )

if st.session_state.get("selected_sheet") not in nav_items:
    st.session_state.selected_sheet = OVERVIEW


# =========================================================
# SHEET NAVIGATION
# =========================================================

def nav_label(name):
    if name == OVERVIEW:
        return "🏠 Overview"
    return f"{SHEET_ICONS.get(project_data[name]['type'], '📄')} {name}"


section_title("📑 PR & PO Sheets")

choice = st.pills(
    "Sheets",
    nav_items,
    selection_mode="single",
    default=st.session_state.selected_sheet,
    format_func=nav_label,
    label_visibility="collapsed",
)

if choice is not None:
    st.session_state.selected_sheet = choice


# =========================================================
# SELECTED SHEET
# =========================================================

selected_sheet = st.session_state.selected_sheet

# keep the page address in sync so the link can be shared as-is
st.query_params["project"] = selected_project

if selected_sheet == OVERVIEW:
    if "sheet" in st.query_params:
        del st.query_params["sheet"]
else:
    st.query_params["sheet"] = selected_sheet

st.write("")

if selected_sheet == OVERVIEW:
    show_overview(project_data, selected_project)
else:
    data = project_data[selected_sheet]

    st.caption(
        f"{SHEET_ICONS.get(data['type'], '📄')} {selected_sheet}  –  "
        f"{SHEET_TYPE_LABELS.get(data['type'], data['type'])}"
    )

    show_sheet(selected_sheet, data)
