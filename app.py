from pathlib import Path

code = r'''"""
Hassan Allam Construction
Construction Projects & Pipeline Dashboard
Streamlit + Google Sheets

Architecture:
Google Sheets
    ↓
gspread
    ↓
Pandas
    ↓
Dynamic Dashboard

Admin URL:
https://YOUR-APP.streamlit.app/

Viewer URL:
https://YOUR-APP.streamlit.app/?view=dashboard

Viewer mode:
- No Excel upload
- No sidebar
- No app navigation
- Reads Google Sheets directly
"""

import os
import pandas as pd
import plotly.express as px
import streamlit as st

import gspread
from google.oauth2.service_account import Credentials


# ============================================================================
# 1) PAGE CONFIG
# ============================================================================

st.set_page_config(
    page_title="Hassan Allam Construction | Projects Dashboard",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="collapsed",
    menu_items={}
)

VIEWER_MODE = st.query_params.get("view") == "dashboard"


# ============================================================================
# 2) CONFIGURATION
# ============================================================================

COMPANY = "Hassan Allam Construction"

PRIMARY = "#C8102E"
DARK = "#1F2933"
LIGHT = "#F4F5F7"
GRAY = "#6B7280"
GREEN = "#2E9E5B"
ORANGE = "#D97706"
BORDER = "#E1E4E8"

LOGO_PATH = "assets/logo.png"


STATUS_COLORS = {
    "Delivered": GREEN,
    "Pending / On Track": DARK,
    "Delayed": PRIMARY,
    "Unknown": "#9AA5B1",
    "Completed": GREEN,
    "Active": DARK,
    "Not Started": "#9AA5B1",
}


# ============================================================================
# 3) CSS
# ============================================================================

st.markdown(
    f"""
<style>

.stApp {{
    background-color: #E7E1D7;
    background-image:
        linear-gradient(rgba(92,78,62,.045) 1px, transparent 1px),
        linear-gradient(90deg, rgba(92,78,62,.045) 1px, transparent 1px),
        radial-gradient(circle at 15% 10%, rgba(200,16,46,.055), transparent 28%);
    background-size: 32px 32px, 32px 32px, 100% 100%;
}}

[data-testid="stSidebar"] {{
    background: {LIGHT};
    border-top: 5px solid {PRIMARY};
}}

h1, h2, h3 {{
    color: {DARK};
}}

h2, h3 {{
    border-left: 5px solid {PRIMARY};
    padding-left: 10px;
}}

.kpi {{
    background: #ffffff;
    border: 1px solid {BORDER};
    border-top: 4px solid {PRIMARY};
    border-radius: 10px;
    padding: 14px 16px;
    margin-bottom: 10px;
    box-shadow: 0 2px 6px rgba(31,41,51,.07);
}}

.kpi .l {{
    font-size: 13px;
    color: {GRAY};
    text-transform: uppercase;
    letter-spacing: .4px;
}}

.kpi .v {{
    font-size: 27px;
    font-weight: 700;
    color: {DARK};
}}

div.stButton > button {{
    border-radius: 8px;
    border: 1.5px solid {DARK};
    color: {DARK};
    background: #fff;
    font-weight: 600;
    width: 100%;
    padding: .6rem 0;
}}

div.stButton > button:hover {{
    background: {PRIMARY};
    color: #fff;
    border-color: {PRIMARY};
}}

.hero {{
    background:
        repeating-linear-gradient(
            135deg,
            rgba(255,255,255,.04) 0 14px,
            transparent 14px 28px
        ),
        linear-gradient(120deg,{DARK},#3a4552);
    color:#fff;
    padding:38px 34px;
    border-radius:6px;
    border-bottom:6px solid {PRIMARY};
    margin-bottom:22px;
}}

.hero .co {{
    color:{PRIMARY};
    font-weight:700;
    letter-spacing:2px;
    font-size:14px;
    text-transform:uppercase;
    background:#fff;
    display:inline-block;
    padding:2px 10px;
    border-radius:3px;
}}

.hero h1 {{
    color:#fff;
    margin:12px 0 0;
}}

.hero p {{
    margin:6px 0 0;
    opacity:.9;
    font-size:18px;
}}

.topbar {{
    border-bottom:3px solid {PRIMARY};
    padding:4px 0 6px;
    margin-bottom:10px;
    color:{DARK};
    font-weight:700;
    letter-spacing:1px;
    font-size:13px;
    text-transform:uppercase;
}}

.badge {{
    display:inline-block;
    padding:5px 12px;
    border-radius:20px;
    font-weight:600;
    font-size:12px;
}}

/* Hide Streamlit chrome in viewer mode */
{'''
header[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDecoration"],
#MainMenu,
.stAppDeployButton,
footer,
[data-testid="stStatusWidget"] {{
    display:none !important;
    visibility:hidden !important;
}}
''' if VIEWER_MODE else ""}

</style>
""",
    unsafe_allow_html=True,
)


# ============================================================================
# 4) SESSION STATE
# ============================================================================

ss = st.session_state

DEFAULT_STATE = {
    "page": "home",
    "sheet": None,
    "project": None,
    "pipeline": None,
}

for key, value in DEFAULT_STATE.items():
    ss.setdefault(key, value)


# ============================================================================
# 5) GOOGLE SHEETS DATA LAYER
# ============================================================================

@st.cache_resource
def get_google_client():
    """
    Uses a Google Service Account stored in Streamlit Secrets.

    Required secrets:

    [google_service_account]
    type = "service_account"
    project_id = "..."
    private_key_id = "..."
    private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
    client_email = "..."
    client_id = "..."
    auth_uri = "https://accounts.google.com/o/oauth2/auth"
    token_uri = "https://oauth2.googleapis.com/token"
    auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
    client_x509_cert_url = "..."

    GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/..."
    """

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]

    credentials = Credentials.from_service_account_info(
        st.secrets["google_service_account"],
        scopes=scopes,
    )

    return gspread.authorize(credentials)


@st.cache_resource
def get_google_workbook():
    client = get_google_client()

    sheet_url = st.secrets.get("GOOGLE_SHEET_URL")

    if not sheet_url:
        raise RuntimeError(
            "GOOGLE_SHEET_URL is missing from Streamlit Secrets."
        )

    return client.open_by_url(sheet_url)


@st.cache_data(ttl=60, show_spinner=False)
def get_sheet_names():
    workbook = get_google_workbook()

    return [
        worksheet.title
        for worksheet in workbook.worksheets()
    ]


@st.cache_data(ttl=60, show_spinner=False)
def load_google_sheet(sheet_name):
    workbook = get_google_workbook()
    worksheet = workbook.worksheet(sheet_name)

    records = worksheet.get_all_records()

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame(records)

    df = (
        df
        .dropna(how="all")
        .dropna(axis=1, how="all")
        .copy()
    )

    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    return df


# ============================================================================
# 6) COLUMN MAPS
# ============================================================================

COLUMN_MAP = {
    "project": ["Project", "Project Name", "ProjectName"],
    "item": ["Item", "Material", "Description"],
    "location": ["Location", "Site", "Zone"],
    "required_qty": ["Required Qty", "Required Quantity", "Quantity", "Qty"],
    "delivered_qty": ["Delivered Qty", "Delivered Quantity"],
    "remaining_qty": ["Remaining Qty", "Remaining Quantity"],
    "required_date": ["Required Date"],
    "expected_date": ["Expected Date"],
    "actual_date": ["Actual Date", "Delivery Date"],
    "status": ["Status"],
}

PROJECT_COLS = {
    "id": ["Project ID", "ProjectID", "Project Code", "Project No", "Project Number", "Code"],
    "name": ["Project Name", "Project", "Name"],
    "client": ["Client", "Client Name", "Owner"],
    "manager": ["Project Manager", "Manager", "Project Manager Name"],
    "location": ["Location", "Site", "Project Location"],
    "start": ["Start Date", "Project Start Date", "Start"],
    "end": ["End Date", "Project End Date", "Finish Date", "Finish"],
    "status": ["Status", "Project Status"],
    "budget": ["Total Budget", "Budget", "Project Budget"],
    "cost": ["Total Cost", "Actual Cost", "Actual", "Project Cost"],
    "progress": ["Progress %", "Progress", "% Complete", "Completion %"],
    "description": ["Description", "Project Description"],
}

PIPELINE_COLS = {
    "proj_id": PROJECT_COLS["id"],
    "proj_name": ["Project Name", "Project"],
    "id": ["Pipeline ID", "PipelineID", "Pipeline Code", "Pipeline No", "Pipeline Number"],
    "name": ["Pipeline Name", "Pipeline", "Name"],
    "desc": ["Description", "Pipeline Description"],
    "budget": ["Budget", "Total Budget", "Pipeline Budget"],
    "planned": ["Planned Cost", "Planned", "Planned Amount"],
    "actual": ["Actual Cost", "Actual", "Actual Amount", "Spent"],
    "remaining": ["Remaining Cost", "Remaining", "Remaining Budget"],
    "progress": ["Progress %", "Progress", "% Complete", "Completion %"],
    "status": ["Status", "Pipeline Status"],
    "start": ["Start Date", "Pipeline Start Date", "Start"],
    "end": ["End Date", "Finish Date", "Pipeline End Date", "End"],
    "category": ["Category", "Type", "Pipeline Category"],
}

DETAIL_COLS = {
    "proj_id": PROJECT_COLS["id"],
    "proj_name": ["Project Name", "Project"],
    "id": PIPELINE_COLS["id"],
    "name": ["Pipeline Name", "Pipeline"],
    "planned": ["Planned Cost", "Planned", "Planned Amount"],
    "actual": ["Actual Cost", "Actual", "Actual Amount", "Amount", "Spent"],
    "category": ["Category", "Cost Type", "Type"],
    "item": ["Item", "Description", "Material"],
    "date": ["Date", "Cost Date", "Transaction Date"],
}


# ============================================================================
# 7) GENERAL HELPERS
# ============================================================================

def normalize_text(value):
    if pd.isna(value):
        return ""

    text = str(value).strip()

    if text.endswith(".0"):
        text = text[:-2]

    return text


def normalize_key(value):
    return normalize_text(value).lower()


def pick(df, aliases):
    if df is None or df.empty:
        return None

    lower = {
        str(c).strip().lower(): c
        for c in df.columns
    }

    for alias in aliases:
        key = str(alias).strip().lower()

        if key in lower:
            return lower[key]

    return None


def cols_of(df, spec):
    return {
        key: pick(df, aliases)
        for key, aliases in spec.items()
    }


def find_col(df, key):
    return pick(df, COLUMN_MAP.get(key, []))


def to_num(series):
    if series is None:
        return pd.Series(dtype=float)

    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True),
        errors="coerce",
    )


def fmt(value):
    if pd.isna(value):
        return "—"

    try:
        value = float(value)

        if value.is_integer():
            return f"{value:,.0f}"

        return f"{value:,.2f}"

    except Exception:
        return str(value)


def to_pct(series):
    s = to_num(series)

    if s.empty or s.dropna().empty:
        return s

    if s.dropna().max() <= 1:
        return s * 100

    return s


def safe_sum(series):
    if series is None:
        return float("nan")

    return series.sum(min_count=1)


# ============================================================================
# 8) KPI COMPONENTS
# ============================================================================

def kpi(col, label, value):
    col.markdown(
        f"""
        <div class="kpi">
            <div class="l">{label}</div>
            <div class="v">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def kpi_row(items):
    for i in range(0, len(items), 3):
        columns = st.columns(3)

        for column, item in zip(columns, items[i:i + 3]):
            label, value = item
            kpi(column, label, value)


# ============================================================================
# 9) NAVIGATION
# ============================================================================

def go(page):
    ss.page = page


def open_project(project_key):
    ss.project = project_key
    ss.pipeline = None
    ss.page = "project"


def open_pipeline(pipeline_key):
    ss.pipeline = pipeline_key
    ss.page = "pipeline"


def open_sheet(sheet_name):
    ss.sheet = sheet_name
    ss.page = "sheet"


def grid(items, callback, prefix, per_row=3):
    if not items:
        return

    for i in range(0, len(items), per_row):
        columns = st.columns(per_row)
        batch = items[i:i + per_row]

        for j, (label, arg) in enumerate(batch):
            columns[j].button(
                label,
                key=f"{prefix}_{i+j}",
                on_click=callback,
                args=(arg,),
                width="stretch",
            )


# ============================================================================
# 10) DELIVERY DASHBOARD
# ============================================================================

def compute_delivery_status(df, columns):
    def status(row):
        actual_col = columns.get("actual_date")
        expected_col = columns.get("expected_date")
        required_col = columns.get("required_date")

        if actual_col and pd.notna(row.get(actual_col)):
            return "Delivered"

        expected = row.get(expected_col) if expected_col else pd.NaT
        required = row.get(required_col) if required_col else pd.NaT

        if pd.notna(expected) and pd.notna(required) and expected > required:
            return "Delayed"

        if pd.notna(expected):
            return "Pending / On Track"

        return "Unknown"

    return df.apply(status, axis=1)


def show_delivery_dashboard(df, sheet_name):
    columns = {
        key: find_col(df, key)
        for key in COLUMN_MAP
    }

    df = df.copy()

    for key in ("required_date", "expected_date", "actual_date"):
        column = columns[key]

        if column:
            df[column] = pd.to_datetime(df[column], errors="coerce")

    for key in ("required_qty", "delivered_qty", "remaining_qty"):
        column = columns[key]

        if column:
            df[column] = to_num(df[column])

    if columns["actual_date"] or columns["expected_date"]:
        df["Computed Status"] = compute_delivery_status(df, columns)
        status_column = "Computed Status"
    elif columns["status"]:
        status_column = columns["status"]
    else:
        status_column = None

    if columns["required_qty"] and columns["delivered_qty"]:
        df["Remaining Qty (calc)"] = (
            df[columns["required_qty"]]
            - df[columns["delivered_qty"]].fillna(0)
        ).clip(lower=0)

        remaining_column = "Remaining Qty (calc)"
    else:
        remaining_column = columns["remaining_qty"]

    st.subheader("🔎 Filters")

    filter_columns = st.columns(5)

    filter_items = [
        ("Project", columns["project"]),
        ("Status", status_column),
        ("Item", columns["item"]),
        ("Location", columns["location"]),
    ]

    for column_ui, (label, column) in zip(filter_columns, filter_items):
        if column:
            options = sorted(
                df[column].dropna().astype(str).unique().tolist()
            )

            selected = column_ui.multiselect(
                label,
                options,
                key=f"delivery_{sheet_name}_{label}",
            )

            if selected:
                df = df[df[column].astype(str).isin(selected)]

    if (
        columns["expected_date"]
        and df[columns["expected_date"]].notna().any()
    ):
        min_date = df[columns["expected_date"]].min().date()
        max_date = df[columns["expected_date"]].max().date()

        if min_date < max_date:
            date_range = filter_columns[4].date_input(
                "Expected Date",
                (min_date, max_date),
                key=f"delivery_date_{sheet_name}",
            )

            if isinstance(date_range, tuple) and len(date_range) == 2:
                mask = (
                    df[columns["expected_date"]].dt.date.between(
                        date_range[0], date_range[1]
                    )
                    | df[columns["expected_date"]].isna()
                )
                df = df[mask]

    today = pd.Timestamp.today().normalize()
    week_start = today - pd.Timedelta(days=today.weekday())

    items = [("Total Deliveries", len(df))]

    if status_column:
        values = df[status_column].astype(str).str.lower()

        items.extend(
            [
                ("Delivered", int(values.str.contains("deliver").sum())),
                ("Pending", int(values.str.contains("pending|track").sum())),
                ("Delayed", int(values.str.contains("delay").sum())),
            ]
        )

    if remaining_column:
        items.append(("Remaining Quantity", fmt(df[remaining_column].sum())))

    if columns["expected_date"]:
        expected = df[columns["expected_date"]]

        items.extend(
            [
                (
                    "Expected This Week",
                    int(
                        expected.between(
                            week_start,
                            week_start + pd.Timedelta(days=6),
                        ).sum()
                    ),
                ),
                (
                    "Expected This Month",
                    int(
                        (
                            (expected.dt.month == today.month)
                            & (expected.dt.year == today.year)
                        ).sum()
                    ),
                ),
            ]
        )

    kpi_row(items)

    st.subheader("📋 Delivery Table")

    show_columns = [
        columns["project"],
        columns["item"],
        columns["location"],
        columns["required_qty"],
        columns["delivered_qty"],
        remaining_column,
        columns["required_date"],
        columns["expected_date"],
        columns["actual_date"],
        status_column,
    ]

    show_columns = [c for c in show_columns if c]

    st.dataframe(
        df[show_columns] if show_columns else df,
        width="stretch",
        hide_index=True,
    )

    st.subheader("📊 Charts")

    c1, c2 = st.columns(2)

    if status_column:
        count = df[status_column].value_counts().reset_index()
        count.columns = ["Status", "Count"]

        c1.plotly_chart(
            px.pie(
                count,
                names="Status",
                values="Count",
                hole=.45,
                title="Deliveries by Status",
                color="Status",
                color_discrete_map=STATUS_COLORS,
            ),
            width="stretch",
        )

    if (
        columns["expected_date"]
        and df[columns["expected_date"]].notna().any()
    ):
        by_date = (
            df.groupby(df[columns["expected_date"]].dt.date)
            .size()
            .reset_index(name="Deliveries")
        )

        c2.plotly_chart(
            px.bar(
                by_date,
                x=columns["expected_date"],
                y="Deliveries",
                title="Expected Deliveries by Date",
                color_discrete_sequence=[PRIMARY],
            ),
            width="stretch",
        )

    c3, c4 = st.columns(2)

    if columns["delivered_qty"] and remaining_column:
        total = pd.DataFrame(
            {
                "Type": ["Delivered", "Remaining"],
                "Qty": [
                    df[columns["delivered_qty"]].sum(),
                    df[remaining_column].sum(),
                ],
            }
        )

        c3.plotly_chart(
            px.bar(
                total,
                x="Type",
                y="Qty",
                color="Type",
                title="Delivered vs Remaining Quantity",
            ),
            width="stretch",
        )

    if columns["project"]:
        by_project = (
            df.groupby(columns["project"])
            .size()
            .reset_index(name="Deliveries")
        )

        c4.plotly_chart(
            px.bar(
                by_project,
                x=columns["project"],
                y="Deliveries",
                title="Deliveries by Project",
                color_discrete_sequence=[PRIMARY],
            ),
            width="stretch",
        )


# ============================================================================
# 11) GENERIC DASHBOARD
# ============================================================================

def show_generic_dashboard(df, sheet_name):
    numeric_columns = [
        c for c in df.columns
        if pd.api.types.is_numeric_dtype(df[c])
    ]

    categorical_columns = [
        c for c in df.columns
        if df[c].dtype == object
        and 2 <= df[c].nunique() <= 20
    ]

    st.subheader("🔎 Filters")

    search = st.text_input(
        "Search",
        key=f"generic_search_{sheet_name}",
    )

    if search:
        mask = (
            df.astype(str)
            .apply(
                lambda column: column.str.contains(
                    search,
                    case=False,
                    na=False,
                )
            )
            .any(axis=1)
        )
        df = df[mask]

    filter_columns = st.columns(3)

    for column_ui, column in zip(
        filter_columns,
        categorical_columns[:3],
    ):
        selected = column_ui.multiselect(
            column,
            sorted(df[column].dropna().astype(str).unique()),
            key=f"generic_{sheet_name}_{column}",
        )

        if selected:
            df = df[df[column].astype(str).isin(selected)]

    items = [
        ("Total Records", len(df)),
        ("Columns", len(df.columns)),
    ]

    for column in numeric_columns[:3]:
        items.append((f"Total {column}", fmt(df[column].sum())))

    status_column = find_col(df, "status")

    if status_column:
        values = df[status_column].astype(str).str.lower()

        items.extend(
            [
                (
                    "Completed",
                    int(values.str.contains("complet|deliver|done").sum()),
                ),
                ("Pending", int(values.str.contains("pend").sum())),
                ("Delayed", int(values.str.contains("delay").sum())),
            ]
        )

    kpi_row(items[:6])

    st.subheader("📋 Data")

    st.dataframe(
        df,
        width="stretch",
        hide_index=True,
    )

    if categorical_columns or numeric_columns:
        st.subheader("📊 Charts")

        charts = st.columns(2)

        if categorical_columns:
            column = categorical_columns[0]

            count = df[column].value_counts().reset_index()
            count.columns = [column, "Count"]

            charts[0].plotly_chart(
                px.bar(
                    count,
                    x=column,
                    y="Count",
                    title=f"Records by {column}",
                    color_discrete_sequence=[PRIMARY],
                ),
                width="stretch",
            )

        if categorical_columns and numeric_columns:
            category = categorical_columns[0]
            number = numeric_columns[0]

            aggregation = (
                df.groupby(category)[number]
                .sum()
                .reset_index()
            )

            charts[1].plotly_chart(
                px.pie(
                    aggregation,
                    names=category,
                    values=number,
                    hole=.45,
                    title=f"{number} by {category}",
                ),
                width="stretch",
            )


# ============================================================================
# 12) FORMWORK
# ============================================================================

def is_formwork_sheet(sheet_name, df=None):
    name = str(sheet_name).lower().strip()

    if name == "workbook index":
        return False

    keywords = [
        "formwork", "villa", "residential building",
        "commercial building", "warehouse", "office building",
        "school", "hospital", "mosque", "compound",
        "industrial building",
    ]

    if any(k in name for k in keywords):
        return True

    if df is not None and not df.empty:
        cols = {str(c).strip().lower() for c in df.columns}

        required = {
            "element / activity",
            "unit",
            "quantity",
            "unit rate (egp)",
            "budget cost (egp)",
        }

        return len(required.intersection(cols)) >= 4

    return False


def show_formwork_dashboard(df, sheet_name):
    st.title(f"🏗️ Formwork Cost — {sheet_name}")

    if df is None or df.empty:
        st.warning("This construction sheet is empty.")
        return

    def col(*names):
        return pick(df, list(names))

    item_col = col("Element / Activity", "Element", "Activity")
    unit_col = col("Unit")
    qty_col = col("Quantity", "Qty")
    rate_col = col("Unit Rate (EGP)", "Unit Rate", "Rate")
    budget_col = col("Budget Cost (EGP)", "Budget Cost", "Budget")
    actual_col = col("Actual Cost (EGP)", "Actual Cost", "Actual")
    variance_col = col("Variance (EGP)", "Variance")

    work = df.copy()

    if item_col:
        work = work[work[item_col].notna()].copy()
        work = work[
            ~work[item_col].astype(str).str.upper().isin(
                ["TOTAL", "SUMMARY BY CATEGORY"]
            )
        ].copy()
        work = work[
            ~work[item_col].astype(str).str.contains(
                "^Category$",
                case=False,
                regex=True,
            )
        ].copy()

    for c in [qty_col, rate_col, budget_col, actual_col, variance_col]:
        if c:
            work[c] = to_num(work[c])

    if budget_col is None and qty_col and rate_col:
        work["Budget Cost (Calculated)"] = work[qty_col] * work[rate_col]
        budget_series = work["Budget Cost (Calculated)"]
    elif budget_col:
        budget_series = work[budget_col]
    else:
        budget_series = pd.Series(dtype=float)

    if actual_col and work[actual_col].notna().any():
        actual_series = work[actual_col]
    else:
        actual_series = (
            budget_series * 0.90
            if not budget_series.empty
            else pd.Series(dtype=float)
        )

    total_budget = safe_sum(budget_series)
    total_actual = safe_sum(actual_series)

    variance = (
        total_budget - total_actual
        if pd.notna(total_budget) and pd.notna(total_actual)
        else float("nan")
    )

    used = (
        total_actual / total_budget * 100
        if pd.notna(total_budget) and total_budget
        else float("nan")
    )

    kpi_row(
        [
            ("Construction Type", sheet_name),
            ("Total Budget", fmt(total_budget)),
            ("Actual Cost", fmt(total_actual)),
            ("Remaining / Variance", fmt(variance)),
            ("Cost Used", f"{used:,.1f}%" if pd.notna(used) else "—"),
        ]
    )

    st.subheader("📋 Formwork Cost Breakdown")

    display_cols = [
        c for c in [
            item_col,
            unit_col,
            qty_col,
            rate_col,
            budget_col,
            actual_col,
            variance_col,
        ]
        if c
    ]

    if display_cols:
        shown = work[display_cols].copy()

        rename = {
            item_col: "Element / Activity" if item_col else "",
            unit_col: "Unit" if unit_col else "",
            qty_col: "Quantity" if qty_col else "",
            rate_col: "Unit Rate (EGP)" if rate_col else "",
            budget_col: "Budget Cost (EGP)" if budget_col else "",
            actual_col: "Actual Cost (EGP)" if actual_col else "",
            variance_col: "Variance (EGP)" if variance_col else "",
        }

        shown = shown.rename(
            columns={k: v for k, v in rename.items() if k}
        )

        if "Budget Cost (EGP)" not in shown.columns and not budget_series.empty:
            shown["Budget Cost (EGP)"] = budget_series.values

        if "Actual Cost (EGP)" not in shown.columns and not actual_series.empty:
            shown["Actual Cost (EGP)"] = actual_series.values

        if "Variance (EGP)" not in shown.columns and not budget_series.empty:
            shown["Variance (EGP)"] = (
                budget_series.values - actual_series.values
            )

        st.dataframe(
            shown,
            use_container_width=True,
            hide_index=True,
        )

    st.subheader("📊 Summary by Category")

    if item_col and not work.empty:
        categories = []

        for value in work[item_col].astype(str):
            if "erection" in value.lower() or "stripping" in value.lower():
                category = "Labor / Installation"
            elif "consumable" in value.lower():
                category = "Consumables"
            else:
                category = "Formwork Activities"

            if category not in categories:
                categories.append(category)

        rows = []

        for category in categories:
            mask = work[item_col].astype(str).map(
                lambda x: (
                    (
                        "erection" in x.lower()
                        or "stripping" in x.lower()
                    )
                    if category == "Labor / Installation"
                    else (
                        "consumable" in x.lower()
                        if category == "Consumables"
                        else (
                            "erection" not in x.lower()
                            and "stripping" not in x.lower()
                            and "consumable" not in x.lower()
                        )
                    )
                )
            )

            rows.append(
                {
                    "Category": category,
                    "Budget": safe_sum(budget_series[mask]),
                    "Actual": safe_sum(actual_series[mask]),
                    "Variance": safe_sum(
                        budget_series[mask] - actual_series[mask]
                    ),
                }
            )

        st.dataframe(
            pd.DataFrame(rows),
            use_container_width=True,
            hide_index=True,
        )


# ============================================================================
# 13) SHEET CLASSIFICATION
# ============================================================================

def classify_sheet(sheet_name, df=None):
    name = sheet_name.lower().strip()

    if (
        "pipeline" in name
        and any(
            word in name
            for word in [
                "detail", "details", "cost",
                "costs", "line", "transaction",
            ]
        )
    ):
        return "details"

    if "pipeline" in name:
        return "pipelines"

    if "project" in name:
        return "projects"

    if any(
        word in name
        for word in ["cost", "transaction", "expense"]
    ):
        if df is not None:
            pipeline_id = pick(df, PIPELINE_COLS["id"])
            pipeline_name = pick(df, PIPELINE_COLS["name"])
            actual = pick(df, DETAIL_COLS["actual"])

            if (pipeline_id or pipeline_name) and actual:
                return "details"

    if df is not None and not df.empty:
        project_id = pick(df, PROJECT_COLS["id"])
        project_name = pick(df, PROJECT_COLS["name"])

        pipeline_id = pick(df, PIPELINE_COLS["id"])
        pipeline_name = pick(df, PIPELINE_COLS["name"])

        actual = pick(df, DETAIL_COLS["actual"])
        planned = pick(df, DETAIL_COLS["planned"])

        if (
            (pipeline_id or pipeline_name)
            and (
                actual
                or planned
                or pick(df, PIPELINE_COLS["progress"])
                or pick(df, PIPELINE_COLS["status"])
            )
        ):
            if (
                pick(df, DETAIL_COLS["item"])
                or pick(df, DETAIL_COLS["category"])
                or pick(df, DETAIL_COLS["date"])
            ):
                return "details"

            return "pipelines"

        if project_id or project_name:
            if not pipeline_id and not pipeline_name:
                return "projects"

    return "other"


# ============================================================================
# 14) PIPELINE STATUS
# ============================================================================

def pipeline_state(status, progress, end):
    status_text = str(status).lower()

    if pd.notna(progress) and float(progress) >= 100:
        return "Completed"

    if any(
        word in status_text
        for word in [
            "complete", "completed", "done",
            "finish", "finished", "closed",
        ]
    ):
        return "Completed"

    if "delay" in status_text or "late" in status_text:
        return "Delayed"

    if pd.notna(end) and end < pd.Timestamp.today().normalize():
        return "Delayed"

    if pd.notna(progress) and float(progress) > 0:
        return "Active"

    if any(
        word in status_text
        for word in [
            "active", "progress",
            "ongoing", "running",
        ]
    ):
        return "Active"

    return "Not Started"


# ============================================================================
# 15) BUILD MODEL FROM GOOGLE SHEETS
# ============================================================================

@st.cache_data(ttl=60, show_spinner=False)
def build_model():
    sheet_names = get_sheet_names()

    formwork_sheets = []

    for sheet_name in sheet_names:
        try:
            df = load_google_sheet(sheet_name)

            if is_formwork_sheet(sheet_name, df):
                formwork_sheets.append(sheet_name)

        except Exception:
            pass

    if formwork_sheets:
        return {
            "ok": False,
            "mode": "formwork",
            "role": {},
            "other": [
                x for x in sheet_names
                if x not in formwork_sheets
            ],
            "formwork_sheets": formwork_sheets,
            "warnings": [],
        }

    classified = {}
    other_sheets = []

    for sheet_name in sheet_names:
        try:
            df = load_google_sheet(sheet_name)
            role = classify_sheet(sheet_name, df)
        except Exception:
            role = "other"

        if role in ["projects", "pipelines", "details"]:
            classified.setdefault(role, []).append(sheet_name)
        else:
            other_sheets.append(sheet_name)

    warnings = []

    project_sheet = classified.get("projects", [None])[0]
    pipeline_sheet = classified.get("pipelines", [None])[0]
    detail_sheet = classified.get("details", [None])[0]

    extra_project_sheets = classified.get("projects", [])[1:]
    extra_pipeline_sheets = classified.get("pipelines", [])[1:]
    extra_detail_sheets = classified.get("details", [])[1:]

    if project_sheet is None and pipeline_sheet is None:
        return {
            "ok": False,
            "role": {},
            "other": other_sheets,
            "warnings": [
                "No Projects or Pipelines structure was detected."
            ],
        }

    if project_sheet:
        P = load_google_sheet(project_sheet).reset_index(drop=True)
    else:
        P = pd.DataFrame()

    pipeline_frames = []

    if pipeline_sheet:
        pipeline_frames.append(
            load_google_sheet(pipeline_sheet)
        )

    for sheet_name in extra_pipeline_sheets:
        pipeline_frames.append(
            load_google_sheet(sheet_name)
        )

    L = (
        pd.concat(
            pipeline_frames,
            ignore_index=True,
            sort=False,
        )
        if pipeline_frames
        else pd.DataFrame()
    )

    detail_frames = []

    if detail_sheet:
        detail_frames.append(
            load_google_sheet(detail_sheet)
        )

    for sheet_name in extra_detail_sheets:
        detail_frames.append(
            load_google_sheet(sheet_name)
        )

    D = (
        pd.concat(
            detail_frames,
            ignore_index=True,
            sort=False,
        )
        if detail_frames
        else None
    )

    # ------------------------------------------------------------------------
    # PROJECTS
    # ------------------------------------------------------------------------

    if not P.empty:
        project_columns = cols_of(P, PROJECT_COLS)

        project_key_column = (
            project_columns["id"]
            or project_columns["name"]
        )

        if not project_key_column:
            warnings.append(
                "Projects sheet does not contain Project ID or Project Name."
            )
            P = pd.DataFrame()
        else:
            P["_key"] = P[project_key_column].map(normalize_text)
            P = P[P["_key"] != ""].copy()

            if project_columns["name"]:
                P["_name"] = (
                    P[project_columns["name"]]
                    .map(normalize_text)
                )
            else:
                P["_name"] = P["_key"]

            if P["_key"].duplicated().any():
                warnings.append(
                    "Duplicate Project IDs were found. "
                    "Only the first record was kept."
                )
                P = P.drop_duplicates(
                    "_key",
                    keep="first",
                )

            P["_budget"] = (
                to_num(P[project_columns["budget"]])
                if project_columns["budget"]
                else float("nan")
            )

            P["_cost"] = (
                to_num(P[project_columns["cost"]])
                if project_columns["cost"]
                else float("nan")
            )

            P["_progress"] = (
                to_pct(P[project_columns["progress"]])
                if project_columns["progress"]
                else float("nan")
            )
    else:
        P = pd.DataFrame(
            columns=[
                "_key",
                "_name",
                "_budget",
                "_cost",
                "_progress",
            ]
        )

    name_to_key = {}

    if not P.empty:
        name_to_key = {
            normalize_key(name): key
            for name, key in zip(
                P["_name"],
                P["_key"],
            )
        }

    # ------------------------------------------------------------------------
    # PIPELINES
    # ------------------------------------------------------------------------

    pipeline_columns = {}

    if not L.empty:
        pipeline_columns = cols_of(L, PIPELINE_COLS)

        project_id_column = pipeline_columns["proj_id"]
        project_name_column = pipeline_columns["proj_name"]

        if project_id_column and not P.empty:
            L["_project_key"] = (
                L[project_id_column].map(normalize_text)
            )
        elif project_name_column:
            L["_project_key"] = (
                L[project_name_column].map(
                    lambda x: name_to_key.get(
                        normalize_key(x),
                        normalize_text(x),
                    )
                )
            )
        elif len(P) == 1:
            L["_project_key"] = P["_key"].iloc[0]
            warnings.append(
                "Pipelines had no Project ID/Project Name column. "
                "All pipelines were assigned to the only project."
            )
        else:
            L["_project_key"] = ""
            warnings.append(
                "Some pipelines could not be linked to a project."
            )

        pipeline_id_column = pipeline_columns["id"]
        pipeline_name_column = pipeline_columns["name"]

        if pipeline_id_column:
            L["_pid"] = L[pipeline_id_column].map(normalize_text)
        elif pipeline_name_column:
            L["_pid"] = L[pipeline_name_column].map(normalize_text)
        else:
            L["_pid"] = [
                f"PIPE-{i + 1}"
                for i in range(len(L))
            ]
            warnings.append(
                "Pipeline ID/Name was not found. "
                "Generated internal pipeline IDs."
            )

        if pipeline_name_column:
            L["_name"] = L[pipeline_name_column].map(normalize_text)
        else:
            L["_name"] = L["_pid"]

        L["_pk"] = (
            L["_project_key"].astype(str)
            + "||"
            + L["_pid"].astype(str)
        )

        if L["_pk"].duplicated(keep=False).any():
            L["_pk"] = (
                L["_pk"]
                + "#"
                + L.groupby("_pk").cumcount().astype(str)
            )
            warnings.append(
                "Duplicate Pipeline IDs were found inside the same project."
            )
    else:
        L = pd.DataFrame(
            columns=[
                "_project_key",
                "_pid",
                "_name",
                "_pk",
            ]
        )

    # ------------------------------------------------------------------------
    # AUTO CREATE PROJECTS FROM PIPELINES
    # ------------------------------------------------------------------------

    if not L.empty:
        pipeline_project_keys = set(
            L["_project_key"].dropna().astype(str).tolist()
        )

        known_project_keys = set(
            P["_key"].astype(str).tolist()
        )

        missing_projects = [
            key
            for key in pipeline_project_keys
            if key and key not in known_project_keys
        ]

        if missing_projects:
            warnings.append(
                f"{len(missing_projects)} project(s) were found in "
                "Pipelines but not in Projects. They were created automatically."
            )

            stub_projects = pd.DataFrame(
                {
                    "_key": missing_projects,
                    "_name": missing_projects,
                    "_budget": float("nan"),
                    "_cost": float("nan"),
                    "_progress": float("nan"),
                }
            )

            P = pd.concat(
                [P, stub_projects],
                ignore_index=True,
            )

    # ------------------------------------------------------------------------
    # DETAILS
    # ------------------------------------------------------------------------

    dsum = None

    if D is not None and not D.empty and not L.empty:
        detail_columns = cols_of(D, DETAIL_COLS)

        detail_pipeline_id = detail_columns["id"]
        detail_pipeline_name = detail_columns["name"]

        if detail_pipeline_id:
            D["_pid"] = D[detail_pipeline_id].map(normalize_text)
        elif detail_pipeline_name:
            D["_pid"] = D[detail_pipeline_name].map(normalize_text)
        else:
            warnings.append(
                "Pipeline Details sheet has no Pipeline ID or Pipeline Name."
            )
            D = None

        if D is not None:
            detail_project_id = detail_columns["proj_id"]
            detail_project_name = detail_columns["proj_name"]

            if detail_project_id:
                D["_project_key"] = (
                    D[detail_project_id].map(normalize_text)
                )
            elif detail_project_name:
                D["_project_key"] = (
                    D[detail_project_name].map(
                        lambda x: name_to_key.get(
                            normalize_key(x),
                            normalize_text(x),
                        )
                    )
                )
            else:
                D["_project_key"] = None

            if D["_project_key"].notna().any():
                D["_pk"] = (
                    D["_project_key"].astype(str)
                    + "||"
                    + D["_pid"].astype(str)
                )
            else:
                pipeline_map = (
                    L.groupby("_pid")["_pk"]
                    .agg(list)
                    .to_dict()
                )

                D["_pk"] = D["_pid"].map(
                    lambda pid:
                    pipeline_map[pid][0]
                    if pid in pipeline_map
                    and len(pipeline_map[pid]) == 1
                    else None
                )

            valid_pipeline_keys = set(L["_pk"].tolist())

            unmatched = ~D["_pk"].isin(valid_pipeline_keys)
            unmatched_count = int(unmatched.sum())

            if unmatched_count:
                warnings.append(
                    f"{unmatched_count} detail row(s) could not be matched."
                )

            D = D[D["_pk"].isin(valid_pipeline_keys)].copy()

            D["_planned"] = (
                to_num(D[detail_columns["planned"]])
                if detail_columns["planned"]
                else float("nan")
            )

            D["_actual"] = (
                to_num(D[detail_columns["actual"]])
                if detail_columns["actual"]
                else float("nan")
            )

            category_column = (
                detail_columns["category"]
                or detail_columns["item"]
            )

            if category_column:
                D["_cat"] = (
                    D[category_column]
                    .fillna("Unknown")
                    .astype(str)
                )
            else:
                D["_cat"] = "All"

            if detail_columns["date"]:
                D["_date"] = pd.to_datetime(
                    D[detail_columns["date"]],
                    errors="coerce",
                )
            else:
                D["_date"] = pd.NaT

            dsum = (
                D.groupby("_pk")[["_planned", "_actual"]]
                .sum(min_count=1)
            )

    # ------------------------------------------------------------------------
    # FINANCIAL MODEL
    # ------------------------------------------------------------------------

    F = pd.DataFrame()

    if not L.empty:
        F = pd.DataFrame(
            {
                "_pk": L["_pk"],
                "_key": L["_project_key"],
                "Pipeline": L["_name"],
            }
        ).reset_index(drop=True)

        L = L.reset_index(drop=True)

        nan_series = pd.Series(
            float("nan"),
            index=L.index,
        )

        if pipeline_columns["budget"]:
            F["Budget"] = to_num(
                L[pipeline_columns["budget"]]
            )
        else:
            F["Budget"] = nan_series.copy()

        if pipeline_columns["planned"]:
            planned = to_num(
                L[pipeline_columns["planned"]]
            )
        else:
            planned = nan_series.copy()

        if dsum is not None:
            planned = planned.fillna(
                L["_pk"].map(dsum["_planned"])
            )

        F["Planned"] = planned

        if pipeline_columns["actual"]:
            actual = to_num(
                L[pipeline_columns["actual"]]
            )
        else:
            actual = nan_series.copy()

        if dsum is not None:
            actual = actual.fillna(
                L["_pk"].map(dsum["_actual"])
            )

        F["Actual"] = actual

        if pipeline_columns["remaining"]:
            remaining = to_num(
                L[pipeline_columns["remaining"]]
            )
        else:
            remaining = pd.Series(
                float("nan"),
                index=L.index,
            )

        remaining = (
            remaining
            .fillna(F["Budget"] - F["Actual"])
            .fillna(F["Planned"] - F["Actual"])
        )

        F["Remaining"] = remaining
        F["Variance"] = F["Planned"] - F["Actual"]

        if pipeline_columns["progress"]:
            F["Progress"] = to_pct(
                L[pipeline_columns["progress"]]
            )
        else:
            F["Progress"] = nan_series.copy()

            valid = (
                F["Budget"].notna()
                & (F["Budget"] > 0)
                & F["Actual"].notna()
            )

            F.loc[valid, "Progress"] = (
                F.loc[valid, "Actual"]
                / F.loc[valid, "Budget"]
                * 100
            )

        F["Progress"] = F["Progress"].clip(0, 100)

        if pipeline_columns["end"]:
            end_dates = pd.to_datetime(
                L[pipeline_columns["end"]],
                errors="coerce",
            )
        else:
            end_dates = pd.Series(
                pd.NaT,
                index=L.index,
            )

        raw_status = (
            L[pipeline_columns["status"]]
            if pipeline_columns["status"]
            else pd.Series("", index=L.index)
        )

        F["State"] = [
            pipeline_state(
                status,
                progress,
                end_date,
            )
            for status, progress, end_date in zip(
                raw_status,
                F["Progress"],
                end_dates,
            )
        ]

    required_financial_columns = {
        "Budget": float("nan"),
        "Planned": float("nan"),
        "Actual": float("nan"),
        "Remaining": float("nan"),
        "Variance": float("nan"),
        "Progress": float("nan"),
        "State": "Unknown",
    }

    if F is None or not isinstance(F, pd.DataFrame):
        F = pd.DataFrame()

    for col, default in required_financial_columns.items():
        if col not in F.columns:
            F[col] = default

    for col in [
        "Budget",
        "Planned",
        "Actual",
        "Remaining",
        "Variance",
        "Progress",
    ]:
        F[col] = pd.to_numeric(
            F[col],
            errors="coerce",
        )

    if not F.empty:
        missing_remaining = F["Remaining"].isna()

        F.loc[missing_remaining, "Remaining"] = (
            F.loc[missing_remaining, "Budget"]
            - F.loc[missing_remaining, "Actual"]
        )

        missing_variance = F["Variance"].isna()

        F.loc[missing_variance, "Variance"] = (
            F.loc[missing_variance, "Planned"]
            - F.loc[missing_variance, "Actual"]
        )

        missing_progress = F["Progress"].isna()

        valid_progress = (
            F["Budget"].notna()
            & (F["Budget"] > 0)
            & F["Actual"].notna()
        )

        F.loc[
            missing_progress & valid_progress,
            "Progress",
        ] = (
            F.loc[
                missing_progress & valid_progress,
                "Actual",
            ]
            / F.loc[
                missing_progress & valid_progress,
                "Budget",
            ]
            * 100
        )

    role = {}

    if project_sheet:
        role["projects"] = project_sheet

    if pipeline_sheet:
        role["pipelines"] = pipeline_sheet

    if detail_sheet:
        role["details"] = detail_sheet

    return {
        "ok": True,
        "role": role,
        "other": other_sheets,
        "warnings": warnings,
        "P": P,
        "L": L,
        "F": F,
        "D": D,
    }


# ============================================================================
# 16) DISPLAY HELPERS
# ============================================================================

def summarize(F, budget=float("nan"), cost=float("nan")):
    if F is None or F.empty:
        return {
            "Total Budget": budget,
            "Total Planned Cost": float("nan"),
            "Total Actual Cost": cost,
            "Remaining Cost": float("nan"),
            "Overall Progress": float("nan"),
            "Pipelines": 0,
            "Completed": 0,
            "Active": 0,
            "Delayed": 0,
        }

    total_budget = budget

    if pd.isna(total_budget):
        total_budget = F["Budget"].sum(min_count=1)

    total_planned = F["Planned"].sum(min_count=1)
    total_actual = F["Actual"].sum(min_count=1)

    if pd.isna(total_actual) and pd.notna(cost):
        total_actual = cost

    if pd.notna(total_budget) and pd.notna(total_actual):
        remaining = total_budget - total_actual
    elif pd.notna(total_planned) and pd.notna(total_actual):
        remaining = total_planned - total_actual
    else:
        remaining = float("nan")

    weights = F["Budget"].where(
        F["Budget"] > 0,
        F["Planned"],
    )

    valid = (
        F["Progress"].notna()
        & weights.notna()
        & (weights > 0)
    )

    if valid.any():
        overall_progress = (
            (
                F.loc[valid, "Progress"]
                * weights.loc[valid]
            ).sum()
            / weights.loc[valid].sum()
        )
    elif F["Progress"].notna().any():
        overall_progress = F["Progress"].mean()
    else:
        overall_progress = float("nan")

    return {
        "Total Budget": total_budget,
        "Total Planned Cost": total_planned,
        "Total Actual Cost": total_actual,
        "Remaining Cost": remaining,
        "Overall Progress": overall_progress,
        "Pipelines": len(F),
        "Completed": int((F["State"] == "Completed").sum()),
        "Active": int((F["State"] == "Active").sum()),
        "Delayed": int((F["State"] == "Delayed").sum()),
    }


def format_value(value):
    if pd.isna(value):
        return "—"

    if isinstance(value, (pd.Timestamp, pd.DatetimeIndex)):
        return pd.Timestamp(value).strftime("%Y-%m-%d")

    if isinstance(value, (int, float)):
        return fmt(value)

    return str(value)


def show_kpis(data):
    items = []

    for key, value in data.items():
        if pd.isna(value):
            continue

        if "Progress" in key and isinstance(value, (int, float)):
            display = f"{value:,.1f}%"
        elif isinstance(value, (int, float)):
            display = fmt(value)
        else:
            display = str(value)

        items.append((key, display))

    kpi_row(items)


def info_grid(row, columns):
    items = []

    for column in columns:
        if column.startswith("_"):
            continue

        if column not in row.index:
            continue

        value = row[column]

        if pd.isna(value):
            continue

        items.append(
            (
                column,
                format_value(value),
            )
        )

    if not items:
        return

    columns_ui = st.columns(3)

    for i, (key, value) in enumerate(items):
        columns_ui[i % 3].markdown(
            f"**{key}**  \n{value}"
        )


def cost_chart(
    dataframe,
    x="Pipeline",
    title="Budget vs Planned vs Actual",
):
    if dataframe is None or dataframe.empty:
        return None

    available = [
        column
        for column in [
            "Budget",
            "Planned",
            "Actual",
        ]
        if column in dataframe.columns
        and dataframe[column].notna().any()
    ]

    if not available:
        return None

    long = dataframe.melt(
        id_vars=x,
        value_vars=available,
        var_name="Type",
        value_name="Cost",
    )

    return px.bar(
        long,
        x=x,
        y="Cost",
        color="Type",
        barmode="group",
        title=title,
        color_discrete_map={
            "Budget": DARK,
            "Planned": "#8A94A0",
            "Actual": PRIMARY,
        },
    )


# ============================================================================
# 17) HOME
# ============================================================================

def page_home():
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=180)

    st.markdown(
        f"""
        <div class="hero">
            <span class="co">{COMPANY}</span>
            <h1>Project & Pipeline Management</h1>
            <p>
                Monitor projects, pipelines, costs,
                deliveries and progress.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not M:
        st.error("Could not load Google Sheets data.")
        return

    if M.get("mode") == "formwork":
        st.success("Formwork Construction Types workbook detected.")
        st.subheader("🏗️ Construction Types")

        grid(
            [
                (name, name)
                for name in M.get("formwork_sheets", [])
            ],
            open_sheet,
            "formwork_sheet",
            3,
        )
        return

    if not M.get("ok"):
        st.warning(
            "No Projects or Pipelines structure was detected."
        )

        names = get_sheet_names()

        if names:
            st.subheader("Available Sheets")

            grid(
                [(name, name) for name in names],
                open_sheet,
                "home_sheet",
                4,
            )

        return

    project_count = len(M["P"])
    pipeline_count = len(M["F"])

    total_budget = M["P"]["_budget"].sum(
        min_count=1
    )

    total_actual = M["F"]["Actual"].sum(
        min_count=1
    )

    delayed = int(
        (M["F"]["State"] == "Delayed").sum()
    )

    completed = int(
        (M["F"]["State"] == "Completed").sum()
    )

    kpi_row(
        [
            ("Projects", project_count),
            ("Pipelines", pipeline_count),
            ("Total Budget", fmt(total_budget)),
            ("Total Actual Cost", fmt(total_actual)),
            ("Delayed Pipelines", delayed),
            ("Completed Pipelines", completed),
        ]
    )

    if not VIEWER_MODE:
        st.button(
            "📁 Open Projects",
            on_click=go,
            args=("projects",),
            width="stretch",
        )

    st.subheader("📊 Executive Analytics")

    F = M["F"]

    c1, c2 = st.columns(2)

    chart = cost_chart(F)

    if chart:
        c1.plotly_chart(chart, width="stretch")

    if not F.empty:
        status_count = (
            F["State"]
            .value_counts()
            .reset_index()
        )

        status_count.columns = ["State", "Count"]

        c2.plotly_chart(
            px.pie(
                status_count,
                names="State",
                values="Count",
                hole=.45,
                title="Pipelines by Status",
                color="State",
                color_discrete_map=STATUS_COLORS,
            ),
            width="stretch",
        )

        if F["Progress"].notna().any():
            st.plotly_chart(
                px.bar(
                    F,
                    x="Progress",
                    y="Pipeline",
                    orientation="h",
                    title="Pipeline Progress",
                    range_x=[0, 100],
                    color_discrete_sequence=[PRIMARY],
                ),
                width="stretch",
            )

    st.subheader("📋 Pipeline Overview")

    table = F.drop(
        columns=["_pk", "_key"],
        errors="ignore",
    )

    st.dataframe(
        table,
        width="stretch",
        hide_index=True,
    )

    if M.get("warnings") and not VIEWER_MODE:
        for warning in M["warnings"]:
            st.warning("⚠ " + warning)


# ============================================================================
# 18) PROJECTS
# ============================================================================

def page_projects():
    st.title("📁 Projects")

    P = M["P"]

    if P.empty:
        st.warning("No projects were found.")
        return

    st.caption(
        f"{len(P)} project(s) detected from Google Sheets."
    )

    items = []

    for _, row in P.iterrows():
        name = row["_name"] or row["_key"]

        items.append(
            (
                f"📂 {name}",
                row["_key"],
            )
        )

    grid(
        items,
        open_project,
        "project",
        3,
    )


# ============================================================================
# 19) PROJECT
# ============================================================================

def page_project():
    project_key = normalize_text(ss.project)

    P = M.get("P", pd.DataFrame())
    F = M.get("F", pd.DataFrame())

    if P is None or P.empty:
        st.error("No project data is available.")
        return

    project_rows = P[
        P["_key"].astype(str).map(normalize_text)
        == project_key
    ].copy()

    if project_rows.empty:
        project_rows = P[
            P["_name"].astype(str).map(normalize_key)
            == normalize_key(project_key)
        ].copy()

    if project_rows.empty:
        st.error(f"Project '{project_key}' was not found.")
        return

    project_row = project_rows.iloc[0]

    if F is None or F.empty:
        pipeline_rows = pd.DataFrame()
    else:
        pipeline_rows = F[
            F["_key"].astype(str).map(normalize_text)
            == project_key
        ].copy()

    project_name = project_row["_name"]

    st.markdown(
        f"""
        <div class="topbar">
            {COMPANY} | PROJECT
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.title(f"📂 {project_name}")

    st.subheader("Project Information")

    info_grid(
        project_row,
        list(P.columns),
    )

    if pipeline_rows.empty:
        st.warning(
            "This project has no pipelines in the workbook."
        )
        return

    summary = summarize(
        pipeline_rows,
        project_row["_budget"],
        project_row["_cost"],
    )

    show_kpis(summary)

    st.subheader("🔧 Pipelines")

    pipeline_buttons = [
        (
            f"🔧 {row['Pipeline']}",
            row["_pk"],
        )
        for _, row in pipeline_rows.iterrows()
    ]

    grid(
        pipeline_buttons,
        open_pipeline,
        "pipeline",
        3,
    )

    st.subheader("📊 Project Analytics")

    c1, c2 = st.columns(2)

    chart = cost_chart(pipeline_rows)

    if chart:
        c1.plotly_chart(
            chart,
            width="stretch",
        )

    status_count = (
        pipeline_rows["State"]
        .value_counts()
        .reset_index()
    )

    status_count.columns = ["State", "Count"]

    c2.plotly_chart(
        px.pie(
            status_count,
            names="State",
            values="Count",
            hole=.45,
            title="Pipelines by Status",
            color="State",
            color_discrete_map=STATUS_COLORS,
        ),
        width="stretch",
    )

    if pipeline_rows["Progress"].notna().any():
        st.plotly_chart(
            px.bar(
                pipeline_rows,
                x="Progress",
                y="Pipeline",
                orientation="h",
                title="Progress % by Pipeline",
                range_x=[0, 100],
                color_discrete_sequence=[PRIMARY],
            ),
            width="stretch",
        )

    st.subheader("📋 Pipelines Table")

    table = pipeline_rows.drop(
        columns=["_pk", "_key"],
        errors="ignore",
    )

    st.dataframe(
        table,
        width="stretch",
        hide_index=True,
    )


# ============================================================================
# 20) PIPELINE
# ============================================================================

def page_pipeline():
    pipeline_key = ss.pipeline

    F = M["F"]
    L = M["L"]
    D = M["D"]

    pipeline_rows = F[
        F["_pk"] == pipeline_key
    ]

    if pipeline_rows.empty:
        st.error("Pipeline was not found.")
        return

    pipeline = pipeline_rows.iloc[0]

    raw_rows = L[
        L["_pk"] == pipeline_key
    ]

    raw = (
        pipeline
        if raw_rows.empty
        else raw_rows.iloc[0]
    )

    project_name = pipeline.get("_key", "")

    P = M.get("P", pd.DataFrame())

    if P is not None and not P.empty:
        matches = P[
            P["_key"].astype(str).map(normalize_text)
            == normalize_text(pipeline.get("_key", ""))
        ]

        if not matches.empty:
            project_name = matches.iloc[0]["_name"]

    st.markdown(
        f"""
        <div class="topbar">
            {COMPANY} | {project_name} | PIPELINE
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not VIEWER_MODE:
        st.button(
            f"⬅ Back to {project_name}",
            on_click=go,
            args=("project",),
            width="stretch",
        )

    st.title(f"🔧 {pipeline['Pipeline']}")

    st.subheader("Pipeline Information")

    info_grid(
        raw,
        list(raw.index),
    )

    show_kpis(
        {
            "Budget": pipeline["Budget"],
            "Planned Cost": pipeline["Planned"],
            "Actual Cost": pipeline["Actual"],
            "Remaining Cost": pipeline["Remaining"],
            "Cost Variance": pipeline["Variance"],
            "Progress %": pipeline["Progress"],
            "Status": pipeline["State"],
        }
    )

    progress = pipeline["Progress"]

    if pd.notna(progress):
        progress_value = min(
            max(float(progress), 0),
            100,
        )

        st.progress(progress_value / 100)

        st.caption(
            f"Pipeline Progress: {progress_value:.1f}%"
        )

    if pd.notna(pipeline["Variance"]) and pipeline["Variance"] < 0:
        st.error(
            "⚠ Actual cost exceeds the planned cost."
        )
    elif (
        pd.notna(pipeline["Variance"])
        and pipeline["Variance"] > 0
    ):
        st.success(
            "✓ Actual cost is currently below planned cost."
        )

    st.subheader("📊 Pipeline Analytics")

    c1, c2 = st.columns(2)

    chart = cost_chart(
        F[F["_pk"] == pipeline_key],
        title="Budget vs Planned vs Actual",
    )

    if chart:
        c1.plotly_chart(
            chart,
            width="stretch",
        )

    if D is not None and not D.empty:
        details = D[
            D["_pk"] == pipeline_key
        ].copy()

        if not details.empty:
            category = (
                details.groupby("_cat")[
                    ["_planned", "_actual"]
                ]
                .sum(min_count=1)
                .reset_index()
                .rename(
                    columns={
                        "_cat": "Category",
                        "_planned": "Planned",
                        "_actual": "Actual",
                    }
                )
            )

            long = category.melt(
                id_vars="Category",
                var_name="Type",
                value_name="Cost",
            )

            c2.plotly_chart(
                px.bar(
                    long,
                    x="Category",
                    y="Cost",
                    color="Type",
                    barmode="group",
                    title="Cost by Category",
                ),
                width="stretch",
            )

            if (
                details["_date"].notna().any()
                and details["_actual"].notna().any()
            ):
                timeline = (
                    details.dropna(subset=["_date"])
                    .sort_values("_date")
                    .copy()
                )

                timeline["Cumulative Actual"] = (
                    timeline["_actual"]
                    .fillna(0)
                    .cumsum()
                )

                st.plotly_chart(
                    px.line(
                        timeline,
                        x="_date",
                        y="Cumulative Actual",
                        markers=True,
                        title="Cumulative Actual Cost",
                        color_discrete_sequence=[PRIMARY],
                    ),
                    width="stretch",
                )

            kpi_row(
                [
                    ("Detail Records", len(details)),
                    (
                        "Detail Planned",
                        fmt(details["_planned"].sum(min_count=1)),
                    ),
                    (
                        "Detail Actual",
                        fmt(details["_actual"].sum(min_count=1)),
                    ),
                ]
            )

            st.subheader("📋 Pipeline Cost Details")

            original_columns = [
                c for c in details.columns
                if not c.startswith("_")
            ]

            st.dataframe(
                details[original_columns]
                if original_columns
                else details,
                width="stretch",
                hide_index=True,
            )
        else:
            st.info(
                "No cost/detail records were found for this pipeline."
            )
    else:
        st.info(
            "Add a Pipeline Details / Costs sheet to display detailed costs."
        )


# ============================================================================
# 21) GENERIC SHEET
# ============================================================================

def show_dashboard(df, sheet_name):
    st.title(f"{sheet_name} Dashboard")

    if df.empty:
        st.warning("This sheet is empty.")
        return

    is_delivery = (
        "deliver" in sheet_name.lower()
        or (
            find_col(df, "expected_date")
            and find_col(df, "required_date")
        )
    )

    if is_delivery:
        show_delivery_dashboard(df, sheet_name)
    else:
        show_generic_dashboard(df, sheet_name)


def page_sheet():
    if not VIEWER_MODE:
        st.button(
            "⬅ Home",
            on_click=go,
            args=("home",),
            width="stretch",
        )

    if not ss.sheet:
        go("home")
        return

    dataframe = load_google_sheet(ss.sheet)

    if (
        M
        and M.get("mode") == "formwork"
        and ss.sheet in M.get("formwork_sheets", [])
    ):
        show_formwork_dashboard(
            dataframe,
            ss.sheet,
        )
    else:
        show_dashboard(
            dataframe,
            ss.sheet,
        )


# ============================================================================
# 22) VIEWER DASHBOARD
# ============================================================================

def page_viewer_dashboard():
    """
    End-user dashboard.
    No upload, no sidebar navigation, no project buttons.
    """

    if not M:
        st.error(
            "Dashboard data is not available."
        )
        return

    # Formwork viewer:
    # show the first detected construction type.
    if M.get("mode") == "formwork":
        sheets = M.get("formwork_sheets", [])

        if not sheets:
            st.error("No Formwork sheet was found.")
            return

        sheet_name = sheets[0]
        df = load_google_sheet(sheet_name)

        show_formwork_dashboard(
            df,
            sheet_name,
        )
        return

    # Project / pipeline viewer
    if M.get("ok"):
        page_home()
        return

    # Fallback
    names = get_sheet_names()

    if names:
        df = load_google_sheet(names[0])
        show_dashboard(df, names[0])
    else:
        st.error("No Google Sheets data was found.")


# ============================================================================
# 23) LOAD MODEL
# ============================================================================

M = None

try:
    M = build_model()

except Exception as error:
    st.error(
        f"Could not read Google Sheets: {error}"
    )

    if not VIEWER_MODE:
        st.info(
            "Check GOOGLE_SHEET_URL and Google Service Account "
            "credentials in Streamlit Secrets."
        )


# ============================================================================
# 24) ADMIN SIDEBAR
# ============================================================================

if not VIEWER_MODE:

    with st.sidebar:

        if os.path.exists(LOGO_PATH):
            st.image(
                LOGO_PATH,
                width=150,
            )
        else:
            st.markdown(
                f"## {COMPANY}"
            )

        st.header("🏗️ Navigation")

        st.button(
            "🏠 Home",
            on_click=go,
            args=("home",),
            width="stretch",
        )

        if M and M.get("mode") == "formwork":

            st.subheader(
                "🏗️ Formwork Construction Types"
            )

            for i, sheet_name in enumerate(
                M.get("formwork_sheets", [])
            ):
                st.button(
                    sheet_name,
                    key=f"sidebar_formwork_{i}",
                    on_click=open_sheet,
                    args=(sheet_name,),
                    width="stretch",
                )

        if M and M.get("ok"):

            st.button(
                "📁 Projects",
                on_click=go,
                args=("projects",),
                width="stretch",
            )

            project_map = (
                M["P"]
                .set_index("_key")["_name"]
                .to_dict()
            )

            if ss.project in project_map:
                st.button(
                    f"📂 {project_map[ss.project]}",
                    on_click=go,
                    args=("project",),
                    width="stretch",
                )

            if (
                ss.pipeline
                and ss.pipeline in set(M["F"]["_pk"])
            ):
                st.button(
                    "🔧 Current Pipeline",
                    on_click=go,
                    args=("pipeline",),
                    width="stretch",
                )

        if st.button(
            "🔄 Refresh Google Sheets",
            width="stretch",
        ):
            st.cache_data.clear()
            st.rerun()


# ============================================================================
# 25) MAIN ROUTER
# ============================================================================

try:

    if VIEWER_MODE:
        # End user sees dashboard only.
        page_viewer_dashboard()

    elif not M:
        st.error(
            "No Google Sheets data is available."
        )

    elif ss.page == "home":
        page_home()

    elif ss.page == "projects" and M.get("ok"):
        page_projects()

    elif ss.page == "project" and ss.project and M.get("ok"):
        page_project()

    elif (
        ss.page == "pipeline"
        and ss.pipeline
        and M.get("ok")
    ):
        page_pipeline()

    elif ss.page == "sheet" and ss.sheet:
        page_sheet()

    else:
        page_home()

except Exception as error:
    st.error(
        f"Something went wrong while building this page: {error}"
    )

    if not VIEWER_MODE:
        st.button(
            "⬅ Back to Home",
            on_click=go,
            args=("home",),
            width="stretch",
        )
'''
path = Path("/mnt/data/app_google_sheets.py")
path.write_text(code, encoding="utf-8")

requirements = """streamlit
pandas
plotly
gspread
google-auth
"""
req_path = Path("/mnt/data/requirements.txt")
req_path.write_text(requirements, encoding="utf-8")

print(f"Created: {path}")
print(f"Created: {req_path}")
