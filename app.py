"""
Hassan Allam Construction
Construction Projects & Pipeline Dashboard
Streamlit MVP - Single File

Excel
  └── Projects
       └── Pipelines
            └── Pipeline Details / Costs

Also supports:
- Delivery Dashboard
- Generic Dashboard
- Dynamic Excel sheets
- Dynamic Projects
- Dynamic Pipelines
- Dynamic Pipeline Details
"""

import os
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st


# ============================================================================
# 1) PAGE CONFIG
# ============================================================================

st.set_page_config(
    page_title="Hassan Allam Construction | Projects Dashboard",
    page_icon="🏗️",
    layout="wide"
)


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
# 3) GENERAL EXCEL COLUMN MAP
# ============================================================================

COLUMN_MAP = {

    "project": [
        "Project",
        "Project Name",
        "ProjectName"
    ],

    "item": [
        "Item",
        "Material",
        "Description"
    ],

    "location": [
        "Location",
        "Site",
        "Zone"
    ],

    "required_qty": [
        "Required Qty",
        "Required Quantity",
        "Quantity",
        "Qty"
    ],

    "delivered_qty": [
        "Delivered Qty",
        "Delivered Quantity"
    ],

    "remaining_qty": [
        "Remaining Qty",
        "Remaining Quantity"
    ],

    "required_date": [
        "Required Date"
    ],

    "expected_date": [
        "Expected Date"
    ],

    "actual_date": [
        "Actual Date",
        "Delivery Date"
    ],

    "status": [
        "Status"
    ],
}


# ============================================================================
# 4) PROJECT / PIPELINE COLUMN MAP
# ============================================================================

PROJECT_COLS = {

    "id": [
        "Project ID",
        "ProjectID",
        "Project Code",
        "Project No",
        "Project Number",
        "Code"
    ],

    "name": [
        "Project Name",
        "Project",
        "Name"
    ],

    "client": [
        "Client",
        "Client Name",
        "Owner"
    ],

    "manager": [
        "Project Manager",
        "Manager",
        "Project Manager Name"
    ],

    "location": [
        "Location",
        "Site",
        "Project Location"
    ],

    "start": [
        "Start Date",
        "Project Start Date",
        "Start"
    ],

    "end": [
        "End Date",
        "Project End Date",
        "Finish Date",
        "Finish"
    ],

    "status": [
        "Status",
        "Project Status"
    ],

    "budget": [
        "Total Budget",
        "Budget",
        "Project Budget"
    ],

    "cost": [
        "Total Cost",
        "Actual Cost",
        "Actual",
        "Project Cost"
    ],

    "progress": [
        "Progress %",
        "Progress",
        "% Complete",
        "Completion %"
    ],

    "description": [
        "Description",
        "Project Description"
    ]
}


PIPELINE_COLS = {

    "proj_id": PROJECT_COLS["id"],

    "proj_name": [
        "Project Name",
        "Project"
    ],

    "id": [
        "Pipeline ID",
        "PipelineID",
        "Pipeline Code",
        "Pipeline No",
        "Pipeline Number"
    ],

    "name": [
        "Pipeline Name",
        "Pipeline",
        "Name"
    ],

    "desc": [
        "Description",
        "Pipeline Description"
    ],

    "budget": [
        "Budget",
        "Total Budget",
        "Pipeline Budget"
    ],

    "planned": [
        "Planned Cost",
        "Planned",
        "Planned Amount"
    ],

    "actual": [
        "Actual Cost",
        "Actual",
        "Actual Amount",
        "Spent"
    ],

    "remaining": [
        "Remaining Cost",
        "Remaining",
        "Remaining Budget"
    ],

    "progress": [
        "Progress %",
        "Progress",
        "% Complete",
        "Completion %"
    ],

    "status": [
        "Status",
        "Pipeline Status"
    ],

    "start": [
        "Start Date",
        "Pipeline Start Date",
        "Start"
    ],

    "end": [
        "End Date",
        "Finish Date",
        "Pipeline End Date",
        "End"
    ],

    "category": [
        "Category",
        "Type",
        "Pipeline Category"
    ]
}


DETAIL_COLS = {

    "proj_id": PROJECT_COLS["id"],

    "proj_name": [
        "Project Name",
        "Project"
    ],

    "id": PIPELINE_COLS["id"],

    "name": [
        "Pipeline Name",
        "Pipeline"
    ],

    "planned": [
        "Planned Cost",
        "Planned",
        "Planned Amount"
    ],

    "actual": [
        "Actual Cost",
        "Actual",
        "Actual Amount",
        "Amount",
        "Spent"
    ],

    "category": [
        "Category",
        "Cost Type",
        "Type"
    ],

    "item": [
        "Item",
        "Description",
        "Material"
    ],

    "date": [
        "Date",
        "Cost Date",
        "Transaction Date"
    ]
}


# ============================================================================
# 5) CSS
# ============================================================================

st.markdown(
    f"""
<style>

.stApp {{
    background-color: #ffffff;
    background-image:
        linear-gradient(rgba(31,41,51,.035) 1px, transparent 1px),
        linear-gradient(90deg, rgba(31,41,51,.035) 1px, transparent 1px);
    background-size: 32px 32px;
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

</style>
""",
    unsafe_allow_html=True
)


# ============================================================================
# 6) SESSION STATE
# ============================================================================

ss = st.session_state

DEFAULT_STATE = {
    "page": "home",
    "file_bytes": None,
    "file_name": None,
    "sheet": None,
    "project": None,
    "pipeline": None,
}

for key, value in DEFAULT_STATE.items():
    ss.setdefault(key, value)


# ============================================================================
# 7) EXCEL DATA LAYER
# ============================================================================

@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes) -> list:
    """
    Return all sheet names from workbook.
    """
    return pd.ExcelFile(
        BytesIO(file_bytes)
    ).sheet_names


@st.cache_data(show_spinner=False)
def load_excel_data(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    """
    Read Excel sheet and automatically detect header row.
    """

    raw = pd.read_excel(
        BytesIO(file_bytes),
        sheet_name=sheet_name,
        header=None,
        nrows=30
    )

    if raw.empty:
        return pd.DataFrame()

    counts = raw.notna().sum(axis=1)

    hdr = 0

    if len(counts) and counts.max() > 0:

        max_count = counts.max()

        if max_count >= 2:

            candidates = [
                i
                for i, c in enumerate(counts)
                if c >= max(2, 0.6 * max_count)
            ]

            if candidates:
                hdr = candidates[0]

    df = pd.read_excel(
        BytesIO(file_bytes),
        sheet_name=sheet_name,
        header=hdr
    )

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
# 8) GENERAL HELPERS
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

    return pick(
        df,
        COLUMN_MAP.get(key, [])
    )


def to_num(series):

    if series is None:
        return pd.Series(dtype=float)

    if pd.api.types.is_numeric_dtype(series):
        return series

    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True),
        errors="coerce"
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

    max_value = s.dropna().max()

    if max_value <= 1:
        return s * 100

    return s


def safe_sum(series):

    if series is None:
        return float("nan")

    return series.sum(min_count=1)


# ============================================================================
# 9) KPI COMPONENTS
# ============================================================================

def kpi(col, label, value):

    col.markdown(
        f"""
        <div class="kpi">
            <div class="l">{label}</div>
            <div class="v">{value}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def kpi_row(items):

    for i in range(0, len(items), 3):

        columns = st.columns(3)

        for column, item in zip(
            columns,
            items[i:i + 3]
        ):

            label, value = item

            kpi(
                column,
                label,
                value
            )


# ============================================================================
# 10) NAVIGATION
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

    for i in range(
        0,
        len(items),
        per_row
    ):

        columns = st.columns(per_row)

        batch = items[
            i:i + per_row
        ]

        for j, (label, arg) in enumerate(batch):

            columns[j].button(
                label,
                key=f"{prefix}_{i+j}",
                on_click=callback,
                args=(arg,),
                width="stretch"
            )


# ============================================================================
# 11) DELIVERY DASHBOARD
# ============================================================================

def compute_delivery_status(df, columns):

    def status(row):

        actual_col = columns.get("actual_date")
        expected_col = columns.get("expected_date")
        required_col = columns.get("required_date")

        if (
            actual_col
            and pd.notna(row.get(actual_col))
        ):
            return "Delivered"

        expected = (
            row.get(expected_col)
            if expected_col
            else pd.NaT
        )

        required = (
            row.get(required_col)
            if required_col
            else pd.NaT
        )

        if (
            pd.notna(expected)
            and pd.notna(required)
            and expected > required
        ):
            return "Delayed"

        if pd.notna(expected):
            return "Pending / On Track"

        return "Unknown"

    return df.apply(
        status,
        axis=1
    )


def show_delivery_dashboard(df, sheet_name):

    columns = {
        key: find_col(df, key)
        for key in COLUMN_MAP
    }

    df = df.copy()

    # ------------------------------------------------------------------------
    # Dates
    # ------------------------------------------------------------------------

    for key in (
        "required_date",
        "expected_date",
        "actual_date"
    ):

        column = columns[key]

        if column:

            df[column] = pd.to_datetime(
                df[column],
                errors="coerce"
            )

    # ------------------------------------------------------------------------
    # Numeric columns
    # ------------------------------------------------------------------------

    for key in (
        "required_qty",
        "delivered_qty",
        "remaining_qty"
    ):

        column = columns[key]

        if column:

            df[column] = to_num(
                df[column]
            )

    # ------------------------------------------------------------------------
    # Warnings
    # ------------------------------------------------------------------------

    warnings = [

        (
            "expected_date",
            "Expected delivery KPIs cannot be calculated."
        ),

        (
            "required_date",
            "Delay detection cannot be calculated."
        ),

        (
            "required_qty",
            "Remaining quantity cannot be calculated."
        ),

        (
            "project",
            "Project filter/chart is unavailable."
        )
    ]

    for key, message in warnings:

        if not columns[key]:

            st.warning(
                f"⚠ {COLUMN_MAP[key][0]} column was not found. "
                f"{message}"
            )

    # ------------------------------------------------------------------------
    # Status
    # ------------------------------------------------------------------------

    if (
        columns["actual_date"]
        or columns["expected_date"]
    ):

        df["Computed Status"] = compute_delivery_status(
            df,
            columns
        )

        status_column = "Computed Status"

    elif columns["status"]:

        status_column = columns["status"]

    else:

        status_column = None

    # ------------------------------------------------------------------------
    # Remaining quantity
    # ------------------------------------------------------------------------

    if (
        columns["required_qty"]
        and columns["delivered_qty"]
    ):

        df["Remaining Qty (calc)"] = (
            df[columns["required_qty"]]
            - df[columns["delivered_qty"]].fillna(0)
        ).clip(lower=0)

        remaining_column = "Remaining Qty (calc)"

    else:

        remaining_column = columns["remaining_qty"]

    # ------------------------------------------------------------------------
    # Filters
    # ------------------------------------------------------------------------

    st.subheader("🔎 Filters")

    filter_columns = st.columns(5)

    filter_items = [
        ("Project", columns["project"]),
        ("Status", status_column),
        ("Item", columns["item"]),
        ("Location", columns["location"])
    ]

    for column_ui, (label, column) in zip(
        filter_columns,
        filter_items
    ):

        if column:

            options = sorted(
                df[column]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

            selected = column_ui.multiselect(
                label,
                options,
                key=f"delivery_{sheet_name}_{label}"
            )

            if selected:

                df = df[
                    df[column]
                    .astype(str)
                    .isin(selected)
                ]

    if (
        columns["expected_date"]
        and df[columns["expected_date"]].notna().any()
    ):

        min_date = (
            df[columns["expected_date"]]
            .min()
            .date()
        )

        max_date = (
            df[columns["expected_date"]]
            .max()
            .date()
        )

        if min_date < max_date:

            date_range = filter_columns[4].date_input(
                "Expected Date",
                (min_date, max_date),
                key=f"delivery_date_{sheet_name}"
            )

            if (
                isinstance(date_range, tuple)
                and len(date_range) == 2
            ):

                mask = (
                    df[columns["expected_date"]]
                    .dt.date
                    .between(
                        date_range[0],
                        date_range[1]
                    )
                    | df[columns["expected_date"]].isna()
                )

                df = df[mask]

    # ------------------------------------------------------------------------
    # KPIs
    # ------------------------------------------------------------------------

    today = pd.Timestamp.today().normalize()

    week_start = (
        today
        - pd.Timedelta(days=today.weekday())
    )

    items = [
        ("Total Deliveries", len(df))
    ]

    if status_column:

        values = (
            df[status_column]
            .astype(str)
            .str.lower()
        )

        items.extend(
            [
                (
                    "Delivered",
                    int(values.str.contains("deliver").sum())
                ),

                (
                    "Pending",
                    int(values.str.contains("pending|track").sum())
                ),

                (
                    "Delayed",
                    int(values.str.contains("delay").sum())
                )
            ]
        )

    if remaining_column:

        items.append(
            (
                "Remaining Quantity",
                fmt(df[remaining_column].sum())
            )
        )

    if columns["expected_date"]:

        expected = df[
            columns["expected_date"]
        ]

        items.extend(
            [
                (
                    "Expected This Week",
                    int(
                        expected.between(
                            week_start,
                            week_start + pd.Timedelta(days=6)
                        ).sum()
                    )
                ),

                (
                    "Expected This Month",
                    int(
                        (
                            (expected.dt.month == today.month)
                            &
                            (expected.dt.year == today.year)
                        ).sum()
                    )
                )
            ]
        )

    kpi_row(items)

    # ------------------------------------------------------------------------
    # Table
    # ------------------------------------------------------------------------

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
        status_column
    ]

    show_columns = [
        column
        for column in show_columns
        if column
    ]

    st.dataframe(
        df[show_columns] if show_columns else df,
        width="stretch",
        hide_index=True
    )

    # ------------------------------------------------------------------------
    # Charts
    # ------------------------------------------------------------------------

    st.subheader("📊 Charts")

    c1, c2 = st.columns(2)

    if status_column:

        count = (
            df[status_column]
            .value_counts()
            .reset_index()
        )

        count.columns = [
            "Status",
            "Count"
        ]

        c1.plotly_chart(
            px.pie(
                count,
                names="Status",
                values="Count",
                hole=.45,
                title="Deliveries by Status",
                color="Status",
                color_discrete_map=STATUS_COLORS
            ),
            width="stretch"
        )

    if (
        columns["expected_date"]
        and df[columns["expected_date"]].notna().any()
    ):

        by_date = (
            df.groupby(
                df[columns["expected_date"]].dt.date
            )
            .size()
            .reset_index(name="Deliveries")
        )

        c2.plotly_chart(
            px.bar(
                by_date,
                x=columns["expected_date"],
                y="Deliveries",
                title="Expected Deliveries by Date",
                color_discrete_sequence=[PRIMARY]
            ),
            width="stretch"
        )

    c3, c4 = st.columns(2)

    if (
        columns["delivered_qty"]
        and remaining_column
    ):

        total = pd.DataFrame(
            {
                "Type": [
                    "Delivered",
                    "Remaining"
                ],

                "Qty": [
                    df[columns["delivered_qty"]].sum(),
                    df[remaining_column].sum()
                ]
            }
        )

        c3.plotly_chart(
            px.bar(
                total,
                x="Type",
                y="Qty",
                color="Type",
                title="Delivered vs Remaining Quantity",
                color_discrete_sequence=[
                    GREEN,
                    PRIMARY
                ]
            ),
            width="stretch"
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
                color_discrete_sequence=[PRIMARY]
            ),
            width="stretch"
        )


# ============================================================================
# 12) GENERIC DASHBOARD
# ============================================================================

def show_generic_dashboard(df, sheet_name):

    numeric_columns = [
        column
        for column in df.columns
        if pd.api.types.is_numeric_dtype(df[column])
    ]

    categorical_columns = [
        column
        for column in df.columns
        if df[column].dtype == object
        and 2 <= df[column].nunique() <= 20
    ]

    # ------------------------------------------------------------------------
    # Filters
    # ------------------------------------------------------------------------

    st.subheader("🔎 Filters")

    search = st.text_input(
        "Search",
        key=f"generic_search_{sheet_name}"
    )

    if search:

        mask = (
            df.astype(str)
            .apply(
                lambda column:
                column.str.contains(
                    search,
                    case=False,
                    na=False
                )
            )
            .any(axis=1)
        )

        df = df[mask]

    filter_columns = st.columns(3)

    for column_ui, column in zip(
        filter_columns,
        categorical_columns[:3]
    ):

        selected = column_ui.multiselect(
            column,
            sorted(
                df[column]
                .dropna()
                .astype(str)
                .unique()
            ),
            key=f"generic_{sheet_name}_{column}"
        )

        if selected:

            df = df[
                df[column]
                .astype(str)
                .isin(selected)
            ]

    # ------------------------------------------------------------------------
    # KPIs
    # ------------------------------------------------------------------------

    items = [
        ("Total Records", len(df)),
        ("Columns", len(df.columns))
    ]

    for column in numeric_columns[:3]:

        items.append(
            (
                f"Total {column}",
                fmt(df[column].sum())
            )
        )

    status_column = find_col(
        df,
        "status"
    )

    if status_column:

        values = (
            df[status_column]
            .astype(str)
            .str.lower()
        )

        items.extend(
            [
                (
                    "Completed",
                    int(
                        values.str.contains(
                            "complet|deliver|done"
                        ).sum()
                    )
                ),

                (
                    "Pending",
                    int(
                        values.str.contains(
                            "pend"
                        ).sum()
                    )
                ),

                (
                    "Delayed",
                    int(
                        values.str.contains(
                            "delay"
                        ).sum()
                    )
                )
            ]
        )

    kpi_row(items[:6])

    # ------------------------------------------------------------------------
    # Data
    # ------------------------------------------------------------------------

    st.subheader("📋 Data")

    st.dataframe(
        df,
        width="stretch",
        hide_index=True
    )

    # ------------------------------------------------------------------------
    # Charts
    # ------------------------------------------------------------------------

    if categorical_columns or numeric_columns:

        st.subheader("📊 Charts")

        charts = st.columns(2)

        if categorical_columns:

            column = categorical_columns[0]

            count = (
                df[column]
                .value_counts()
                .reset_index()
            )

            count.columns = [
                column,
                "Count"
            ]

            charts[0].plotly_chart(
                px.bar(
                    count,
                    x=column,
                    y="Count",
                    title=f"Records by {column}",
                    color_discrete_sequence=[PRIMARY]
                ),
                width="stretch"
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
                    title=f"{number} by {category}"
                ),
                width="stretch"
            )


def show_dashboard(df, sheet_name):

    st.title(
        f"{sheet_name} Dashboard"
    )

    if df.empty:

        st.warning(
            "This sheet is empty."
        )

        return

    is_delivery = (
        "deliver" in sheet_name.lower()
        or (
            find_col(df, "expected_date")
            and find_col(df, "required_date")
        )
    )

    if is_delivery:

        show_delivery_dashboard(
            df,
            sheet_name
        )

    else:

        show_generic_dashboard(
            df,
            sheet_name
        )


# ============================================================================
# 13) SHEET CLASSIFICATION
# ============================================================================

def classify_sheet(sheet_name, df=None):

    name = sheet_name.lower().strip()

    # Strong name-based detection
    if (
        "pipeline" in name
        and any(
            word in name
            for word in [
                "detail",
                "details",
                "cost",
                "costs",
                "line",
                "transaction"
            ]
        )
    ):

        return "details"

    if "pipeline" in name:

        return "pipelines"

    if "project" in name:

        return "projects"

    # Cost / transaction sheets
    if any(
        word in name
        for word in [
            "cost",
            "transaction",
            "expense"
        ]
    ):

        if df is not None:

            pipeline_id = pick(
                df,
                PIPELINE_COLS["id"]
            )

            pipeline_name = pick(
                df,
                PIPELINE_COLS["name"]
            )

            actual = pick(
                df,
                DETAIL_COLS["actual"]
            )

            if (
                (pipeline_id or pipeline_name)
                and actual
            ):

                return "details"

    # Column-based detection
    if df is not None and not df.empty:

        project_id = pick(
            df,
            PROJECT_COLS["id"]
        )

        project_name = pick(
            df,
            PROJECT_COLS["name"]
        )

        pipeline_id = pick(
            df,
            PIPELINE_COLS["id"]
        )

        pipeline_name = pick(
            df,
            PIPELINE_COLS["name"]
        )

        actual = pick(
            df,
            DETAIL_COLS["actual"]
        )

        planned = pick(
            df,
            DETAIL_COLS["planned"]
        )

        # Pipeline
        if (
            (pipeline_id or pipeline_name)
            and (
                actual
                or planned
                or pick(df, PIPELINE_COLS["progress"])
                or pick(df, PIPELINE_COLS["status"])
            )
        ):

            # If it has detail-level item/category/date
            if (
                pick(df, DETAIL_COLS["item"])
                or pick(df, DETAIL_COLS["category"])
                or pick(df, DETAIL_COLS["date"])
            ):

                return "details"

            return "pipelines"

        # Project
        if (
            project_id
            or project_name
        ):

            if not pipeline_id and not pipeline_name:

                return "projects"

    return "other"


# ============================================================================
# 14) PIPELINE STATUS
# ============================================================================

def pipeline_state(status, progress, end):

    status_text = str(
        status
    ).lower()

    if (
        pd.notna(progress)
        and float(progress) >= 100
    ):

        return "Completed"

    if any(
        word in status_text
        for word in [
            "complete",
            "completed",
            "done",
            "finish",
            "finished",
            "closed"
        ]
    ):

        return "Completed"

    if (
        "delay" in status_text
        or "late" in status_text
    ):

        return "Delayed"

    if (
        pd.notna(end)
        and end < pd.Timestamp.today().normalize()
    ):

        return "Delayed"

    if (
        pd.notna(progress)
        and float(progress) > 0
    ):

        return "Active"

    if any(
        word in status_text
        for word in [
            "active",
            "progress",
            "ongoing",
            "running"
        ]
    ):

        return "Active"

    return "Not Started"


# ============================================================================
# 15) BUILD PROJECT / PIPELINE MODEL
# ============================================================================

@st.cache_data(show_spinner=False)
def build_model(file_bytes: bytes) -> dict:

    sheet_names = get_sheet_names(
        file_bytes
    )

    classified = {}

    other_sheets = []

    # ------------------------------------------------------------------------
    # Classify every sheet
    # ------------------------------------------------------------------------

    for sheet_name in sheet_names:

        try:

            df = load_excel_data(
                file_bytes,
                sheet_name
            )

            role = classify_sheet(
                sheet_name,
                df
            )

        except Exception:

            role = "other"

        if role in [
            "projects",
            "pipelines",
            "details"
        ]:

            classified.setdefault(
                role,
                []
            ).append(sheet_name)

        else:

            other_sheets.append(
                sheet_name
            )

    warnings = []

    # ------------------------------------------------------------------------
    # Select primary sheets
    # ------------------------------------------------------------------------

    project_sheet = (
        classified.get("projects", [None])[0]
    )

    pipeline_sheet = (
        classified.get("pipelines", [None])[0]
    )

    detail_sheet = (
        classified.get("details", [None])[0]
    )

    # Additional project/pipeline/detail sheets
    extra_project_sheets = (
        classified.get("projects", [])[1:]
    )

    extra_pipeline_sheets = (
        classified.get("pipelines", [])[1:]
    )

    extra_detail_sheets = (
        classified.get("details", [])[1:]
    )

    # ------------------------------------------------------------------------
    # No project/pipeline model
    # ------------------------------------------------------------------------

    if (
        project_sheet is None
        and pipeline_sheet is None
    ):

        return {
            "ok": False,
            "role": {},
            "other": other_sheets,
            "warnings": [
                "No Projects or Pipelines structure was detected."
            ]
        }

    # ------------------------------------------------------------------------
    # Load Projects
    # ------------------------------------------------------------------------

    if project_sheet:

        P = load_excel_data(
            file_bytes,
            project_sheet
        ).reset_index(drop=True)

    else:

        P = pd.DataFrame()

    # ------------------------------------------------------------------------
    # Load Pipelines
    # ------------------------------------------------------------------------

    pipeline_frames = []

    if pipeline_sheet:

        pipeline_frames.append(
            load_excel_data(
                file_bytes,
                pipeline_sheet
            )
        )

    for sheet_name in extra_pipeline_sheets:

        pipeline_frames.append(
            load_excel_data(
                file_bytes,
                sheet_name
            )
        )

    if pipeline_frames:

        L = pd.concat(
            pipeline_frames,
            ignore_index=True,
            sort=False
        )

    else:

        L = pd.DataFrame()

    # ------------------------------------------------------------------------
    # Load Details
    # ------------------------------------------------------------------------

    detail_frames = []

    if detail_sheet:

        detail_frames.append(
            load_excel_data(
                file_bytes,
                detail_sheet
            )
        )

    for sheet_name in extra_detail_sheets:

        detail_frames.append(
            load_excel_data(
                file_bytes,
                sheet_name
            )
        )

    if detail_frames:

        D = pd.concat(
            detail_frames,
            ignore_index=True,
            sort=False
        )

    else:

        D = None

    # =========================================================================
    # PROJECTS
    # =========================================================================

    if not P.empty:

        project_columns = cols_of(
            P,
            PROJECT_COLS
        )

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

            P["_key"] = (
                P[project_key_column]
                .map(normalize_text)
            )

            P = P[
                P["_key"] != ""
            ].copy()

            # Project Name
            if project_columns["name"]:

                P["_name"] = (
                    P[project_columns["name"]]
                    .map(normalize_text)
                )

            else:

                P["_name"] = P["_key"]

            # Remove duplicate projects
            duplicate_projects = P[
                P["_key"].duplicated()
            ]

            if not duplicate_projects.empty:

                warnings.append(
                    "Duplicate Project IDs were found. "
                    "Only the first record was kept."
                )

                P = P.drop_duplicates(
                    "_key",
                    keep="first"
                )

            # Financial columns
            if project_columns["budget"]:

                P["_budget"] = to_num(
                    P[project_columns["budget"]]
                )

            else:

                P["_budget"] = float("nan")

            if project_columns["cost"]:

                P["_cost"] = to_num(
                    P[project_columns["cost"]]
                )

            else:

                P["_cost"] = float("nan")

            if project_columns["progress"]:

                P["_progress"] = to_pct(
                    P[project_columns["progress"]]
                )

            else:

                P["_progress"] = float("nan")

    else:

        P = pd.DataFrame(
            columns=[
                "_key",
                "_name",
                "_budget",
                "_cost",
                "_progress"
            ]
        )

    # Map Project Name -> Project Key
    name_to_key = {}

    if not P.empty:

        name_to_key = {
            normalize_key(name): key
            for name, key
            in zip(
                P["_name"],
                P["_key"]
            )
        }

    # =========================================================================
    # PIPELINES
    # =========================================================================

    if not L.empty:

        pipeline_columns = cols_of(
            L,
            PIPELINE_COLS
        )

        # ------------------------------------------------------------
        # Project relation
        # ------------------------------------------------------------

        project_id_column = (
            pipeline_columns["proj_id"]
        )

        project_name_column = (
            pipeline_columns["proj_name"]
        )

        if (
            project_id_column
            and not P.empty
        ):

            L["_project_key"] = (
                L[project_id_column]
                .map(normalize_text)
            )

        elif project_name_column:

            L["_project_key"] = (
                L[project_name_column]
                .map(
                    lambda x:
                    name_to_key.get(
                        normalize_key(x),
                        normalize_text(x)
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

        # ------------------------------------------------------------
        # Pipeline ID / Name
        # ------------------------------------------------------------

        pipeline_id_column = (
            pipeline_columns["id"]
        )

        pipeline_name_column = (
            pipeline_columns["name"]
        )

        if pipeline_id_column:

            L["_pid"] = (
                L[pipeline_id_column]
                .map(normalize_text)
            )

        elif pipeline_name_column:

            L["_pid"] = (
                L[pipeline_name_column]
                .map(normalize_text)
            )

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

            L["_name"] = (
                L[pipeline_name_column]
                .map(normalize_text)
            )

        else:

            L["_name"] = L["_pid"]

        # ------------------------------------------------------------
        # Internal unique pipeline key
        # Project + Pipeline
        # ------------------------------------------------------------

        L["_pk"] = (
            L["_project_key"].astype(str)
            + "||"
            + L["_pid"].astype(str)
        )

        # Duplicate pipelines
        duplicate_mask = L["_pk"].duplicated(
            keep=False
        )

        if duplicate_mask.any():

            L["_pk"] = (
                L["_pk"]
                + "#"
                + L.groupby("_pk").cumcount().astype(str)
            )

            warnings.append(
                "Duplicate Pipeline IDs were found "
                "inside the same project. "
                "Internal unique keys were generated."
            )

    else:

        L = pd.DataFrame(
            columns=[
                "_project_key",
                "_pid",
                "_name",
                "_pk"
            ]
        )

    # =========================================================================
    # AUTO CREATE PROJECTS FROM PIPELINES
    # =========================================================================

    if not L.empty:

        pipeline_project_keys = set(
            L["_project_key"]
            .dropna()
            .astype(str)
            .tolist()
        )

        known_project_keys = set(
            P["_key"]
            .astype(str)
            .tolist()
        )

        missing_projects = [
            key
            for key in pipeline_project_keys
            if key
            and key not in known_project_keys
        ]

        if missing_projects:

            warnings.append(
                f"{len(missing_projects)} project(s) "
                "were found in Pipelines but not in Projects. "
                "They were created automatically."
            )

            stub_projects = pd.DataFrame(
                {
                    "_key": missing_projects,
                    "_name": missing_projects,
                    "_budget": float("nan"),
                    "_cost": float("nan"),
                    "_progress": float("nan")
                }
            )

            P = pd.concat(
                [
                    P,
                    stub_projects
                ],
                ignore_index=True
            )

    # =========================================================================
    # DETAILS
    # =========================================================================

    dsum = None

    if D is not None and not D.empty and not L.empty:

        detail_columns = cols_of(
            D,
            DETAIL_COLS
        )

        # ------------------------------------------------------------
        # Pipeline ID / Name
        # ------------------------------------------------------------

        detail_pipeline_id = (
            detail_columns["id"]
        )

        detail_pipeline_name = (
            detail_columns["name"]
        )

        if detail_pipeline_id:

            D["_pid"] = (
                D[detail_pipeline_id]
                .map(normalize_text)
            )

        elif detail_pipeline_name:

            D["_pid"] = (
                D[detail_pipeline_name]
                .map(normalize_text)
            )

        else:

            warnings.append(
                "Pipeline Details sheet has no Pipeline ID or Pipeline Name."
            )

            D = None

        if D is not None:

            # --------------------------------------------------------
            # Project relation
            # --------------------------------------------------------

            detail_project_id = (
                detail_columns["proj_id"]
            )

            detail_project_name = (
                detail_columns["proj_name"]
            )

            if detail_project_id:

                D["_project_key"] = (
                    D[detail_project_id]
                    .map(normalize_text)
                )

            elif detail_project_name:

                D["_project_key"] = (
                    D[detail_project_name]
                    .map(
                        lambda x:
                        name_to_key.get(
                            normalize_key(x),
                            normalize_text(x)
                        )
                    )
                )

            else:

                D["_project_key"] = None

            # --------------------------------------------------------
            # Match details to pipeline
            # --------------------------------------------------------

            if D["_project_key"].notna().any():

                D["_pk"] = (
                    D["_project_key"]
                    .astype(str)
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

            valid_pipeline_keys = set(
                L["_pk"].tolist()
            )

            unmatched = (
                ~D["_pk"].isin(
                    valid_pipeline_keys
                )
            )

            unmatched_count = int(
                unmatched.sum()
            )

            if unmatched_count:

                warnings.append(
                    f"{unmatched_count} detail row(s) "
                    "could not be matched to a pipeline."
                )

            D = D[
                D["_pk"].isin(
                    valid_pipeline_keys
                )
            ].copy()

            # --------------------------------------------------------
            # Numeric fields
            # --------------------------------------------------------

            if detail_columns["planned"]:

                D["_planned"] = to_num(
                    D[detail_columns["planned"]]
                )

            else:

                D["_planned"] = float("nan")

            if detail_columns["actual"]:

                D["_actual"] = to_num(
                    D[detail_columns["actual"]]
                )

            else:

                D["_actual"] = float("nan")

            # --------------------------------------------------------
            # Category
            # --------------------------------------------------------

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

            # --------------------------------------------------------
            # Date
            # --------------------------------------------------------

            if detail_columns["date"]:

                D["_date"] = pd.to_datetime(
                    D[detail_columns["date"]],
                    errors="coerce"
                )

            else:

                D["_date"] = pd.NaT

            # --------------------------------------------------------
            # Detail aggregation
            # --------------------------------------------------------

            dsum = (
                D.groupby("_pk")[
                    ["_planned", "_actual"]
                ]
                .sum(min_count=1)
            )

    # =========================================================================
    # FINANCIAL MODEL
    # =========================================================================

    F = pd.DataFrame()

    if not L.empty:

        F = pd.DataFrame(
            {
                "_pk": L["_pk"],
                "_key": L["_project_key"],
                "Pipeline": L["_name"]
            }
        ).reset_index(drop=True)

        L = L.reset_index(
            drop=True
        )

        nan_series = pd.Series(
            float("nan"),
            index=L.index
        )

        # ------------------------------------------------------------
        # Budget
        # ------------------------------------------------------------

        if pipeline_columns["budget"]:

            F["Budget"] = to_num(
                L[pipeline_columns["budget"]]
            )

        else:

            F["Budget"] = nan_series.copy()

        # ------------------------------------------------------------
        # Planned
        # ------------------------------------------------------------

        if pipeline_columns["planned"]:

            planned = to_num(
                L[pipeline_columns["planned"]]
            )

        else:

            planned = nan_series.copy()

        if dsum is not None:

            detail_planned = (
                L["_pk"]
                .map(
                    dsum["_planned"]
                )
            )

            planned = (
                planned
                .fillna(detail_planned)
            )

        F["Planned"] = planned

        # ------------------------------------------------------------
        # Actual
        # ------------------------------------------------------------

        if pipeline_columns["actual"]:

            actual = to_num(
                L[pipeline_columns["actual"]]
            )

        else:

            actual = nan_series.copy()

        if dsum is not None:

            detail_actual = (
                L["_pk"]
                .map(
                    dsum["_actual"]
                )
            )

            actual = (
                actual
                .fillna(detail_actual)
            )

        F["Actual"] = actual

        # ------------------------------------------------------------
        # Remaining
        # Priority:
        # Excel Remaining
        # Budget - Actual
        # Planned - Actual
        # ------------------------------------------------------------

        if pipeline_columns["remaining"]:

            remaining = to_num(
                L[pipeline_columns["remaining"]]
            )

        else:

            remaining = pd.Series(
                float("nan"),
                index=L.index
            )

        budget_actual = (
            F["Budget"]
            - F["Actual"]
        )

        planned_actual = (
            F["Planned"]
            - F["Actual"]
        )

        remaining = (
            remaining
            .fillna(budget_actual)
            .fillna(planned_actual)
        )

        F["Remaining"] = remaining

        # ------------------------------------------------------------
        # Variance
        # Positive = under planned
        # Negative = over planned
        # ------------------------------------------------------------

        F["Variance"] = (
            F["Planned"]
            - F["Actual"]
        )

        # ------------------------------------------------------------
        # Progress
        # ------------------------------------------------------------

        if pipeline_columns["progress"]:

            F["Progress"] = to_pct(
                L[pipeline_columns["progress"]]
            )

        else:

            # Automatic progress from cost
            F["Progress"] = pd.Series(
                float("nan"),
                index=L.index
            )

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

        F["Progress"] = (
            F["Progress"]
            .clip(lower=0, upper=100)
        )

        # ------------------------------------------------------------
        # Status
        # ------------------------------------------------------------

        if pipeline_columns["end"]:

            end_dates = pd.to_datetime(
                L[pipeline_columns["end"]],
                errors="coerce"
            )

        else:

            end_dates = pd.Series(
                pd.NaT,
                index=L.index
            )

        if pipeline_columns["status"]:

            raw_status = L[
                pipeline_columns["status"]
            ]

        else:

            raw_status = pd.Series(
                "",
                index=L.index
            )

        F["State"] = [
            pipeline_state(
                status,
                progress,
                end_date
            )
            for status, progress, end_date
            in zip(
                raw_status,
                F["Progress"],
                end_dates
            )
        ]

    # =========================================================================
    # RETURN MODEL
    # =========================================================================

    # ============================================================================
    # NORMALIZE FINANCIAL MODEL COLUMNS
    # ============================================================================
    # Keep a stable internal schema even when optional financial columns are
    # missing from the uploaded Excel workbook. This prevents KeyError failures
    # such as KeyError: "Actual".
    required_financial_columns = {
        "Budget": "float",
        "Planned": "float",
        "Actual": "float",
        "Remaining": "float",
        "Variance": "float",
        "Progress": "float",
        "State": "object",
    }

    for _column, _dtype in required_financial_columns.items():
        if _column not in F.columns:
            F[_column] = pd.Series(index=F.index, dtype=_dtype)

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
        "D": D
    }


# ============================================================================
# 16) PROJECT SUMMARY
# ============================================================================

def summarize(
    F,
    budget=float("nan"),
    cost=float("nan")
):

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
            "Delayed": 0
        }

    # Budget
    total_budget = budget

    if pd.isna(total_budget):

        total_budget = F[
            "Budget"
        ].sum(
            min_count=1
        )

    # Planned
    total_planned = F[
        "Planned"
    ].sum(
        min_count=1
    )

    # Actual
    total_actual = F[
        "Actual"
    ].sum(
        min_count=1
    )

    if (
        pd.isna(total_actual)
        and pd.notna(cost)
    ):

        total_actual = cost

    # Remaining
    if (
        pd.notna(total_budget)
        and pd.notna(total_actual)
    ):

        remaining = (
            total_budget
            - total_actual
        )

    elif (
        pd.notna(total_planned)
        and pd.notna(total_actual)
    ):

        remaining = (
            total_planned
            - total_actual
        )

    else:

        remaining = float("nan")

    # Weighted progress
    weights = F[
        "Budget"
    ].where(
        F["Budget"] > 0,
        F["Planned"]
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

        overall_progress = (
            F["Progress"]
            .mean()
        )

    else:

        overall_progress = float("nan")

    return {
        "Total Budget": total_budget,
        "Total Planned Cost": total_planned,
        "Total Actual Cost": total_actual,
        "Remaining Cost": remaining,
        "Overall Progress": overall_progress,
        "Pipelines": len(F),
        "Completed": int(
            (F["State"] == "Completed").sum()
        ),
        "Active": int(
            (F["State"] == "Active").sum()
        ),
        "Delayed": int(
            (F["State"] == "Delayed").sum()
        )
    }


# ============================================================================
# 17) DISPLAY HELPERS
# ============================================================================

def format_value(value):

    if pd.isna(value):

        return "—"

    if isinstance(
        value,
        (
            pd.Timestamp,
            pd.DatetimeIndex
        )
    ):

        return (
            pd.Timestamp(value)
            .strftime("%Y-%m-%d")
        )

    if isinstance(
        value,
        (int, float)
    ):

        return fmt(value)

    return str(value)


def show_kpis(data):

    items = []

    for key, value in data.items():

        if pd.isna(value):

            continue

        if (
            "Progress" in key
            and isinstance(value, (int, float))
        ):

            display = f"{value:,.1f}%"

        elif isinstance(
            value,
            (int, float)
        ):

            display = fmt(value)

        else:

            display = str(value)

        items.append(
            (key, display)
        )

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
                format_value(value)
            )
        )

    if not items:
        return

    columns_ui = st.columns(3)

    for i, (key, value) in enumerate(items):

        columns_ui[
            i % 3
        ].markdown(
            f"**{key}**  \n{value}"
        )


def cost_chart(
    dataframe,
    x="Pipeline",
    title="Budget vs Planned vs Actual"
):

    if dataframe is None or dataframe.empty:
        return None

    available = [
        column
        for column in [
            "Budget",
            "Planned",
            "Actual"
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
        value_name="Cost"
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
            "Actual": PRIMARY
        }
    )


# ============================================================================
# 18) HOME PAGE
# ============================================================================

def page_home():

    if os.path.exists(
        LOGO_PATH
    ):

        st.image(
            LOGO_PATH,
            width=180
        )

    st.markdown(
        f"""
        <div class="hero">

            <span class="co">
                {COMPANY}
            </span>

            <h1>
                Project & Pipeline Management
            </h1>

            <p>
                Monitor projects, pipelines, costs,
                deliveries and progress.
            </p>

        </div>
        """,
        unsafe_allow_html=True
    )

    if not ss.file_bytes:

        st.info(
            "⬅ Upload an Excel (.xlsx) workbook "
            "from the sidebar to start."
        )

        return

    st.success(
        f"Loaded: {ss.file_name}"
    )

    if not M or not M.get("ok"):

        st.warning(
            "No Projects or Pipelines structure "
            "was detected."
        )

        names = (
            get_sheet_names(
                ss.file_bytes
            )
        )

        if names:

            st.subheader(
                "Available Sheets"
            )

            grid(
                [
                    (name, name)
                    for name in names
                ],
                open_sheet,
                "home_sheet",
                4
            )

        return

    # ------------------------------------------------------------------------
    # Workbook KPIs
    # ------------------------------------------------------------------------

    project_count = len(
        M["P"]
    )

    pipeline_count = len(
        M["F"]
    )

    total_budget = M["P"][
        "_budget"
    ].sum(
        min_count=1
    )

    total_actual = M["F"][
        "Actual"
    ].sum(
        min_count=1
    )

    delayed = int(
        (
            M["F"]["State"]
            == "Delayed"
        ).sum()
    )

    completed = int(
        (
            M["F"]["State"]
            == "Completed"
        ).sum()
    )

    kpi_row(
        [
            ("Projects", project_count),
            ("Pipelines", pipeline_count),
            (
                "Total Budget",
                fmt(total_budget)
            ),
            (
                "Total Actual Cost",
                fmt(total_actual)
            ),
            (
                "Delayed Pipelines",
                delayed
            ),
            (
                "Completed Pipelines",
                completed
            )
        ]
    )

    st.button(
        "📁 Open Projects",
        on_click=go,
        args=("projects",),
        width="stretch"
    )

    # ------------------------------------------------------------------------
    # Detected Sheets
    # ------------------------------------------------------------------------

    if M.get("role"):

        st.subheader(
            "Detected Project Structure"
        )

        for role, sheet_name in M[
            "role"
        ].items():

            st.write(
                f"**{role.title()}** → `{sheet_name}`"
            )

    # ------------------------------------------------------------------------
    # Warnings
    # ------------------------------------------------------------------------

    for warning in M.get(
        "warnings",
        []
    ):

        st.warning(
            "⚠ " + warning
        )

    # ------------------------------------------------------------------------
    # Other Sheets
    # ------------------------------------------------------------------------

    if M.get("other"):

        st.subheader(
            "Other Sheets"
        )

        grid(
            [
                (name, name)
                for name in M["other"]
            ],
            open_sheet,
            "other_sheet",
            4
        )


# ============================================================================
# 19) PROJECTS PAGE
# ============================================================================

def page_projects():

    st.title(
        "📁 Projects"
    )

    P = M["P"]

    if P.empty:

        st.warning(
            "No projects were found."
        )

        return

    st.caption(
        f"{len(P)} project(s) detected from Excel."
    )

    items = []

    for _, row in P.iterrows():

        name = (
            row["_name"]
            if row["_name"]
            else row["_key"]
        )

        items.append(
            (
                f"📂 {name}",
                row["_key"]
            )
        )

    grid(
        items,
        open_project,
        "project",
        3
    )


# ============================================================================
# 20) PROJECT PAGE
# ============================================================================

def page_project():

    project_key = ss.project

    P = M["P"]
    F = M["F"]

    project_rows = P[
        P["_key"] == project_key
    ]

    if project_rows.empty:

        st.error(
            "Project was not found."
        )

        return

    project_row = (
        project_rows.iloc[0]
    )

    pipeline_rows = F[
        F["_key"] == project_key
    ].copy()

    project_name = (
        project_row["_name"]
    )

    # ------------------------------------------------------------------------
    # Header
    # ------------------------------------------------------------------------

    st.markdown(
        f"""
        <div class="topbar">
            {COMPANY} | PROJECT
        </div>
        """,
        unsafe_allow_html=True
    )

    st.title(
        f"📂 {project_name}"
    )

    # ------------------------------------------------------------------------
    # Project information
    # ------------------------------------------------------------------------

    st.subheader(
        "Project Information"
    )

    info_grid(
        project_row,
        list(P.columns)
    )

    # ------------------------------------------------------------------------
    # No pipelines
    # ------------------------------------------------------------------------

    if pipeline_rows.empty:

        st.warning(
            "This project has no pipelines "
            "in the workbook."
        )

        return

    # ------------------------------------------------------------------------
    # KPIs
    # ------------------------------------------------------------------------

    summary = summarize(
        pipeline_rows,
        project_row["_budget"],
        project_row["_cost"]
    )

    show_kpis(
        summary
    )

    # ------------------------------------------------------------------------
    # Budget validation
    # ------------------------------------------------------------------------

    if (
        pd.notna(
            project_row["_budget"]
        )
        and pipeline_rows["Budget"]
        .sum(
            min_count=1
        ) > project_row["_budget"] * 1.0001
    ):

        st.warning(
            "⚠ The sum of pipeline budgets "
            "is higher than the project's Total Budget."
        )

    # ------------------------------------------------------------------------
    # Pipelines
    # ------------------------------------------------------------------------

    st.subheader(
        "🔧 Pipelines"
    )

    pipeline_buttons = []

    for _, row in pipeline_rows.iterrows():

        pipeline_buttons.append(
            (
                f"🔧 {row['Pipeline']}",
                row["_pk"]
            )
        )

    grid(
        pipeline_buttons,
        open_pipeline,
        "pipeline",
        3
    )

    # ------------------------------------------------------------------------
    # Charts
    # ------------------------------------------------------------------------

    st.subheader(
        "📊 Project Analytics"
    )

    c1, c2 = st.columns(2)

    chart = cost_chart(
        pipeline_rows
    )

    if chart:

        c1.plotly_chart(
            chart,
            width="stretch"
        )

    status_count = (
        pipeline_rows["State"]
        .value_counts()
        .reset_index()
    )

    status_count.columns = [
        "State",
        "Count"
    ]

    c2.plotly_chart(
        px.pie(
            status_count,
            names="State",
            values="Count",
            hole=.45,
            title="Pipelines by Status",
            color="State",
            color_discrete_map=STATUS_COLORS
        ),
        width="stretch"
    )

    # ------------------------------------------------------------------------
    # Progress
    # ------------------------------------------------------------------------

    if pipeline_rows[
        "Progress"
    ].notna().any():

        st.plotly_chart(
            px.bar(
                pipeline_rows,
                x="Progress",
                y="Pipeline",
                orientation="h",
                title="Progress % by Pipeline",
                range_x=[0, 100],
                color_discrete_sequence=[
                    PRIMARY
                ]
            ),
            width="stretch"
        )

    # ------------------------------------------------------------------------
    # Pipeline table
    # ------------------------------------------------------------------------

    st.subheader(
        "📋 Pipelines Table"
    )

    table = pipeline_rows.drop(
        columns=[
            "_pk",
            "_key"
        ],
        errors="ignore"
    )

    st.dataframe(
        table,
        width="stretch",
        hide_index=True
    )


# ============================================================================
# 21) PIPELINE PAGE
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

        st.error(
            "Pipeline was not found."
        )

        return

    pipeline = (
        pipeline_rows.iloc[0]
    )

    raw_rows = L[
        L["_pk"] == pipeline_key
    ]

    if raw_rows.empty:

        raw = pipeline

    else:

        raw = raw_rows.iloc[0]

    project_name = (
        M["P"]
        .set_index("_key")["_name"]
        .get(
            pipeline["_key"],
            pipeline["_key"]
        )
    )

    # ------------------------------------------------------------------------
    # Header
    # ------------------------------------------------------------------------

    st.markdown(
        f"""
        <div class="topbar">
            {COMPANY} |
            {project_name} |
            PIPELINE
        </div>
        """,
        unsafe_allow_html=True
    )

    st.button(
        f"⬅ Back to {project_name}",
        on_click=go,
        args=("project",),
        width="stretch"
    )

    st.title(
        f"🔧 {pipeline['Pipeline']}"
    )

    # ------------------------------------------------------------------------
    # Pipeline information
    # ------------------------------------------------------------------------

    st.subheader(
        "Pipeline Information"
    )

    info_grid(
        raw,
        list(raw.index)
    )

    # ------------------------------------------------------------------------
    # Financial KPIs
    # ------------------------------------------------------------------------

    show_kpis(
        {
            "Budget": pipeline["Budget"],
            "Planned Cost": pipeline["Planned"],
            "Actual Cost": pipeline.get("Actual", float("nan")),
            "Remaining Cost": pipeline.get("Remaining", float("nan")),
            "Cost Variance (Planned − Actual)": pipeline.get("Variance", float("nan")),
            "Progress %": pipeline.get("Progress", float("nan")),
            "Status": pipeline.get("State", "Unknown")
        }
    )

    # ------------------------------------------------------------------------
    # Progress bar
    # ------------------------------------------------------------------------

    progress = pipeline.get(
        "Progress",
        float("nan")
    )

    if pd.notna(progress):

        progress_value = min(
            max(
                float(progress),
                0
            ),
            100
        )

        st.progress(
            progress_value / 100
        )

        st.caption(
            f"Pipeline Progress: {progress_value:.1f}%"
        )

    # ------------------------------------------------------------------------
    # Cost warning
    # ------------------------------------------------------------------------

    if (
        pd.notna(
            pipeline.get("Variance", float("nan"))
        )
        and pipeline.get("Variance", float("nan")) < 0
    ):

        st.error(
            "⚠ Actual cost exceeds the planned cost."
        )

    elif (
        pd.notna(
            pipeline.get("Variance", float("nan"))
        )
        and pipeline.get("Variance", float("nan")) > 0
    ):

        st.success(
            "✓ Actual cost is currently below planned cost."
        )

    # ------------------------------------------------------------------------
    # Cost chart
    # ------------------------------------------------------------------------

    st.subheader(
        "📊 Pipeline Analytics"
    )

    c1, c2 = st.columns(2)

    one_pipeline = F[
        F["_pk"] == pipeline_key
    ]

    chart = cost_chart(
        one_pipeline,
        title="Budget vs Planned vs Actual"
    )

    if chart:

        c1.plotly_chart(
            chart,
            width="stretch"
        )

    # ------------------------------------------------------------------------
    # Details
    # ------------------------------------------------------------------------

    if (
        D is not None
        and not D.empty
    ):

        details = D[
            D["_pk"] == pipeline_key
        ].copy()

        if not details.empty:

            # ------------------------------------------------------------
            # Cost by Category
            # ------------------------------------------------------------

            category = (
                details
                .groupby("_cat")[
                    ["_planned", "_actual"]
                ]
                .sum(
                    min_count=1
                )
                .reset_index()
            )

            category = category.rename(
                columns={
                    "_cat": "Category",
                    "_planned": "Planned",
                    "_actual": "Actual"
                }
            )

            long = category.melt(
                id_vars="Category",
                var_name="Type",
                value_name="Cost"
            )

            c2.plotly_chart(
                px.bar(
                    long,
                    x="Category",
                    y="Cost",
                    color="Type",
                    barmode="group",
                    title="Cost by Category",
                    color_discrete_map={
                        "Planned": "#8A94A0",
                        "Actual": PRIMARY
                    }
                ),
                width="stretch"
            )

            # ------------------------------------------------------------
            # Cumulative Actual
            # ------------------------------------------------------------

            if (
                details["_date"].notna().any()
                and details["_actual"].notna().any()
            ):

                timeline = (
                    details
                    .dropna(
                        subset=["_date"]
                    )
                    .sort_values("_date")
                    .copy()
                )

                timeline[
                    "Cumulative Actual"
                ] = (
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
                        color_discrete_sequence=[
                            PRIMARY
                        ]
                    ),
                    width="stretch"
                )

            # ------------------------------------------------------------
            # Detail KPIs
            # ------------------------------------------------------------

            detail_planned = details[
                "_planned"
            ].sum(
                min_count=1
            )

            detail_actual = details[
                "_actual"
            ].sum(
                min_count=1
            )

            kpi_row(
                [
                    (
                        "Detail Records",
                        len(details)
                    ),

                    (
                        "Detail Planned",
                        fmt(detail_planned)
                    ),

                    (
                        "Detail Actual",
                        fmt(detail_actual)
                    )
                ]
            )

            # ------------------------------------------------------------
            # Detail table
            # ------------------------------------------------------------

            st.subheader(
                "📋 Pipeline Cost Details"
            )

            original_columns = [
                column
                for column in details.columns
                if not column.startswith("_")
            ]

            if original_columns:

                st.dataframe(
                    details[
                        original_columns
                    ],
                    width="stretch",
                    hide_index=True
                )

            else:

                st.dataframe(
                    details,
                    width="stretch",
                    hide_index=True
                )

        else:

            st.info(
                "No cost/detail records were found "
                "for this pipeline."
            )

    else:

        st.info(
            "Add a Pipeline Details / Costs sheet "
            "with Pipeline ID, Category, Planned, "
            "Actual and Date columns to display "
            "detailed pipeline costs."
        )


# ============================================================================
# 22) GENERIC SHEET PAGE
# ============================================================================

def page_sheet():

    st.button(
        "⬅ Home",
        on_click=go,
        args=("home",),
        width="stretch"
    )

    if not ss.sheet:

        go("home")

        return

    dataframe = load_excel_data(
        ss.file_bytes,
        ss.sheet
    )

    show_dashboard(
        dataframe,
        ss.sheet
    )


# ============================================================================
# 23) SIDEBAR
# ============================================================================

M = None

try:

    if ss.file_bytes:

        M = build_model(
            ss.file_bytes
        )

except Exception as error:

    st.error(
        f"Could not read the Excel file: {error}"
    )


with st.sidebar:

    # ------------------------------------------------------------------------
    # Logo
    # ------------------------------------------------------------------------

    if os.path.exists(
        LOGO_PATH
    ):

        st.image(
            LOGO_PATH,
            width=150
        )

    else:

        st.markdown(
            f"## {COMPANY}"
        )

    st.header(
        "🏗️ Navigation"
    )

    # ------------------------------------------------------------------------
    # Upload
    # ------------------------------------------------------------------------

    uploaded_file = st.file_uploader(
        "Upload Excel Workbook",
        type=["xlsx"]
    )

    if uploaded_file is not None:

        file_content = (
            uploaded_file.getvalue()
        )

        is_new_file = (
            uploaded_file.name
            != ss.file_name
            or
            len(file_content)
            != len(
                ss.file_bytes or b""
            )
        )

        if is_new_file:

            ss.file_bytes = file_content
            ss.file_name = uploaded_file.name

            ss.sheet = None
            ss.project = None
            ss.pipeline = None
            ss.page = "home"

            st.cache_data.clear()

            st.rerun()

    # ------------------------------------------------------------------------
    # Home
    # ------------------------------------------------------------------------

    st.button(
        "🏠 Home",
        on_click=go,
        args=("home",),
        width="stretch"
    )

    # ------------------------------------------------------------------------
    # Project navigation
    # ------------------------------------------------------------------------

    if M and M.get("ok"):

        st.button(
            "📁 Projects",
            on_click=go,
            args=("projects",),
            width="stretch"
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
                width="stretch"
            )

        if (
            ss.pipeline
            and
            ss.pipeline in set(
                M["F"]["_pk"]
            )
        ):

            st.button(
                "🔧 Current Pipeline",
                on_click=go,
                args=("pipeline",),
                width="stretch"
            )

    # ------------------------------------------------------------------------
    # Refresh
    # ------------------------------------------------------------------------

    if ss.file_bytes:

        if st.button(
            "🔄 Refresh Data",
            width="stretch"
        ):

            st.cache_data.clear()

            st.rerun()


# ============================================================================
# 24) MAIN APPLICATION ROUTER
# ============================================================================

try:

    # No Excel
    if not ss.file_bytes:

        page_home()

    # Home
    elif ss.page == "home":

        page_home()

    # Generic sheet
    elif (
        ss.page == "sheet"
        and ss.sheet
    ):

        page_sheet()

    # Project model
    elif M and M.get("ok"):

        if (
            ss.page == "projects"
        ):

            page_projects()

        elif (
            ss.page == "project"
            and ss.project
        ):

            page_project()

        elif (
            ss.page == "pipeline"
            and ss.pipeline
            and ss.pipeline in set(
                M["F"]["_pk"]
            )
        ):

            page_pipeline()

        else:

            page_projects()

    else:

        page_home()


except Exception as error:

    st.error(
        f"Something went wrong while building this page: {error}"
    )

    st.button(
        "⬅ Back to Home",
        on_click=go,
        args=("home",),
        width="stretch"
    )
