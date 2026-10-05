"""
Hassan Allam Construction
Construction Projects & Pipeline Dashboard
Streamlit + Google Sheets
"""

import os
import re
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st

import gspread
from google.oauth2.service_account import Credentials


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Hassan Allam Construction | Projects Dashboard",
    page_icon="🏗️",
    layout="wide",
)


# =========================================================
# CONFIG
# =========================================================

COMPANY = "Hassan Allam Construction"

PRIMARY = "#C8102E"
DARK = "#1F2933"
LIGHT = "#F4F5F7"
GRAY = "#6B7280"
GREEN = "#2E9E5B"
ORANGE = "#D97706"
BORDER = "#E1E4E8"

LOGO_PATH = "assets/logo.png"


# =========================================================
# VIEWER MODE
# =========================================================

VIEWER_MODE = st.query_params.get("view") == "dashboard"


# =========================================================
# STATUS COLORS
# =========================================================

STATUS_COLORS = {
    "Delivered": GREEN,
    "Completed": GREEN,
    "Active": PRIMARY,
    "On Track": GREEN,
    "Pending": ORANGE,
    "Pending / On Track": ORANGE,
    "Delayed": PRIMARY,
    "Not Started": GRAY,
    "Unknown": GRAY,
}


# =========================================================
# COLUMN MAPPING
# =========================================================

COLUMN_MAP = {
    "project": [
        "project",
        "project name",
        "project_name",
        "projectname",
        "name",
    ],
    "project_id": [
        "project id",
        "project_id",
        "project code",
        "project_code",
        "code",
        "id",
    ],
    "client": [
        "client",
        "customer",
        "owner",
        "client name",
    ],
    "location": [
        "location",
        "city",
        "area",
        "site",
    ],
    "manager": [
        "manager",
        "project manager",
        "project_manager",
        "pm",
    ],
    "start": [
        "start",
        "start date",
        "start_date",
        "planned start",
        "planned_start",
    ],
    "end": [
        "end",
        "end date",
        "end_date",
        "planned end",
        "planned_end",
        "completion date",
    ],
    "status": [
        "status",
        "project status",
        "project_status",
        "state",
    ],
    "budget": [
        "budget",
        "contract value",
        "contract_value",
        "value",
        "budget amount",
    ],
    "planned": [
        "planned",
        "planned cost",
        "planned_cost",
        "planned amount",
    ],
    "actual": [
        "actual",
        "actual cost",
        "actual_cost",
        "actual amount",
        "spent",
    ],
    "progress": [
        "progress",
        "progress %",
        "progress_pct",
        "completion",
        "completion %",
        "completion percentage",
    ],
    "pipeline": [
        "pipeline",
        "pipeline name",
        "pipeline_name",
        "opportunity",
        "opportunity name",
    ],
    "stage": [
        "stage",
        "pipeline stage",
        "pipeline_stage",
        "phase",
    ],
    "probability": [
        "probability",
        "probability %",
        "probability_pct",
        "win probability",
    ],
    "expected_value": [
        "expected value",
        "expected_value",
        "weighted value",
    ],
}


PROJECT_COLS = [
    "Project",
    "Project ID",
    "Client",
    "Location",
    "Manager",
    "Start Date",
    "End Date",
    "Status",
    "Budget",
    "Planned",
    "Actual",
    "Progress",
]


PIPELINE_COLS = [
    "Pipeline",
    "Project",
    "Client",
    "Stage",
    "Probability",
    "Expected Value",
    "Status",
]


DETAIL_COLS = [
    "Project",
    "Pipeline",
    "Category",
    "Item",
    "Description",
    "Quantity",
    "Unit Cost",
    "Total Cost",
    "Status",
]


# =========================================================
# CSS
# =========================================================

BASE_CSS = f"""
<style>

html, body, [class*="css"] {{
    font-family: Arial, sans-serif;
}}

.stApp {{
    background-color: #F8F9FA;
}}

section[data-testid="stSidebar"] {{
    background-color: white;
    border-right: 1px solid {BORDER};
}}

section[data-testid="stSidebar"] > div {{
    padding-top: 1rem;
}}

h1, h2, h3 {{
    color: {DARK};
}}

.dashboard-title {{
    font-size: 30px;
    font-weight: 700;
    color: {DARK};
    margin-bottom: 5px;
}}

.dashboard-subtitle {{
    color: {GRAY};
    font-size: 15px;
    margin-bottom: 25px;
}}

.hero {{
    background: linear-gradient(
        135deg,
        {DARK} 0%,
        #374151 100%
    );
    padding: 30px;
    border-radius: 16px;
    color: white;
    margin-bottom: 25px;
}}

.hero h1 {{
    color: white;
    margin-bottom: 5px;
}}

.hero p {{
    color: #E5E7EB;
}}

.kpi-card {{
    background: white;
    border: 1px solid {BORDER};
    border-radius: 14px;
    padding: 18px;
    min-height: 120px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
}}

.kpi-title {{
    color: {GRAY};
    font-size: 13px;
    font-weight: 600;
    margin-bottom: 8px;
}}

.kpi-value {{
    color: {DARK};
    font-size: 28px;
    font-weight: 700;
}}

.kpi-sub {{
    color: {GRAY};
    font-size: 12px;
    margin-top: 5px;
}}

.section-title {{
    font-size: 21px;
    font-weight: 700;
    color: {DARK};
    margin-top: 20px;
    margin-bottom: 12px;
}}

.badge {{
    display: inline-block;
    padding: 5px 10px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 600;
}}

.info-box {{
    background: white;
    border: 1px solid {BORDER};
    border-radius: 12px;
    padding: 18px;
    margin-bottom: 15px;
}}

[data-testid="stSidebar"] .stButton > button {{
    text-align: left;
    width: 100%;
    border: none;
    background: transparent;
}}

[data-testid="stSidebar"] .stButton > button:hover {{
    background-color: #F3F4F6;
}}

div.stButton > button[kind="primary"] {{
    background-color: {PRIMARY};
    border-color: {PRIMARY};
}}

</style>
"""

st.markdown(BASE_CSS, unsafe_allow_html=True)


# =========================================================
# VIEWER CSS
# IMPORTANT:
# This is a normal string, NOT an f-string.
# Therefore CSS {{ }} does not create Python syntax errors.
# =========================================================

if VIEWER_MODE:

    VIEWER_CSS = """
    <style>

    section[data-testid="stSidebar"] {
        display: none !important;
    }

    header[data-testid="stHeader"],
    [data-testid="stToolbar"],
    [data-testid="stDecoration"],
    #MainMenu,
    .stAppDeployButton,
    footer,
    [data-testid="stStatusWidget"] {
        display: none !important;
        visibility: hidden !important;
    }

    .block-container {
        padding-top: 1.5rem !important;
        max-width: 1400px !important;
    }

    </style>
    """

    st.markdown(
        VIEWER_CSS,
        unsafe_allow_html=True,
    )


# =========================================================
# SESSION STATE
# =========================================================

DEFAULT_STATE = {
    "page": "home",
    "sheet": None,
    "project": None,
    "pipeline": None,
}

for key, value in DEFAULT_STATE.items():

    if key not in st.session_state:
        st.session_state[key] = value


# =========================================================
# GOOGLE SHEETS CONNECTION
# =========================================================

@st.cache_resource
def get_google_client():

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


# =========================================================
# GET SHEET NAMES
# =========================================================

@st.cache_data(
    ttl=60,
    show_spinner=False,
)
def get_sheet_names():

    workbook = get_google_workbook()

    return [
        worksheet.title
        for worksheet in workbook.worksheets()
    ]


# =========================================================
# LOAD GOOGLE SHEET
# =========================================================

@st.cache_data(
    ttl=60,
    show_spinner=False,
)
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
        str(column).strip()
        for column in df.columns
    ]

    return df


# =========================================================
# HELPERS
# =========================================================

def normalize_text(value):

    if pd.isna(value):
        return ""

    value = str(value)

    value = value.strip().lower()

    value = re.sub(
        r"[^a-z0-9]+",
        " ",
        value,
    )

    return " ".join(value.split())


def normalize_key(value):

    return normalize_text(value)


def pick(df, key, default=None):

    if df is None or df.empty:
        return default

    mapping = COLUMN_MAP.get(key, [])

    normalized_columns = {
        normalize_text(column): column
        for column in df.columns
    }

    for candidate in mapping:

        normalized_candidate = normalize_text(
            candidate
        )

        if normalized_candidate in normalized_columns:

            return normalized_columns[
                normalized_candidate
            ]

    return default


def cols_of(df, names):

    if df is None or df.empty:
        return []

    result = []

    for name in names:

        column = pick(df, name)

        if column:
            result.append(column)

    return result


def find_col(df, candidates):

    if df is None or df.empty:
        return None

    normalized_columns = {
        normalize_text(column): column
        for column in df.columns
    }

    for candidate in candidates:

        key = normalize_text(candidate)

        if key in normalized_columns:
            return normalized_columns[key]

    return None


def to_num(value):

    if pd.isna(value):
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()

    text = (
        text
        .replace(",", "")
        .replace("$", "")
        .replace("EGP", "")
        .replace("%", "")
    )

    try:
        return float(text)

    except Exception:
        return 0.0


def fmt(value):

    try:
        return f"{float(value):,.0f}"

    except Exception:
        return "0"


def to_pct(value):

    number = to_num(value)

    if number <= 1:
        number *= 100

    return max(
        0,
        min(100, number),
    )


def safe_sum(series):

    if series is None:
        return 0

    return pd.to_numeric(
        series,
        errors="coerce",
    ).fillna(0).sum()


# =========================================================
# KPI
# =========================================================

def kpi(
    title,
    value,
    subtitle="",
):

    st.markdown(
        f"""
        <div class="kpi-card">

            <div class="kpi-title">
                {title}
            </div>

            <div class="kpi-value">
                {value}
            </div>

            <div class="kpi-sub">
                {subtitle}
            </div>

        </div>
        """,
        unsafe_allow_html=True,
    )


def kpi_row(items):

    columns = st.columns(
        len(items)
    )

    for column, item in zip(
        columns,
        items,
    ):

        with column:

            kpi(
                item[0],
                item[1],
                item[2] if len(item) > 2 else "",
            )


# =========================================================
# NAVIGATION
# =========================================================

def go(page):

    st.session_state.page = page

    st.session_state.project = None

    st.session_state.pipeline = None

    st.rerun()


def open_project(project):

    st.session_state.page = "project"

    st.session_state.project = project

    st.rerun()


def open_pipeline(pipeline):

    st.session_state.page = "pipeline"

    st.session_state.pipeline = pipeline

    st.rerun()


def open_sheet(sheet):

    st.session_state.page = "sheet"

    st.session_state.sheet = sheet

    st.rerun()


# =========================================================
# DATA MODEL
# =========================================================

def classify_sheet(sheet_name):

    name = normalize_text(sheet_name)

    if any(
        keyword in name
        for keyword in [
            "project",
            "projects",
        ]
    ):
        return "projects"

    if any(
        keyword in name
        for keyword in [
            "pipeline",
            "pipelines",
            "opportunity",
            "opportunities",
        ]
    ):
        return "pipelines"

    if any(
        keyword in name
        for keyword in [
            "detail",
            "details",
            "material",
            "materials",
            "formwork",
        ]
    ):
        return "details"

    return "other"


def pipeline_state(row):

    status_column = pick(
        row.to_frame().T,
        "status",
    )

    if status_column:

        status = str(
            row.get(status_column, "")
        ).strip()

        if status:
            return status

    probability_column = pick(
        row.to_frame().T,
        "probability",
    )

    if probability_column:

        probability = to_pct(
            row.get(
                probability_column,
                0,
            )
        )

        if probability >= 70:
            return "High Probability"

        if probability >= 40:
            return "Medium Probability"

        return "Low Probability"

    return "Unknown"


# =========================================================
# BUILD MODEL
# =========================================================

@st.cache_data(
    ttl=60,
    show_spinner=False,
)
def build_model():

    sheet_names = get_sheet_names()

    model = {
        "ok": True,
        "role": {},
        "other": [],
        "warnings": [],
        "P": pd.DataFrame(),
        "L": pd.DataFrame(),
        "F": pd.DataFrame(),
        "D": pd.DataFrame(),
    }

    if not sheet_names:

        model["ok"] = False

        model["warnings"].append(
            "No sheets found in Google Sheets."
        )

        return model

    project_sheets = []
    pipeline_sheets = []
    detail_sheets = []

    for sheet in sheet_names:

        role = classify_sheet(sheet)

        model["role"][sheet] = role

        if role == "projects":
            project_sheets.append(sheet)

        elif role == "pipelines":
            pipeline_sheets.append(sheet)

        elif role == "details":
            detail_sheets.append(sheet)

        else:
            model["other"].append(sheet)

    # -----------------------------------------------------
    # PROJECT DATA
    # -----------------------------------------------------

    project_frames = []

    for sheet in project_sheets:

        try:

            df = load_google_sheet(sheet)

            if not df.empty:

                df = df.copy()

                df["_source_sheet"] = sheet

                project_frames.append(df)

        except Exception as error:

            model["warnings"].append(
                f"{sheet}: {error}"
            )

    if project_frames:

        P = pd.concat(
            project_frames,
            ignore_index=True,
        )

    else:

        P = pd.DataFrame()

    model["P"] = P

    # -----------------------------------------------------
    # PIPELINE DATA
    # -----------------------------------------------------

    pipeline_frames = []

    for sheet in pipeline_sheets:

        try:

            df = load_google_sheet(sheet)

            if not df.empty:

                df = df.copy()

                df["_source_sheet"] = sheet

                pipeline_frames.append(df)

        except Exception as error:

            model["warnings"].append(
                f"{sheet}: {error}"
            )

    if pipeline_frames:

        L = pd.concat(
            pipeline_frames,
            ignore_index=True,
        )

    else:

        L = pd.DataFrame()

    model["L"] = L

    # -----------------------------------------------------
    # DETAILS
    # -----------------------------------------------------

    detail_frames = []

    for sheet in detail_sheets:

        try:

            df = load_google_sheet(sheet)

            if not df.empty:

                df = df.copy()

                df["_source_sheet"] = sheet

                detail_frames.append(df)

        except Exception as error:

            model["warnings"].append(
                f"{sheet}: {error}"
            )

    if detail_frames:

        D = pd.concat(
            detail_frames,
            ignore_index=True,
        )

    else:

        D = pd.DataFrame()

    model["D"] = D

    # -----------------------------------------------------
    # FINANCIAL MODEL
    # -----------------------------------------------------

    if not P.empty:

        F = P.copy()

        project_column = pick(
            F,
            "project",
        )

        budget_column = pick(
            F,
            "budget",
        )

        planned_column = pick(
            F,
            "planned",
        )

        actual_column = pick(
            F,
            "actual",
        )

        progress_column = pick(
            F,
            "progress",
        )

        status_column = pick(
            F,
            "status",
        )

        if project_column:

            F["Project"] = F[
                project_column
            ].astype(str)

        else:

            F["Project"] = [
                f"Project {i + 1}"
                for i in range(len(F))
            ]

        if budget_column:

            F["Budget"] = F[
                budget_column
            ].apply(to_num)

        else:

            F["Budget"] = 0

        if planned_column:

            F["Planned"] = F[
                planned_column
            ].apply(to_num)

        else:

            F["Planned"] = 0

        if actual_column:

            F["Actual"] = F[
                actual_column
            ].apply(to_num)

        else:

            F["Actual"] = 0

        F["Remaining"] = (
            F["Budget"] - F["Actual"]
        )

        F["Variance"] = (
            F["Planned"] - F["Actual"]
        )

        if progress_column:

            F["Progress"] = F[
                progress_column
            ].apply(to_pct)

        else:

            F["Progress"] = 0

        if status_column:

            F["State"] = (
                F[status_column]
                .fillna("")
                .astype(str)
                .str.strip()
            )

        else:

            F["State"] = "Unknown"

        model["F"] = F

    return model


# =========================================================
# DELIVERY STATUS
# =========================================================

def compute_delivery_status(df):

    if df.empty:

        return pd.Series(
            dtype=str
        )

    status_column = pick(
        df,
        "status",
    )

    if status_column:

        return (
            df[status_column]
            .fillna("Unknown")
            .astype(str)
            .str.strip()
        )

    return pd.Series(
        ["Unknown"] * len(df),
        index=df.index,
    )


# =========================================================
# FORMATTING
# =========================================================

def format_value(value):

    if pd.isna(value):
        return ""

    if isinstance(
        value,
        (int, float),
    ):

        return f"{value:,.2f}"

    return str(value)


# =========================================================
# COST CHART
# =========================================================

def cost_chart(df):

    if df.empty:
        return

    budget = safe_sum(
        df.get(
            "Budget",
            pd.Series(dtype=float),
        )
    )

    planned = safe_sum(
        df.get(
            "Planned",
            pd.Series(dtype=float),
        )
    )

    actual = safe_sum(
        df.get(
            "Actual",
            pd.Series(dtype=float),
        )
    )

    chart_df = pd.DataFrame(
        {
            "Metric": [
                "Budget",
                "Planned",
                "Actual",
            ],
            "Amount": [
                budget,
                planned,
                actual,
            ],
        }
    )

    fig = px.bar(
        chart_df,
        x="Metric",
        y="Amount",
        text_auto=True,
        title="Financial Overview",
    )

    fig.update_layout(
        height=400,
        showlegend=False,
    )

    st.plotly_chart(
        fig,
        width="stretch",
    )


# =========================================================
# PROJECT DASHBOARD
# =========================================================

def show_delivery_dashboard(model):

    F = model["F"]

    if F.empty:

        st.info(
            "No project data available."
        )

        return

    budget = safe_sum(
        F["Budget"]
    )

    planned = safe_sum(
        F["Planned"]
    )

    actual = safe_sum(
        F["Actual"]
    )

    remaining = safe_sum(
        F["Remaining"]
    )

    avg_progress = (
        F["Progress"].mean()
        if "Progress" in F
        else 0
    )

    kpi_row(
        [
            (
                "Projects",
                f"{len(F):,}",
                "Total projects",
            ),
            (
                "Budget",
                fmt(budget),
                "Total budget",
            ),
            (
                "Actual Cost",
                fmt(actual),
                "Actual spending",
            ),
            (
                "Remaining",
                fmt(remaining),
                "Budget remaining",
            ),
            (
                "Avg. Progress",
                f"{avg_progress:.1f}%",
                "Average completion",
            ),
        ]
    )

    st.markdown(
        '<div class="section-title">Financial Performance</div>',
        unsafe_allow_html=True,
    )

    cost_chart(F)

    st.markdown(
        '<div class="section-title">Project Status</div>',
        unsafe_allow_html=True,
    )

    state_counts = (
        F["State"]
        .fillna("Unknown")
        .replace("", "Unknown")
        .value_counts()
        .reset_index()
    )

    state_counts.columns = [
        "Status",
        "Projects",
    ]

    if not state_counts.empty:

        fig = px.bar(
            state_counts,
            x="Status",
            y="Projects",
            text_auto=True,
            title="Projects by Status",
        )

        fig.update_layout(
            height=400,
        )

        st.plotly_chart(
            fig,
            width="stretch",
        )

    st.markdown(
        '<div class="section-title">Projects</div>',
        unsafe_allow_html=True,
    )

    display_columns = [
        column
        for column in [
            "Project",
            "Budget",
            "Planned",
            "Actual",
            "Remaining",
            "Progress",
            "State",
        ]
        if column in F.columns
    ]

    st.dataframe(
        F[display_columns],
        width="stretch",
        hide_index=True,
    )


# =========================================================
# GENERIC DASHBOARD
# =========================================================

def show_generic_dashboard(
    sheet_name,
    df,
):

    if df.empty:

        st.warning(
            f"Sheet '{sheet_name}' is empty."
        )

        return

    st.markdown(
        f"""
        <div class="dashboard-title">
            {sheet_name}
        </div>

        <div class="dashboard-subtitle">
            Data from Google Sheets
        </div>
        """,
        unsafe_allow_html=True,
    )

    numeric_columns = df.select_dtypes(
        include="number"
    ).columns.tolist()

    kpis = [
        (
            "Rows",
            f"{len(df):,}",
            "Records",
        ),
        (
            "Columns",
            f"{len(df.columns):,}",
            "Fields",
        ),
    ]

    if numeric_columns:

        total = safe_sum(
            df[numeric_columns[0]]
        )

        kpis.append(
            (
                numeric_columns[0],
                fmt(total),
                "Total",
            )
        )

    kpi_row(kpis)

    st.markdown(
        '<div class="section-title">Data</div>',
        unsafe_allow_html=True,
    )

    st.dataframe(
        df,
        width="stretch",
        hide_index=True,
    )

    # Numeric charts
    if numeric_columns:

        chart_column = st.selectbox(
            "Select numeric column",
            numeric_columns,
            key=f"chart_{sheet_name}",
        )

        chart_df = df[
            [chart_column]
        ].copy()

        chart_df[chart_column] = pd.to_numeric(
            chart_df[chart_column],
            errors="coerce",
        )

        chart_df = chart_df.dropna()

        if not chart_df.empty:

            fig = px.bar(
                chart_df,
                y=chart_column,
                title=f"{chart_column} Overview",
            )

            st.plotly_chart(
                fig,
                width="stretch",
            )


# =========================================================
# FORM WORK DASHBOARD
# =========================================================

def is_formwork_sheet(sheet_name):

    name = normalize_text(
        sheet_name
    )

    return "formwork" in name


def show_formwork_dashboard(
    sheet_name,
    df,
):

    if df.empty:

        st.warning(
            "Formwork sheet is empty."
        )

        return

    st.markdown(
        f"""
        <div class="hero">

            <h1>
                🏗️ Formwork Dashboard
            </h1>

            <p>
                {sheet_name}
            </p>

        </div>
        """,
        unsafe_allow_html=True,
    )

    numeric_columns = df.select_dtypes(
        include="number"
    ).columns.tolist()

    kpi_items = [
        (
            "Records",
            f"{len(df):,}",
            "Total records",
        )
    ]

    for column in numeric_columns[:4]:

        kpi_items.append(
            (
                column,
                fmt(
                    safe_sum(
                        df[column]
                    )
                ),
                "Total",
            )
        )

    kpi_row(kpi_items)

    st.markdown(
        '<div class="section-title">Formwork Data</div>',
        unsafe_allow_html=True,
    )

    st.dataframe(
        df,
        width="stretch",
        hide_index=True,
    )


# =========================================================
# PROJECT SUMMARY
# =========================================================

def summarize(model):

    F = model["F"]

    if F.empty:

        return {
            "projects": 0,
            "budget": 0,
            "planned": 0,
            "actual": 0,
            "remaining": 0,
            "progress": 0,
        }

    return {
        "projects": len(F),
        "budget": safe_sum(F["Budget"]),
        "planned": safe_sum(F["Planned"]),
        "actual": safe_sum(F["Actual"]),
        "remaining": safe_sum(F["Remaining"]),
        "progress": (
            F["Progress"].mean()
            if "Progress" in F
            else 0
        ),
    }


# =========================================================
# HOME PAGE
# =========================================================

def page_home(model):

    summary = summarize(model)

    st.markdown(
        f"""
        <div class="hero">

            <h1>
                {COMPANY}
            </h1>

            <p>
                Construction Projects & Pipeline Dashboard
            </p>

        </div>
        """,
        unsafe_allow_html=True,
    )

    kpi_row(
        [
            (
                "Projects",
                f"{summary['projects']:,}",
                "Total projects",
            ),
            (
                "Budget",
                fmt(summary["budget"]),
                "Total project budget",
            ),
            (
                "Actual Cost",
                fmt(summary["actual"]),
                "Actual cost",
            ),
            (
                "Remaining",
                fmt(summary["remaining"]),
                "Remaining budget",
            ),
            (
                "Progress",
                f"{summary['progress']:.1f}%",
                "Average project progress",
            ),
        ]
    )

    st.markdown(
        '<div class="section-title">Executive Overview</div>',
        unsafe_allow_html=True,
    )

    if not model["F"].empty:

        cost_chart(
            model["F"]
        )

    # Pipeline
    L = model["L"]

    if not L.empty:

        st.markdown(
            '<div class="section-title">Pipeline</div>',
            unsafe_allow_html=True,
        )

        pipeline_display = L.copy()

        probability_column = pick(
            pipeline_display,
            "probability",
        )

        if probability_column:

            pipeline_display[
                "Probability"
            ] = pipeline_display[
                probability_column
            ].apply(to_pct)

        expected_column = pick(
            pipeline_display,
            "expected_value",
        )

        if expected_column:

            pipeline_display[
                "Expected Value"
            ] = pipeline_display[
                expected_column
            ].apply(to_num)

        columns_to_show = []

        for logical_name in [
            "pipeline",
            "project",
            "client",
            "stage",
            "probability",
            "expected_value",
            "status",
        ]:

            column = pick(
                pipeline_display,
                logical_name,
            )

            if column:
                columns_to_show.append(
                    column
                )

        if columns_to_show:

            st.dataframe(
                pipeline_display[
                    columns_to_show
                ],
                width="stretch",
                hide_index=True,
            )

    else:

        st.info(
            "No pipeline sheet was detected."
        )


# =========================================================
# PROJECTS PAGE
# =========================================================

def page_projects(model):

    st.markdown(
        """
        <div class="dashboard-title">
            Projects
        </div>

        <div class="dashboard-subtitle">
            All construction projects
        </div>
        """,
        unsafe_allow_html=True,
    )

    F = model["F"]

    if F.empty:

        st.warning(
            "No project data available."
        )

        return

    search = st.text_input(
        "🔎 Search project",
        "",
    )

    filtered = F.copy()

    if search:

        mask = (
            filtered["Project"]
            .astype(str)
            .str.contains(
                search,
                case=False,
                na=False,
            )
        )

        filtered = filtered[
            mask
        ]

    st.dataframe(
        filtered,
        width="stretch",
        hide_index=True,
    )

    st.markdown(
        '<div class="section-title">Open Project</div>',
        unsafe_allow_html=True,
    )

    projects = (
        filtered["Project"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    if projects:

        selected = st.selectbox(
            "Project",
            projects,
        )

        if st.button(
            "Open Project",
            type="primary",
        ):

            open_project(
                selected
            )


# =========================================================
# PROJECT DETAILS
# =========================================================

def page_project(model):

    project = st.session_state.project

    if not project:

        go("projects")

        return

    F = model["F"]

    project_rows = F[
        F["Project"]
        .astype(str)
        .str.strip()
        == str(project).strip()
    ]

    st.markdown(
        f"""
        <div class="dashboard-title">
            {project}
        </div>

        <div class="dashboard-subtitle">
            Project details
        </div>
        """,
        unsafe_allow_html=True,
    )

    if project_rows.empty:

        st.warning(
            "Project not found."
        )

        return

    row = project_rows.iloc[0]

    kpi_row(
        [
            (
                "Budget",
                fmt(row.get("Budget", 0)),
                "Project budget",
            ),
            (
                "Planned",
                fmt(row.get("Planned", 0)),
                "Planned cost",
            ),
            (
                "Actual",
                fmt(row.get("Actual", 0)),
                "Actual cost",
            ),
            (
                "Remaining",
                fmt(row.get("Remaining", 0)),
                "Remaining",
            ),
            (
                "Progress",
                f"{to_pct(row.get('Progress', 0)):.1f}%",
                "Completion",
            ),
        ]
    )

    st.markdown(
        '<div class="section-title">Project Information</div>',
        unsafe_allow_html=True,
    )

    info = {}

    for key, logical_name in [
        ("Client", "client"),
        ("Location", "location"),
        ("Manager", "manager"),
        ("Status", "status"),
        ("Start Date", "start"),
        ("End Date", "end"),
    ]:

        column = pick(
            model["P"],
            logical_name,
        )

        if column:

            matches = model["P"][
                model["P"][column]
                .astype(str)
                .str.strip()
                == str(project).strip()
            ]

            if not matches.empty:

                info[key] = matches.iloc[
                    0
                ].get(column, "")

    if info:

        info_df = pd.DataFrame(
            list(info.items()),
            columns=[
                "Field",
                "Value",
            ],
        )

        st.dataframe(
            info_df,
            width="stretch",
            hide_index=True,
        )

    # Details
    D = model["D"]

    if not D.empty:

        project_column = pick(
            D,
            "project",
        )

        if project_column:

            details = D[
                D[project_column]
                .astype(str)
                .str.strip()
                == str(project).strip()
            ]

            if not details.empty:

                st.markdown(
                    '<div class="section-title">Project Details</div>',
                    unsafe_allow_html=True,
                )

                st.dataframe(
                    details,
                    width="stretch",
                    hide_index=True,
                )


# =========================================================
# PIPELINE PAGE
# =========================================================

def page_pipeline(model):

    pipeline = st.session_state.pipeline

    L = model["L"]

    if L.empty:

        st.info(
            "No pipeline data available."
        )

        return

    pipeline_column = pick(
        L,
        "pipeline",
    )

    if not pipeline_column:

        st.dataframe(
            L,
            width="stretch",
            hide_index=True,
        )

        return

    rows = L[
        L[pipeline_column]
        .astype(str)
        .str.strip()
        == str(pipeline).strip()
    ]

    st.markdown(
        f"""
        <div class="dashboard-title">
            {pipeline}
        </div>

        <div class="dashboard-subtitle">
            Pipeline opportunity
        </div>
        """,
        unsafe_allow_html=True,
    )

    if rows.empty:

        st.warning(
            "Pipeline not found."
        )

        return

    row = rows.iloc[0]

    probability_column = pick(
        L,
        "probability",
    )

    expected_column = pick(
        L,
        "expected_value",
    )

    probability = (
        to_pct(
            row.get(
                probability_column,
                0,
            )
        )
        if probability_column
        else 0
    )

    expected_value = (
        to_num(
            row.get(
                expected_column,
                0,
            )
        )
        if expected_column
        else 0
    )

    kpi_row(
        [
            (
                "Probability",
                f"{probability:.1f}%",
                "Win probability",
            ),
            (
                "Expected Value",
                fmt(expected_value),
                "Weighted opportunity value",
            ),
            (
                "Status",
                pipeline_state(row),
                "Current state",
            ),
        ]
    )

    st.markdown(
        '<div class="section-title">Pipeline Details</div>',
        unsafe_allow_html=True,
    )

    st.dataframe(
        rows,
        width="stretch",
        hide_index=True,
    )


# =========================================================
# SHEET PAGE
# =========================================================

def page_sheet():

    sheet_name = st.session_state.sheet

    if not sheet_name:

        go("home")

        return

    df = load_google_sheet(
        sheet_name
    )

    if is_formwork_sheet(
        sheet_name
    ):

        show_formwork_dashboard(
            sheet_name,
            df,
        )

    else:

        show_generic_dashboard(
            sheet_name,
            df,
        )


# =========================================================
# SIDEBAR
# =========================================================

def sidebar(model):

    with st.sidebar:

        # Logo
        if os.path.exists(
            LOGO_PATH
        ):

            st.image(
                LOGO_PATH,
                width=180,
            )

        else:

            st.markdown(
                f"""
                <h2>
                    🏗️ {COMPANY}
                </h2>
                """,
                unsafe_allow_html=True,
            )

        st.divider()

        st.markdown(
            "### 🏗️ Navigation"
        )

        if st.button(
            "🏠 Home",
            width="stretch",
        ):

            go("home")

        if st.button(
            "📊 Projects",
            width="stretch",
        ):

            go("projects")

        # -------------------------------------------------
        # Sheets
        # -------------------------------------------------

        st.markdown(
            "### 📄 Google Sheets"
        )

        try:

            sheet_names = get_sheet_names()

            for sheet in sheet_names:

                if st.button(
                    f"📄 {sheet}",
                    key=f"sheet_btn_{sheet}",
                    width="stretch",
                ):

                    open_sheet(
                        sheet
                    )

        except Exception as error:

            st.error(
                f"Unable to load sheets: {error}"
            )

        st.divider()

        # -------------------------------------------------
        # Refresh
        # -------------------------------------------------

        if st.button(
            "🔄 Refresh Google Sheets",
            width="stretch",
        ):

            st.cache_data.clear()

            st.rerun()

        st.caption(
            "Data source: Google Sheets"
        )


# =========================================================
# VIEWER DASHBOARD
# =========================================================

def page_viewer_dashboard(model):

    """
    Public dashboard.

    Viewer URL:

    https://YOUR-APP.streamlit.app/?view=dashboard

    """

    st.markdown(
        f"""
        <div class="hero">

            <h1>
                🏗️ {COMPANY}
            </h1>

            <p>
                Construction Projects Dashboard
            </p>

        </div>
        """,
        unsafe_allow_html=True,
    )

    show_delivery_dashboard(
        model
    )


# =========================================================
# MAIN
# =========================================================

try:

    model = build_model()

except Exception as error:

    st.error(
        "Unable to connect to Google Sheets."
    )

    st.exception(error)

    st.stop()


# =========================================================
# VIEWER MODE
# =========================================================

if VIEWER_MODE:

    page_viewer_dashboard(
        model
    )

    st.stop()


# =========================================================
# ADMIN MODE
# =========================================================

sidebar(model)


page = st.session_state.page


if page == "home":

    page_home(
        model
    )

elif page == "projects":

    page_projects(
        model
    )

elif page == "project":

    page_project(
        model
    )

elif page == "pipeline":

    page_pipeline(
        model
    )

elif page == "sheet":

    page_sheet()

else:

    page_home(
        model
    )
