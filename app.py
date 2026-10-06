"""
Hassan Allam Construction
Dynamic Construction Projects Dashboard
---------------------------------------

Excel Workbook
│
├── Project 1
├── Project 2
├── Project 3
├── Project 4
└── Project 5

Every new Excel Sheet is automatically detected
as a Project.

No project names need to be added to Python code.
"""

import os
import re
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Hassan Allam Construction | Projects Dashboard",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# COMPANY SETTINGS
# ============================================================

COMPANY = "Hassan Allam Construction"

PRIMARY = "#C8102E"
DARK = "#1F2933"
LIGHT = "#F4F5F7"
GRAY = "#6B7280"
GREEN = "#2E9E5B"
ORANGE = "#D97706"
BORDER = "#E1E4E8"

LOGO_PATH = "assets/logo.png"


# ============================================================
# STATUS COLORS
# ============================================================

STATUS_COLORS = {
    "Completed": "#2E9E5B",
    "Active": "#2563EB",
    "In Progress": "#2563EB",
    "Delayed": "#C8102E",
    "Pending": "#D97706",
    "Not Started": "#6B7280",
    "Cancelled": "#7F1D1D",
}


# ============================================================
# CSS
# ============================================================

st.markdown(
    f"""
    <style>

    .stApp {{
        background-color: #F7F7F5;
    }}

    [data-testid="stSidebar"] {{
        background-color: #FFFFFF;
        border-right: 1px solid {BORDER};
    }}

    [data-testid="stSidebar"] * {{
        color: {DARK};
    }}

    .main-title {{
        font-size: 34px;
        font-weight: 800;
        color: {DARK};
        margin-bottom: 5px;
    }}

    .subtitle {{
        color: {GRAY};
        font-size: 15px;
        margin-bottom: 25px;
    }}

    .hero {{
        background: linear-gradient(
            135deg,
            #FFFFFF 0%,
            #F7F7F7 100%
        );
        border: 1px solid {BORDER};
        border-radius: 16px;
        padding: 28px;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.04);
    }}

    .hero-title {{
        color: {DARK};
        font-size: 30px;
        font-weight: 800;
    }}

    .hero-text {{
        color: {GRAY};
        font-size: 15px;
    }}

    .kpi-card {{
        background: white;
        border: 1px solid {BORDER};
        border-radius: 14px;
        padding: 18px;
        min-height: 120px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.03);
    }}

    .kpi-label {{
        color: {GRAY};
        font-size: 13px;
        font-weight: 600;
    }}

    .kpi-value {{
        color: {DARK};
        font-size: 27px;
        font-weight: 800;
        margin-top: 8px;
    }}

    .section-title {{
        font-size: 22px;
        font-weight: 800;
        color: {DARK};
        margin-top: 25px;
        margin-bottom: 15px;
    }}

    .project-card {{
        background: white;
        border: 1px solid {BORDER};
        border-radius: 14px;
        padding: 15px;
        margin-bottom: 10px;
    }}

    .badge {{
        display: inline-block;
        padding: 5px 10px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 700;
        background: #F1F5F9;
    }}

    div.stButton > button {{
        border-radius: 10px;
        border: 1px solid {BORDER};
        background-color: white;
        color: {DARK};
        font-weight: 600;
        min-height: 44px;
    }}

    div.stButton > button:hover {{
        border-color: {PRIMARY};
        color: {PRIMARY};
    }}

    .footer {{
        color: {GRAY};
        font-size: 12px;
        text-align: center;
        padding: 25px;
    }}

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SESSION STATE
# ============================================================

defaults = {
    "page": "home",
    "file_bytes": None,
    "file_name": None,
    "selected_project": None,
    "selected_sheet": None,
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# BASIC HELPERS
# ============================================================

def normalize_text(value):
    if value is None:
        return ""

    text = str(value).strip().lower()

    text = re.sub(r"\s+", " ", text)

    return text


def normalize_key(value):
    text = normalize_text(value)

    text = re.sub(r"[^a-z0-9]+", "_", text)

    return text.strip("_")


def clean_columns(df):
    df = df.copy()

    new_columns = []

    for col in df.columns:
        value = str(col).strip()

        if value == "":
            value = "Unnamed"

        new_columns.append(value)

    df.columns = new_columns

    return df


def fmt_number(value):
    try:
        if pd.isna(value):
            return "0"

        value = float(value)

        if value.is_integer():
            return f"{int(value):,}"

        return f"{value:,.2f}"

    except Exception:
        return str(value)


def to_numeric(series):
    if series is None:
        return pd.Series(dtype=float)

    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("EGP", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.strip(),
        errors="coerce",
    )


def safe_sum(series):
    try:
        return float(to_numeric(series).fillna(0).sum())
    except Exception:
        return 0


def percentage(value):
    try:
        return f"{float(value):.1f}%"
    except Exception:
        return "0.0%"


# ============================================================
# FILE FUNCTIONS
# ============================================================

def get_excel_sheets(file_bytes):
    try:
        excel = pd.ExcelFile(BytesIO(file_bytes))
        return excel.sheet_names
    except Exception:
        return []


def load_sheet(file_bytes, sheet_name):
    """
    Load Excel sheet and automatically detect the header row.
    """

    try:

        raw = pd.read_excel(
            BytesIO(file_bytes),
            sheet_name=sheet_name,
            header=None,
        )

        raw = raw.dropna(
            axis=0,
            how="all",
        )

        raw = raw.dropna(
            axis=1,
            how="all",
        )

        if raw.empty:
            return pd.DataFrame()

        # Detect best header row
        best_row = 0
        best_score = -1

        max_check = min(len(raw), 30)

        for i in range(max_check):

            row = raw.iloc[i]

            non_empty = row.notna().sum()

            text_like = 0

            for value in row:
                if pd.notna(value):
                    if isinstance(value, str):
                        text_like += 1

            score = non_empty * 2 + text_like

            if score > best_score:
                best_score = score
                best_row = i

        df = pd.read_excel(
            BytesIO(file_bytes),
            sheet_name=sheet_name,
            header=best_row,
        )

        df = df.dropna(
            axis=0,
            how="all",
        )

        df = df.dropna(
            axis=1,
            how="all",
        )

        df = clean_columns(df)

        return df

    except Exception as e:

        st.error(
            f"Could not load sheet `{sheet_name}`: {e}"
        )

        return pd.DataFrame()


# ============================================================
# COLUMN DETECTION
# ============================================================

def find_column(df, keywords):

    if df is None or df.empty:
        return None

    normalized = {
        col: normalize_text(col)
        for col in df.columns
    }

    # Exact / contains matching
    for keyword in keywords:

        keyword = normalize_text(keyword)

        for col, normalized_col in normalized.items():

            if keyword == normalized_col:
                return col

    for keyword in keywords:

        keyword = normalize_text(keyword)

        for col, normalized_col in normalized.items():

            if keyword in normalized_col:
                return col

    return None


def numeric_columns(df):

    result = []

    for col in df.columns:

        converted = to_numeric(df[col])

        numeric_count = converted.notna().sum()

        total = len(df)

        if total > 0:

            ratio = numeric_count / total

            if ratio >= 0.5:
                result.append(col)

    return result


def categorical_columns(df):

    result = []

    for col in df.columns:

        if df[col].dtype == "object":

            unique = df[col].nunique(dropna=True)

            if 1 < unique <= 30:
                result.append(col)

    return result


def date_columns(df):

    result = []

    for col in df.columns:

        parsed = pd.to_datetime(
            df[col],
            errors="coerce",
        )

        if len(df) > 0:

            ratio = parsed.notna().sum() / len(df)

            if ratio >= 0.6:
                result.append(col)

    return result


# ============================================================
# PROJECT NAME
# ============================================================

def pretty_project_name(sheet_name):

    name = str(sheet_name)

    name = re.sub(
        r"[_\-]+",
        " ",
        name,
    )

    name = re.sub(
        r"\s+",
        " ",
        name,
    )

    return name.strip()


# ============================================================
# PROJECT STATUS
# ============================================================

def detect_status(df):

    status_col = find_column(
        df,
        [
            "status",
            "project status",
            "state",
            "progress status",
            "delivery status",
        ],
    )

    if not status_col:
        return "Active"

    values = (
        df[status_col]
        .dropna()
        .astype(str)
        .str.strip()
    )

    if values.empty:
        return "Active"

    value_counts = values.value_counts()

    raw_status = str(
        value_counts.index[0]
    ).strip()

    normalized = normalize_text(raw_status)

    if "complete" in normalized:
        return "Completed"

    if "delay" in normalized:
        return "Delayed"

    if "cancel" in normalized:
        return "Cancelled"

    if "pending" in normalized:
        return "Pending"

    if "not started" in normalized:
        return "Not Started"

    if "progress" in normalized:
        return "In Progress"

    if "active" in normalized:
        return "Active"

    return raw_status


# ============================================================
# PROGRESS DETECTION
# ============================================================

def detect_progress(df):

    progress_col = find_column(
        df,
        [
            "progress",
            "completion",
            "completion %",
            "progress %",
            "% complete",
            "% completion",
        ],
    )

    if progress_col:

        values = to_numeric(
            df[progress_col]
        )

        values = values.dropna()

        if not values.empty:

            avg = values.mean()

            if avg <= 1:
                avg *= 100

            return max(
                0,
                min(100, avg),
            )

    # Quantity based progress
    planned_col = find_column(
        df,
        [
            "planned quantity",
            "planned qty",
            "target quantity",
            "planned",
        ],
    )

    actual_col = find_column(
        df,
        [
            "actual quantity",
            "actual qty",
            "completed quantity",
            "actual",
        ],
    )

    if planned_col and actual_col:

        planned = safe_sum(
            df[planned_col]
        )

        actual = safe_sum(
            df[actual_col]
        )

        if planned > 0:

            return max(
                0,
                min(
                    100,
                    actual / planned * 100,
                ),
            )

    return None


# ============================================================
# COST DETECTION
# ============================================================

def detect_cost_columns(df):

    budget = find_column(
        df,
        [
            "budget",
            "budget cost",
            "budget amount",
            "estimated cost",
            "contract value",
        ],
    )

    planned = find_column(
        df,
        [
            "planned cost",
            "planned amount",
            "planned",
        ],
    )

    actual = find_column(
        df,
        [
            "actual cost",
            "actual amount",
            "actual",
            "spent",
            "cost incurred",
        ],
    )

    return budget, planned, actual


# ============================================================
# FORMATTING
# ============================================================

def kpi_card(label, value):

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def kpi_row(items):

    cols = st.columns(len(items))

    for col, (label, value) in zip(
        cols,
        items,
    ):

        with col:
            kpi_card(
                label,
                value,
            )


# ============================================================
# NAVIGATION
# ============================================================

def go(page):

    st.session_state.page = page
    st.rerun()


def open_project(sheet_name):

    st.session_state.selected_project = sheet_name
    st.session_state.selected_sheet = sheet_name
    st.session_state.page = "project"

    st.rerun()


def back_home():

    st.session_state.page = "home"
    st.session_state.selected_project = None

    st.rerun()


# ============================================================
# PROJECT SHEET DISCOVERY
# ============================================================

def get_project_sheets(file_bytes):

    sheets = get_excel_sheets(file_bytes)

    ignored = {
        "readme",
        "settings",
        "config",
        "lookup",
        "lookups",
        "parameters",
        "parameter",
    }

    projects = []

    for sheet in sheets:

        normalized = normalize_key(sheet)

        if normalized in ignored:
            continue

        if not normalized:
            continue

        projects.append(sheet)

    return projects


# ============================================================
# PROJECT SUMMARY
# ============================================================

def project_summary(sheet_name, df):

    budget_col, planned_col, actual_col = detect_cost_columns(
        df
    )

    budget = (
        safe_sum(df[budget_col])
        if budget_col
        else 0
    )

    planned = (
        safe_sum(df[planned_col])
        if planned_col
        else 0
    )

    actual = (
        safe_sum(df[actual_col])
        if actual_col
        else 0
    )

    if budget == 0 and planned > 0:
        budget = planned

    remaining = budget - actual

    progress = detect_progress(df)

    if progress is None:

        if budget > 0:
            progress = actual / budget * 100
        else:
            progress = 0

    status = detect_status(df)

    return {
        "Project": pretty_project_name(sheet_name),
        "Budget": budget,
        "Planned": planned,
        "Actual": actual,
        "Remaining": remaining,
        "Progress": progress,
        "Status": status,
        "Rows": len(df),
        "Columns": len(df.columns),
    }


# ============================================================
# PROJECT TYPE DETECTION
# ============================================================

def detect_project_type(df):

    if df.empty:
        return "Generic"

    columns = " ".join(
        normalize_text(col)
        for col in df.columns
    )

    if any(
        word in columns
        for word in [
            "formwork",
            "unit rate",
            "budget cost",
            "element / activity",
        ]
    ):
        return "Formwork / Cost"

    if any(
        word in columns
        for word in [
            "delivery",
            "expected date",
            "delivery date",
            "delivered",
        ]
    ):
        return "Delivery"

    if any(
        word in columns
        for word in [
            "pipeline",
            "tender",
            "opportunity",
        ]
    ):
        return "Pipeline"

    if any(
        word in columns
        for word in [
            "cost",
            "budget",
            "actual",
            "planned",
        ]
    ):
        return "Financial"

    return "Generic"


# ============================================================
# FILTER DATA
# ============================================================

def apply_filters(df, key_prefix="filter"):

    filtered = df.copy()

    st.markdown(
        '<div class="section-title">Filters</div>',
        unsafe_allow_html=True,
    )

    categorical = categorical_columns(
        filtered
    )

    # Search
    search = st.text_input(
        "Search",
        key=f"{key_prefix}_search",
        placeholder="Search in project data...",
    )

    if search:

        search = search.lower()

        mask = filtered.astype(
            str
        ).apply(
            lambda col: col.str.lower().str.contains(
                search,
                na=False,
            )
        ).any(axis=1)

        filtered = filtered[mask]

    # Select filters
    filter_columns = categorical[:4]

    if filter_columns:

        cols = st.columns(
            len(filter_columns)
        )

        for col_ui, column in zip(
            cols,
            filter_columns,
        ):

            values = (
                filtered[column]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

            values = sorted(
                values,
                key=str,
            )

            if len(values) <= 30:

                with col_ui:

                    selected = st.multiselect(
                        column,
                        values,
                        key=f"{key_prefix}_{normalize_key(column)}",
                    )

                    if selected:

                        filtered = filtered[
                            filtered[column]
                            .astype(str)
                            .isin(selected)
                        ]

    return filtered


# ============================================================
# GENERIC CHARTS
# ============================================================

def show_generic_charts(df):

    numeric = numeric_columns(df)

    categorical = categorical_columns(df)

    if not numeric:
        return

    st.markdown(
        '<div class="section-title">Analytics</div>',
        unsafe_allow_html=True,
    )

    if categorical:

        category = categorical[0]
        value = numeric[0]

        chart_df = (
            df.groupby(
                category,
                dropna=False,
            )[value]
            .sum()
            .reset_index()
        )

        if not chart_df.empty:

            fig = px.bar(
                chart_df,
                x=category,
                y=value,
                title=f"{value} by {category}",
            )

            fig.update_layout(
                template="plotly_white",
                height=400,
            )

            st.plotly_chart(
                fig,
                use_container_width=True,
            )

    if len(numeric) >= 2:

        x_col = numeric[0]
        y_col = numeric[1]

        fig = px.scatter(
            df,
            x=x_col,
            y=y_col,
            title=f"{y_col} vs {x_col}",
        )

        fig.update_layout(
            template="plotly_white",
            height=400,
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
        )


# ============================================================
# FINANCIAL CHART
# ============================================================

def show_financial_chart(df):

    budget_col, planned_col, actual_col = detect_cost_columns(
        df
    )

    if not any(
        [
            budget_col,
            planned_col,
            actual_col,
        ]
    ):
        return

    data = []

    if budget_col:
        data.append(
            [
                "Budget",
                safe_sum(df[budget_col]),
            ]
        )

    if planned_col:
        data.append(
            [
                "Planned",
                safe_sum(df[planned_col]),
            ]
        )

    if actual_col:
        data.append(
            [
                "Actual",
                safe_sum(df[actual_col]),
            ]
        )

    if not data:
        return

    chart_df = pd.DataFrame(
        data,
        columns=[
            "Type",
            "Amount",
        ],
    )

    fig = px.bar(
        chart_df,
        x="Type",
        y="Amount",
        text_auto=".2s",
        title="Project Cost Overview",
    )

    fig.update_layout(
        template="plotly_white",
        height=400,
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
    )


# ============================================================
# STATUS CHART
# ============================================================

def show_status_chart(df):

    status_col = find_column(
        df,
        [
            "status",
            "project status",
            "state",
            "delivery status",
        ],
    )

    if not status_col:
        return

    chart_df = (
        df[status_col]
        .fillna("Unknown")
        .astype(str)
        .value_counts()
        .reset_index()
    )

    chart_df.columns = [
        "Status",
        "Count",
    ]

    if chart_df.empty:
        return

    fig = px.pie(
        chart_df,
        names="Status",
        values="Count",
        hole=0.45,
        title="Status Distribution",
    )

    fig.update_layout(
        template="plotly_white",
        height=400,
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
    )


# ============================================================
# DATE CHART
# ============================================================

def show_date_chart(df):

    dates = date_columns(df)

    numeric = numeric_columns(df)

    if not dates:
        return

    date_col = dates[0]

    temp = df.copy()

    temp["_date_"] = pd.to_datetime(
        temp[date_col],
        errors="coerce",
    )

    temp = temp.dropna(
        subset=["_date_"]
    )

    if temp.empty:
        return

    if numeric:

        value_col = numeric[0]

        temp["_value_"] = to_numeric(
            temp[value_col]
        )

        chart_df = (
            temp.groupby(
                "_date_"
            )["_value_"]
            .sum()
            .reset_index()
        )

        fig = px.line(
            chart_df,
            x="_date_",
            y="_value_",
            markers=True,
            title=f"{value_col} Trend",
        )

    else:

        chart_df = (
            temp.groupby(
                "_date_"
            )
            .size()
            .reset_index(
                name="Count"
            )
        )

        fig = px.line(
            chart_df,
            x="_date_",
            y="Count",
            markers=True,
            title="Records Over Time",
        )

    fig.update_layout(
        template="plotly_white",
        height=400,
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
    )


# ============================================================
# HOME PAGE
# ============================================================

def page_home():

    st.markdown(
        '<div class="main-title">Construction Projects Dashboard</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        f'<div class="subtitle">{COMPANY}</div>',
        unsafe_allow_html=True,
    )

    if not st.session_state.file_bytes:

        st.markdown(
            f"""
            <div class="hero">

                <div class="hero-title">
                    🏗️ Construction Projects & Pipeline
                </div>

                <br>

                <div class="hero-text">
                    Upload your Excel workbook to automatically
                    discover Projects and build dashboards.
                </div>

                <br>

                <div class="hero-text">
                    Every new Excel Sheet can represent a new Project.
                    The application will detect it automatically.
                </div>

            </div>
            """,
            unsafe_allow_html=True,
        )

        st.info(
            "Upload an Excel file from the sidebar to start."
        )

        return

    sheets = get_project_sheets(
        st.session_state.file_bytes
    )

    st.markdown(
        f"""
        <div class="hero">

            <div class="hero-title">
                🏗️ {COMPANY}
            </div>

            <br>

            <div class="hero-text">
                Dynamic Construction Projects Dashboard
            </div>

            <div class="hero-text">
                Workbook: {st.session_state.file_name}
            </div>

        </div>
        """,
        unsafe_allow_html=True,
    )

    # Load summaries
    summaries = []

    for sheet in sheets:

        df = load_sheet(
            st.session_state.file_bytes,
            sheet,
        )

        if not df.empty:

            summaries.append(
                project_summary(
                    sheet,
                    df,
                )
            )

    # Workbook KPIs
    total_projects = len(summaries)

    total_budget = sum(
        item["Budget"]
        for item in summaries
    )

    total_actual = sum(
        item["Actual"]
        for item in summaries
    )

    total_remaining = sum(
        item["Remaining"]
        for item in summaries
    )

    avg_progress = (
        sum(
            item["Progress"]
            for item in summaries
        )
        / total_projects
        if total_projects
        else 0
    )

    kpi_row(
        [
            (
                "Projects",
                fmt_number(total_projects),
            ),
            (
                "Total Budget",
                fmt_number(total_budget),
            ),
            (
                "Actual Cost",
                fmt_number(total_actual),
            ),
            (
                "Remaining",
                fmt_number(total_remaining),
            ),
            (
                "Avg. Progress",
                percentage(avg_progress),
            ),
        ]
    )

    # Projects
    st.markdown(
        '<div class="section-title">Projects</div>',
        unsafe_allow_html=True,
    )

    if not sheets:

        st.warning(
            "No project sheets were detected."
        )

        return

    # Buttons
    columns_per_row = 3

    for start in range(
        0,
        len(sheets),
        columns_per_row,
    ):

        row_sheets = sheets[
            start:start + columns_per_row
        ]

        cols = st.columns(
            columns_per_row
        )

        for col, sheet in zip(
            cols,
            row_sheets,
        ):

            with col:

                df = load_sheet(
                    st.session_state.file_bytes,
                    sheet,
                )

                summary = project_summary(
                    sheet,
                    df,
                )

                st.markdown(
                    f"""
                    <div class="project-card">

                    <b>
                        🏗️
                        {pretty_project_name(sheet)}
                    </b>

                    <br><br>

                    <span class="badge">
                        {summary["Status"]}
                    </span>

                    <br><br>

                    Progress:
                    <b>
                        {percentage(summary["Progress"])}
                    </b>

                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                if st.button(
                    "Open Project",
                    key=f"open_project_{start}_{normalize_key(sheet)}",
                    use_container_width=True,
                ):

                    open_project(sheet)

    # Structure
    st.markdown(
        '<div class="section-title">Detected Workbook Structure</div>',
        unsafe_allow_html=True,
    )

    structure_data = []

    for sheet in sheets:

        df = load_sheet(
            st.session_state.file_bytes,
            sheet,
        )

        structure_data.append(
            {
                "Sheet": sheet,
                "Project": pretty_project_name(sheet),
                "Rows": len(df),
                "Columns": len(df.columns),
                "Type": detect_project_type(df),
            }
        )

    if structure_data:

        st.dataframe(
            pd.DataFrame(structure_data),
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# PROJECT PAGE
# ============================================================

def page_project():

    sheet = st.session_state.selected_project

    if not sheet:

        back_home()
        return

    df = load_sheet(
        st.session_state.file_bytes,
        sheet,
    )

    if df.empty:

        st.error(
            "The selected project sheet is empty."
        )

        if st.button("← Back to Projects"):
            back_home()

        return

    project_name = pretty_project_name(
        sheet
    )

    summary = project_summary(
        sheet,
        df,
    )

    # Header
    col1, col2 = st.columns(
        [5, 1]
    )

    with col1:

        st.markdown(
            f"""
            <div class="main-title">
                🏗️ {project_name}
            </div>

            <div class="subtitle">
                {COMPANY}
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col2:

        if st.button(
            "← Projects",
            use_container_width=True,
        ):

            back_home()

    # KPIs
    kpi_row(
        [
            (
                "Budget",
                fmt_number(summary["Budget"]),
            ),
            (
                "Planned",
                fmt_number(summary["Planned"]),
            ),
            (
                "Actual",
                fmt_number(summary["Actual"]),
            ),
            (
                "Remaining",
                fmt_number(summary["Remaining"]),
            ),
            (
                "Progress",
                percentage(summary["Progress"]),
            ),
        ]
    )

    st.markdown("")

    # Status + type
    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Status",
            summary["Status"],
        )

    with col2:

        st.metric(
            "Data Rows",
            f"{summary['Rows']:,}",
        )

    with col3:

        st.metric(
            "Project Type",
            detect_project_type(df),
        )

    # Progress
    st.markdown(
        '<div class="section-title">Project Progress</div>',
        unsafe_allow_html=True,
    )

    progress = summary["Progress"]

    st.progress(
        max(
            0,
            min(
                100,
                int(progress),
            ),
        )
    )

    st.write(
        f"Project completion: **{percentage(progress)}**"
    )

    # Analytics
    col1, col2 = st.columns(2)

    with col1:
        show_financial_chart(df)

    with col2:
        show_status_chart(df)

    show_date_chart(df)

    show_generic_charts(df)

    # Filters
    filtered = apply_filters(
        df,
        key_prefix=f"project_{normalize_key(sheet)}",
    )

    # Data table
    st.markdown(
        '<div class="section-title">Project Data</div>',
        unsafe_allow_html=True,
    )

    st.caption(
        f"Showing {len(filtered):,} of {len(df):,} rows"
    )

    st.dataframe(
        filtered,
        use_container_width=True,
        height=500,
        hide_index=True,
    )

    # Download
    csv = filtered.to_csv(
        index=False
    ).encode("utf-8")

    st.download_button(
        "Download Project Data",
        data=csv,
        file_name=f"{normalize_key(sheet)}_data.csv",
        mime="text/csv",
        use_container_width=True,
    )


# ============================================================
# SIDEBAR
# ============================================================

def sidebar():

    with st.sidebar:

        # Logo
        if os.path.exists(LOGO_PATH):

            st.image(
                LOGO_PATH,
                use_container_width=True,
            )

        else:

            st.markdown(
                f"""
                <h2 style="color:{PRIMARY};">
                    🏗️
                </h2>

                <h3>
                    {COMPANY}
                </h3>
                """,
                unsafe_allow_html=True,
            )

        st.divider()

        st.subheader(
            "Excel Data"
        )

        uploaded_file = st.file_uploader(
            "Upload Excel Workbook",
            type=[
                "xlsx",
                "xls",
            ],
            key="excel_uploader",
        )

        if uploaded_file:

            file_bytes = uploaded_file.getvalue()

            # Only update when new file
            if (
                st.session_state.file_name
                != uploaded_file.name
            ):

                st.session_state.file_bytes = (
                    file_bytes
                )

                st.session_state.file_name = (
                    uploaded_file.name
                )

                st.session_state.page = "home"

                st.session_state.selected_project = None

        if st.session_state.file_bytes:

            st.success(
                f"Loaded: {st.session_state.file_name}"
            )

        st.divider()

        if st.button(
            "🏠 Home",
            use_container_width=True,
        ):

            st.session_state.page = "home"
            st.session_state.selected_project = None

            st.rerun()

        # Dynamic projects
        if st.session_state.file_bytes:

            st.markdown(
                "**Projects**"
            )

            sheets = get_project_sheets(
                st.session_state.file_bytes
            )

            for sheet in sheets:

                label = pretty_project_name(
                    sheet
                )

                if st.button(
                    f"🏗️ {label}",
                    key=f"sidebar_{normalize_key(sheet)}",
                    use_container_width=True,
                ):

                    open_project(sheet)

        st.divider()

        if st.button(
            "🔄 Refresh Data",
            use_container_width=True,
        ):

            st.rerun()

        st.markdown(
            """
            <div class="footer">
                Dynamic Excel Dashboard<br>
                Hassan Allam Construction
            </div>
            """,
            unsafe_allow_html=True,
        )


# ============================================================
# MAIN ROUTER
# ============================================================

sidebar()


if not st.session_state.file_bytes:

    page_home()

elif st.session_state.page == "home":

    page_home()

elif st.session_state.page == "project":

    page_project()

else:

    page_home()
