```python
"""
Hassan Allam Construction
Dynamic Multi-Project PR Dashboard

============================================================

DATA MODEL

Each Excel file = ONE PROJECT

Example:

    Sofitel Hotel.xlsx
    New Cairo Project.xlsx
    North Coast Project.xlsx
    Al Shadda Project.xlsx

The file name automatically becomes the Project name.

The application supports multiple Excel files.

Dashboard flow:

    Project Filter
          ↓
       PR Filter
          ↓
    Other Filters
          ↓
         KPIs
          ↓
        Tables

NO CHARTS.

============================================================
"""

import os
import re
from io import BytesIO

import pandas as pd
import streamlit as st


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Hassan Allam Construction | PR Dashboard",
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
        min-height: 115px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.03);
    }}

    .kpi-label {{
        color: {GRAY};
        font-size: 13px;
        font-weight: 600;
    }}

    .kpi-value {{
        color: {DARK};
        font-size: 26px;
        font-weight: 800;
        margin-top: 8px;
    }}

    .section-title {{
        font-size: 21px;
        font-weight: 800;
        color: {DARK};
        margin-top: 25px;
        margin-bottom: 15px;
    }}

    .info-box {{
        background: #FFFFFF;
        border: 1px solid {BORDER};
        border-radius: 12px;
        padding: 16px;
        margin-bottom: 15px;
    }}

    .info-label {{
        color: {GRAY};
        font-size: 12px;
        font-weight: 600;
    }}

    .info-value {{
        color: {DARK};
        font-size: 17px;
        font-weight: 700;
        margin-top: 4px;
    }}

    div.stButton > button {{
        border-radius: 10px;
        border: 1px solid {BORDER};
        background-color: white;
        color: {DARK};
        font-weight: 600;
        min-height: 42px;
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
    "projects": {},
    "selected_project": "All Projects",
    "selected_pr": "All PRs",
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

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text


def normalize_key(value):

    text = normalize_text(value)

    text = re.sub(
        r"[^a-z0-9]+",
        "_",
        text,
    )

    return text.strip("_")


def clean_columns(df):

    df = df.copy()

    columns = []

    for col in df.columns:

        value = str(col).strip()

        if not value:
            value = "Unnamed"

        columns.append(value)

    df.columns = columns

    return df


def pretty_project_name(file_name):

    name = os.path.splitext(
        os.path.basename(file_name)
    )[0]

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


def percentage(value):

    try:
        return f"{float(value):.1f}%"

    except Exception:
        return "0.0%"


def to_numeric(series):

    if series is None:
        return pd.Series(dtype=float)

    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("EGP", "", regex=False)
        .str.replace("LE", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.strip(),
        errors="coerce",
    )


def safe_sum(series):

    try:

        return float(
            to_numeric(series)
            .fillna(0)
            .sum()
        )

    except Exception:

        return 0


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

    # Exact
    for keyword in keywords:

        keyword = normalize_text(keyword)

        for col, normalized_col in normalized.items():

            if keyword == normalized_col:
                return col

    # Contains
    for keyword in keywords:

        keyword = normalize_text(keyword)

        for col, normalized_col in normalized.items():

            if keyword in normalized_col:
                return col

    return None


# ============================================================
# PR COLUMN
# ============================================================

def find_pr_column(df):

    return find_column(
        df,
        [
            "pr number",
            "pr no",
            "pr no.",
            "pr#",
            "pr",
            "purchase requisition",
            "purchase requisition number",
            "purchase requisition no",
            "requisition number",
            "requisition no",
        ],
    )


# ============================================================
# COST COLUMNS
# ============================================================

def find_budget_column(df):

    return find_column(
        df,
        [
            "budget cost",
            "budget amount",
            "budget",
            "estimated cost",
            "estimated amount",
            "contract value",
            "total budget",
        ],
    )


def find_actual_column(df):

    return find_column(
        df,
        [
            "actual cost",
            "actual amount",
            "actual",
            "spent",
            "cost incurred",
            "paid amount",
            "total actual",
        ],
    )


def find_planned_column(df):

    return find_column(
        df,
        [
            "planned cost",
            "planned amount",
            "planned",
            "forecast",
            "forecast cost",
        ],
    )


def find_quantity_column(df):

    return find_column(
        df,
        [
            "quantity",
            "qty",
            "total quantity",
            "ordered quantity",
            "order quantity",
        ],
    )


# ============================================================
# LOAD EXCEL
# ============================================================

def detect_header_row(raw):

    if raw.empty:
        return 0

    best_row = 0
    best_score = -1

    max_check = min(
        len(raw),
        40,
    )

    for i in range(max_check):

        row = raw.iloc[i]

        non_empty = row.notna().sum()

        text_like = 0

        for value in row:

            if pd.notna(value):

                if isinstance(
                    value,
                    str,
                ):
                    text_like += 1

        score = (
            non_empty * 2
            + text_like
        )

        # Give extra weight to rows containing
        # important business headers
        row_text = " ".join(
            normalize_text(value)
            for value in row
            if pd.notna(value)
        )

        important_words = [
            "pr",
            "purchase",
            "item",
            "description",
            "quantity",
            "qty",
            "cost",
            "amount",
            "budget",
            "supplier",
            "vendor",
        ]

        for word in important_words:

            if word in row_text:
                score += 3

        if score > best_score:

            best_score = score
            best_row = i

    return best_row


def load_excel_file(file_bytes):

    try:

        excel = pd.ExcelFile(
            BytesIO(file_bytes)
        )

        sheets = excel.sheet_names

        all_data = []

        for sheet in sheets:

            raw = pd.read_excel(
                BytesIO(file_bytes),
                sheet_name=sheet,
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
                continue

            header_row = detect_header_row(
                raw
            )

            df = pd.read_excel(
                BytesIO(file_bytes),
                sheet_name=sheet,
                header=header_row,
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

            if not df.empty:

                # Add source sheet
                df["_Source Sheet"] = sheet

                all_data.append(df)

        if not all_data:
            return pd.DataFrame()

        # Combine sheets.
        # Missing columns automatically become NaN.
        result = pd.concat(
            all_data,
            ignore_index=True,
            sort=False,
        )

        return result

    except Exception as e:

        st.error(
            f"Could not read Excel file: {e}"
        )

        return pd.DataFrame()


# ============================================================
# PROJECT REGISTRY
# ============================================================

def register_uploaded_files(uploaded_files):

    if not uploaded_files:
        return

    for uploaded_file in uploaded_files:

        file_name = uploaded_file.name

        file_bytes = uploaded_file.getvalue()

        project_name = pretty_project_name(
            file_name
        )

        df = load_excel_file(
            file_bytes
        )

        if df.empty:
            continue

        st.session_state.projects[
            project_name
        ] = {
            "file_name": file_name,
            "file_bytes": file_bytes,
            "data": df,
        }


# ============================================================
# PROJECT FILTER
# ============================================================

def project_filter():

    projects = sorted(
        st.session_state.projects.keys(),
        key=str.lower,
    )

    options = [
        "All Projects"
    ] + projects

    current = st.session_state.get(
        "selected_project",
        "All Projects",
    )

    if current not in options:
        current = "All Projects"

    selected = st.selectbox(
        "Project",
        options,
        index=options.index(current),
        key="project_selector",
    )

    st.session_state.selected_project = selected

    return selected


# ============================================================
# GET SELECTED DATA
# ============================================================

def get_project_data():

    selected = st.session_state.selected_project

    if selected == "All Projects":

        frames = []

        for project_name, info in (
            st.session_state.projects.items()
        ):

            df = info["data"].copy()

            df.insert(
                0,
                "Project",
                project_name,
            )

            frames.append(df)

        if not frames:
            return pd.DataFrame()

        return pd.concat(
            frames,
            ignore_index=True,
            sort=False,
        )

    if selected in st.session_state.projects:

        df = st.session_state.projects[
            selected
        ]["data"].copy()

        if "Project" not in df.columns:

            df.insert(
                0,
                "Project",
                selected,
            )

        return df

    return pd.DataFrame()


# ============================================================
# PR FILTER
# ============================================================

def pr_filter(df):

    if df.empty:
        return df

    pr_col = find_pr_column(df)

    if not pr_col:

        st.info(
            "No PR column was detected in the selected project."
        )

        st.session_state.selected_pr = "All PRs"

        return df

    values = (
        df[pr_col]
        .dropna()
        .astype(str)
        .str.strip()
    )

    values = [
        value
        for value in values.unique().tolist()
        if value
    ]

    values = sorted(
        values,
        key=str.lower,
    )

    options = [
        "All PRs"
    ] + values

    current = st.session_state.get(
        "selected_pr",
        "All PRs",
    )

    if current not in options:
        current = "All PRs"

    selected = st.selectbox(
        f"PR ({pr_col})",
        options,
        index=options.index(current),
        key=f"pr_selector_{normalize_key(str(st.session_state.selected_project))}",
    )

    st.session_state.selected_pr = selected

    if selected != "All PRs":

        df = df[
            df[pr_col]
            .astype(str)
            .str.strip()
            .eq(selected)
        ].copy()

    return df


# ============================================================
# DYNAMIC FILTERS
# ============================================================

def apply_dynamic_filters(df):

    if df.empty:
        return df

    excluded = {
        "project",
        "_source sheet",
    }

    candidate_columns = []

    for col in df.columns:

        normalized = normalize_text(col)

        if normalized in excluded:
            continue

        # Do not create filters for obvious numeric columns
        if pd.api.types.is_numeric_dtype(
            df[col]
        ):
            continue

        unique_count = (
            df[col]
            .dropna()
            .astype(str)
            .nunique()
        )

        if (
            unique_count > 1
            and unique_count <= 40
        ):
            candidate_columns.append(
                col
            )

    # PR is already handled separately
    pr_col = find_pr_column(df)

    if pr_col in candidate_columns:
        candidate_columns.remove(
            pr_col
        )

    # Prefer business columns
    priority_words = [
        "supplier",
        "vendor",
        "category",
        "status",
        "department",
        "section",
        "item",
        "material",
        "type",
        "location",
    ]

    candidate_columns.sort(
        key=lambda col: (
            0
            if any(
                word in normalize_text(col)
                for word in priority_words
            )
            else 1,
            normalize_text(col),
        )
    )

    # Maximum filters displayed
    candidate_columns = candidate_columns[
        :8
    ]

    if not candidate_columns:
        return df

    st.markdown(
        '<div class="section-title">Additional Filters</div>',
        unsafe_allow_html=True,
    )

    filtered = df.copy()

    # 4 filters per row
    for start in range(
        0,
        len(candidate_columns),
        4,
    ):

        row_columns = candidate_columns[
            start:start + 4
        ]

        cols = st.columns(4)

        for ui_col, data_col in zip(
            cols,
            row_columns,
        ):

            values = (
                filtered[data_col]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
                .tolist()
            )

            values = sorted(
                values,
                key=str.lower,
            )

            with ui_col:

                selected = st.multiselect(
                    data_col,
                    values,
                    key=(
                        f"filter_"
                        f"{normalize_key(str(st.session_state.selected_project))}_"
                        f"{normalize_key(str(st.session_state.selected_pr))}_"
                        f"{normalize_key(data_col)}"
                    ),
                )

                if selected:

                    filtered = filtered[
                        filtered[data_col]
                        .astype(str)
                        .str.strip()
                        .isin(selected)
                    ]

    return filtered


# ============================================================
# SEARCH
# ============================================================

def search_filter(df):

    if df.empty:
        return df

    search = st.text_input(
        "Search",
        placeholder="Search within selected data...",
        key=(
            "global_search_"
            + normalize_key(
                str(
                    st.session_state.selected_project
                )
            )
            + "_"
            + normalize_key(
                str(
                    st.session_state.selected_pr
                )
            )
        ),
    )

    if not search:
        return df

    search = search.lower().strip()

    mask = (
        df.astype(str)
        .apply(
            lambda col:
            col.str.lower()
            .str.contains(
                search,
                na=False,
            )
        )
        .any(axis=1)
    )

    return df[mask].copy()


# ============================================================
# KPI HELPERS
# ============================================================

def kpi_card(label, value):

    st.markdown(
        f"""
        <div class="kpi-card">

            <div class="kpi-label">
                {label}
            </div>

            <div class="kpi-value">
                {value}
            </div>

        </div>
        """,
        unsafe_allow_html=True,
    )


def kpi_row(items):

    if not items:
        return

    cols = st.columns(
        len(items)
    )

    for col, item in zip(
        cols,
        items,
    ):

        label, value = item

        with col:

            kpi_card(
                label,
                value,
            )


# ============================================================
# CALCULATE KPIs
# ============================================================

def calculate_kpis(df):

    if df.empty:

        return {
            "records": 0,
            "prs": 0,
            "budget": 0,
            "planned": 0,
            "actual": 0,
            "remaining": 0,
            "quantity": 0,
        }

    pr_col = find_pr_column(df)

    budget_col = find_budget_column(df)

    planned_col = find_planned_column(df)

    actual_col = find_actual_column(df)

    quantity_col = find_quantity_column(df)

    prs = (
        df[pr_col]
        .dropna()
        .astype(str)
        .str.strip()
        .nunique()
        if pr_col
        else 0
    )

    budget = (
        safe_sum(
            df[budget_col]
        )
        if budget_col
        else 0
    )

    planned = (
        safe_sum(
            df[planned_col]
        )
        if planned_col
        else 0
    )

    actual = (
        safe_sum(
            df[actual_col]
        )
        if actual_col
        else 0
    )

    quantity = (
        safe_sum(
            df[quantity_col]
        )
        if quantity_col
        else 0
    )

    # If budget does not exist,
    # use planned as fallback.
    if budget == 0 and planned != 0:

        budget = planned

    remaining = budget - actual

    return {
        "records": len(df),
        "prs": prs,
        "budget": budget,
        "planned": planned,
        "actual": actual,
        "remaining": remaining,
        "quantity": quantity,
    }


# ============================================================
# KPI SECTION
# ============================================================

def show_kpis(df):

    kpis = calculate_kpis(
        df
    )

    kpi_row(
        [
            (
                "Records",
                fmt_number(
                    kpis["records"]
                ),
            ),
            (
                "PRs",
                fmt_number(
                    kpis["prs"]
                ),
            ),
            (
                "Budget",
                f"{fmt_number(kpis['budget'])} EGP",
            ),
            (
                "Planned",
                f"{fmt_number(kpis['planned'])} EGP",
            ),
            (
                "Actual",
                f"{fmt_number(kpis['actual'])} EGP",
            ),
        ]
    )

    st.markdown("")

    kpi_row(
        [
            (
                "Remaining",
                f"{fmt_number(kpis['remaining'])} EGP",
            ),
            (
                "Quantity",
                fmt_number(
                    kpis["quantity"]
                ),
            ),
        ]
    )


# ============================================================
# PR SUMMARY TABLE
# ============================================================

def show_pr_summary(df):

    pr_col = find_pr_column(df)

    if not pr_col:
        return

    budget_col = find_budget_column(df)

    planned_col = find_planned_column(df)

    actual_col = find_actual_column(df)

    quantity_col = find_quantity_column(df)

    aggregation = {}

    if quantity_col:
        aggregation[
            quantity_col
        ] = "sum"

    if budget_col:
        aggregation[
            budget_col
        ] = "sum"

    if planned_col:
        aggregation[
            planned_col
        ] = "sum"

    if actual_col:
        aggregation[
            actual_col
        ] = "sum"

    if not aggregation:
        return

    summary = (
        df.groupby(
            pr_col,
            dropna=False,
        )
        .agg(aggregation)
        .reset_index()
    )

    # Remaining
    if budget_col and actual_col:

        summary[
            "Remaining"
        ] = (
            to_numeric(
                summary[budget_col]
            )
            - to_numeric(
                summary[actual_col]
            )
        )

    elif planned_col and actual_col:

        summary[
            "Remaining"
        ] = (
            to_numeric(
                summary[planned_col]
            )
            - to_numeric(
                summary[actual_col]
            )
        )

    st.markdown(
        '<div class="section-title">PR Summary</div>',
        unsafe_allow_html=True,
    )

    st.dataframe(
        summary,
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# PROJECT SUMMARY TABLE
# ============================================================

def show_project_summary():

    rows = []

    for project_name, info in (
        st.session_state.projects.items()
    ):

        df = info["data"]

        kpis = calculate_kpis(
            df
        )

        rows.append(
            {
                "Project": project_name,
                "PRs": kpis["prs"],
                "Records": kpis["records"],
                "Budget": kpis["budget"],
                "Planned": kpis["planned"],
                "Actual": kpis["actual"],
                "Remaining": kpis["remaining"],
                "Quantity": kpis["quantity"],
            }
        )

    if not rows:
        return

    summary = pd.DataFrame(
        rows
    )

    st.markdown(
        '<div class="section-title">Projects Summary</div>',
        unsafe_allow_html=True,
    )

    st.dataframe(
        summary,
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# PROJECT INFO
# ============================================================

def show_project_info(df):

    if df.empty:
        return

    project_name = (
        st.session_state.selected_project
    )

    if project_name == "All Projects":
        return

    info_columns = [
        "supplier",
        "vendor",
        "contractor",
        "project manager",
        "department",
        "location",
    ]

    detected = []

    for word in info_columns:

        col = find_column(
            df,
            [word],
        )

        if col and col not in detected:
            detected.append(col)

    if not detected:
        return

    st.markdown(
        '<div class="section-title">Project Information</div>',
        unsafe_allow_html=True,
    )

    cols = st.columns(
        min(
            4,
            len(detected),
        )
    )

    for ui_col, column in zip(
        cols,
        detected,
    ):

        values = (
            df[column]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        if not values:
            continue

        with ui_col:

            st.markdown(
                f"""
                <div class="info-box">

                    <div class="info-label">
                        {column}
                    </div>

                    <div class="info-value">
                        {", ".join(values[:3])}
                    </div>

                </div>
                """,
                unsafe_allow_html=True,
            )


# ============================================================
# MAIN DASHBOARD
# ============================================================

def dashboard():

    st.markdown(
        '<div class="main-title">Construction Cost Dashboard</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        f'<div class="subtitle">{COMPANY}</div>',
        unsafe_allow_html=True,
    )

    if not st.session_state.projects:

        st.markdown(
            f"""
            <div class="hero">

                <div class="hero-title">
                    🏗️ Construction Projects Dashboard
                </div>

                <br>

                <div class="hero-text">
                    Upload one or more Excel files.
                </div>

                <br>

                <div class="hero-text">
                    Each Excel file represents one Project.
                    The file name automatically becomes
                    the Project name.
                </div>

                <br>

                <div class="hero-text">
                    Select Project → Select PR → Apply filters
                    → View KPIs and tables.
                </div>

            </div>
            """,
            unsafe_allow_html=True,
        )

        st.info(
            "Please upload Excel project files from the sidebar."
        )

        return

    # --------------------------------------------------------
    # PROJECT FILTER
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Selection</div>',
        unsafe_allow_html=True,
    )

    project = project_filter()

    # --------------------------------------------------------
    # DATA
    # --------------------------------------------------------

    df = get_project_data()

    if df.empty:

        st.warning(
            "No data found for the selected project."
        )

        return

    # --------------------------------------------------------
    # PR FILTER
    # --------------------------------------------------------

    col1, col2 = st.columns(
        [1, 3]
    )

    with col1:

        df = pr_filter(
            df
        )

    # --------------------------------------------------------
    # PROJECT INFO
    # --------------------------------------------------------

    show_project_info(
        df
    )

    # --------------------------------------------------------
    # ADDITIONAL FILTERS
    # --------------------------------------------------------

    df = apply_dynamic_filters(
        df
    )

    # --------------------------------------------------------
    # SEARCH
    # --------------------------------------------------------

    df = search_filter(
        df
    )

    # --------------------------------------------------------
    # KPI
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">KPIs</div>',
        unsafe_allow_html=True,
    )

    show_kpis(
        df
    )

    # --------------------------------------------------------
    # PR SUMMARY
    # --------------------------------------------------------

    if (
        st.session_state.selected_pr
        == "All PRs"
    ):

        show_pr_summary(
            df
        )

    # --------------------------------------------------------
    # MAIN TABLE
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Selected Data</div>',
        unsafe_allow_html=True,
    )

    st.caption(
        f"Showing {len(df):,} records"
    )

    if df.empty:

        st.warning(
            "No records match the selected filters."
        )

    else:

        st.dataframe(
            df,
            use_container_width=True,
            height=550,
            hide_index=True,
        )

    # --------------------------------------------------------
    # DOWNLOAD
    # --------------------------------------------------------

    if not df.empty:

        csv = df.to_csv(
            index=False
        ).encode(
            "utf-8-sig"
        )

        st.download_button(
            "Download Selected Data",
            data=csv,
            file_name=(
                f"{normalize_key(str(project))}_"
                f"{normalize_key(str(st.session_state.selected_pr))}.csv"
            ),
            mime="text/csv",
            use_container_width=True,
        )


# ============================================================
# SIDEBAR
# ============================================================

def sidebar():

    with st.sidebar:

        # Logo
        if os.path.exists(
            LOGO_PATH
        ):

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

        # ----------------------------------------------------
        # MULTIPLE FILE UPLOAD
        # ----------------------------------------------------

        st.subheader(
            "Project Files"
        )

        uploaded_files = st.file_uploader(
            "Upload Project Excel Files",
            type=[
                "xlsx",
                "xls",
            ],
            accept_multiple_files=True,
            key="project_files_uploader",
            help=(
                "Each Excel file represents one project. "
                "The file name will be used as the project name."
            ),
        )

        if uploaded_files:

            register_uploaded_files(
                uploaded_files
            )

        # ----------------------------------------------------
        # REGISTERED PROJECTS
        # ----------------------------------------------------

        if st.session_state.projects:

            st.success(
                f"{len(st.session_state.projects)} "
                f"project(s) loaded"
            )

            st.markdown(
                "**Registered Projects**"
            )

            for project_name in sorted(
                st.session_state.projects.keys(),
                key=str.lower,
            ):

                st.caption(
                    f"🏗️ {project_name}"
                )

        else:

            st.info(
                "No project files loaded."
            )

        st.divider()

        # ----------------------------------------------------
        # RESET
        # ----------------------------------------------------

        if st.button(
            "🗑️ Clear All Projects",
            use_container_width=True,
        ):

            st.session_state.projects = {}

            st.session_state.selected_project = (
                "All Projects"
            )

            st.session_state.selected_pr = (
                "All PRs"
            )

            st.rerun()

        # ----------------------------------------------------
        # REFRESH
        # ----------------------------------------------------

        if st.button(
            "🔄 Refresh Dashboard",
            use_container_width=True,
        ):

            st.rerun()

        st.divider()

        st.markdown(
            """
            <div class="footer">
                Dynamic PR Dashboard<br>
                Hassan Allam Construction
            </div>
            """,
            unsafe_allow_html=True,
        )


# ============================================================
# RUN
# ============================================================

sidebar()

dashboard()
```
