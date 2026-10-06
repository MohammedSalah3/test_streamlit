import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re
import hashlib

# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Hassan Allam Construction | PR Dashboard",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# CONSTANTS
# =========================================================

COMPANY = "Hassan Allam Construction"

PRIMARY = "#C8102E"
DARK = "#1F2933"
LIGHT = "#F4F5F7"
GRAY = "#6B7280"
GREEN = "#2E9E5B"
ORANGE = "#D97706"
BORDER = "#E1E4E8"

# =========================================================
# CUSTOM CSS
# =========================================================

st.markdown(
    f"""
    <style>

    .main {{
        background-color: #FFFFFF;
    }}

    .block-container {{
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }}

    .company-title {{
        font-size: 30px;
        font-weight: 700;
        color: {DARK};
        margin-bottom: 0px;
    }}

    .company-subtitle {{
        color: {GRAY};
        font-size: 15px;
        margin-top: 3px;
        margin-bottom: 20px;
    }}

    .section-title {{
        font-size: 21px;
        font-weight: 700;
        color: {DARK};
        margin-top: 20px;
        margin-bottom: 12px;
    }}

    .kpi-card {{
        background: white;
        border: 1px solid {BORDER};
        border-radius: 12px;
        padding: 18px;
        min-height: 110px;
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
        font-size: 25px;
        font-weight: 700;
    }}

    .project-card {{
        background: #FFFFFF;
        border: 1px solid {BORDER};
        border-radius: 10px;
        padding: 12px;
        margin-bottom: 8px;
    }}

    .project-name {{
        color: {DARK};
        font-weight: 600;
        font-size: 14px;
    }}

    .info-box {{
        background: #F8F9FA;
        border-left: 4px solid {PRIMARY};
        padding: 12px 15px;
        border-radius: 6px;
        margin-bottom: 15px;
    }}

    </style>
    """,
    unsafe_allow_html=True
)

# =========================================================
# SESSION STATE
# =========================================================

if "projects" not in st.session_state:
    st.session_state.projects = {}

# =========================================================
# HELPER FUNCTIONS
# =========================================================

def clean_column_name(column):
    """
    Clean Excel column names.
    """
    if column is None:
        return ""

    column = str(column).strip()

    column = re.sub(r"\s+", " ", column)

    return column


def make_unique_columns(columns):
    """
    Make duplicate Excel column names unique.
    """

    result = []
    counter = {}

    for col in columns:

        col = clean_column_name(col)

        if not col:
            col = "Unnamed"

        if col not in counter:
            counter[col] = 0
            result.append(col)
        else:
            counter[col] += 1
            result.append(f"{col}_{counter[col]}")

    return result


def pretty_project_name(filename):
    """
    Convert filename to readable project name.
    """

    name = filename.rsplit(".", 1)[0]

    name = re.sub(r"[_\-]+", " ", name)

    name = re.sub(r"\s+", " ", name).strip()

    return name


def file_hash(uploaded_file):
    """
    Generate unique hash for uploaded file.
    """

    content = uploaded_file.getvalue()

    return hashlib.md5(content).hexdigest()


def detect_header_row(raw_df):
    """
    Try to detect the actual header row.
    """

    if raw_df.empty:
        return 0

    max_rows_to_check = min(len(raw_df), 15)

    best_row = 0
    best_score = -1

    for i in range(max_rows_to_check):

        row = raw_df.iloc[i]

        values = [
            str(x).strip()
            for x in row.tolist()
            if pd.notna(x) and str(x).strip()
        ]

        if not values:
            continue

        score = len(values)

        keywords = [
            "pr",
            "project",
            "supplier",
            "vendor",
            "item",
            "material",
            "quantity",
            "amount",
            "cost",
            "date",
            "status",
            "description"
        ]

        for value in values:
            value_lower = value.lower()

            for keyword in keywords:
                if keyword in value_lower:
                    score += 2

        if score > best_score:
            best_score = score
            best_row = i

    return best_row


def read_excel_sheet(excel_bytes, sheet_name):
    """
    Read one Excel sheet and detect its header.
    """

    try:

        raw = pd.read_excel(
            BytesIO(excel_bytes),
            sheet_name=sheet_name,
            header=None
        )

        if raw.empty:
            return pd.DataFrame()

        header_row = detect_header_row(raw)

        headers = raw.iloc[header_row].tolist()

        headers = make_unique_columns(headers)

        df = raw.iloc[header_row + 1:].copy()

        df.columns = headers

        df = df.dropna(how="all")

        df = df.reset_index(drop=True)

        # Remove columns where everything is empty
        df = df.dropna(axis=1, how="all")

        # Clean columns
        df.columns = make_unique_columns(df.columns)

        return df

    except Exception:
        return pd.DataFrame()


def load_excel_file(excel_bytes):
    """
    Read all sheets from one Excel file
    and combine them into one project dataset.
    """

    try:

        excel = pd.ExcelFile(BytesIO(excel_bytes))

        all_data = []

        for sheet_name in excel.sheet_names:

            df = read_excel_sheet(
                excel_bytes,
                sheet_name
            )

            if df.empty:
                continue

            df["_Source Sheet"] = sheet_name

            all_data.append(df)

        if not all_data:
            return pd.DataFrame()

        combined = pd.concat(
            all_data,
            ignore_index=True,
            sort=False
        )

        combined = combined.dropna(
            how="all"
        )

        combined = combined.reset_index(drop=True)

        return combined

    except Exception as e:

        st.error(
            f"Error reading Excel file: {str(e)}"
        )

        return pd.DataFrame()


def find_column(df, candidates):
    """
    Find a column using exact and partial matching.
    """

    if df is None or df.empty:
        return None

    columns = list(df.columns)

    # Exact match
    for candidate in candidates:

        for col in columns:

            if str(col).strip().lower() == candidate.lower():
                return col

    # Partial match
    for candidate in candidates:

        candidate_lower = candidate.lower()

        for col in columns:

            col_lower = str(col).strip().lower()

            if candidate_lower in col_lower:
                return col

    return None


def find_pr_column(df):
    """
    Detect PR column.
    """

    if df is None or df.empty:
        return None

    columns = list(df.columns)

    # Strong exact matches first
    exact_candidates = [
        "PR",
        "PR Number",
        "PR No",
        "PR No.",
        "PR#",
        "Purchase Requisition",
        "Purchase Requisition Number",
        "Purchase Requisition No",
        "Requisition Number",
        "Requisition No"
    ]

    for candidate in exact_candidates:

        for col in columns:

            if str(col).strip().lower() == candidate.lower():
                return col

    # Regex / partial detection
    for col in columns:

        col_text = str(col).strip().lower()

        if re.fullmatch(r"pr", col_text):
            return col

        if re.search(r"\bpr\s*(number|no|no\.)\b", col_text):
            return col

        if "purchase requisition" in col_text:
            return col

        if "requisition number" in col_text:
            return col

        if "requisition no" in col_text:
            return col

    return None


def numeric_series(df, column):
    """
    Convert column to numeric safely.
    """

    if column is None or column not in df.columns:
        return pd.Series(dtype=float)

    series = df[column]

    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(
            series,
            errors="coerce"
        )

    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("EGP", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace("€", "", regex=False)
        .str.strip()
    )

    return pd.to_numeric(
        cleaned,
        errors="coerce"
    )


def find_numeric_column(df, candidates):
    """
    Find likely financial/numeric column.
    """

    if df.empty:
        return None

    columns = list(df.columns)

    # Exact first
    for candidate in candidates:

        for col in columns:

            if str(col).strip().lower() == candidate.lower():
                return col

    # Partial
    for candidate in candidates:

        candidate_lower = candidate.lower()

        for col in columns:

            col_lower = str(col).lower()

            if candidate_lower in col_lower:

                test = numeric_series(df, col)

                if test.notna().sum() > 0:
                    return col

    return None


def format_number(value):
    """
    Format KPI numbers.
    """

    if value is None:
        return "0"

    try:

        if pd.isna(value):
            return "0"

        value = float(value)

        if value.is_integer():
            return f"{int(value):,}"

        return f"{value:,.2f}"

    except Exception:
        return str(value)


def format_currency(value):
    """
    Format financial KPI.
    """

    if value is None:
        return "0"

    try:

        if pd.isna(value):
            return "0"

        value = float(value)

        return f"{value:,.2f}"

    except Exception:
        return str(value)


def show_kpi(title, value):
    """
    Render KPI card.
    """

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">{title}</div>
            <div class="kpi-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def clean_dataframe_for_display(df):
    """
    Prepare dataframe for display.
    """

    result = df.copy()

    for col in result.columns:

        if pd.api.types.is_datetime64_any_dtype(
            result[col]
        ):
            result[col] = result[col].dt.strftime(
                "%Y-%m-%d"
            )

    return result


def register_uploaded_files(uploaded_files):
    """
    Register uploaded Excel files as projects.
    """

    if not uploaded_files:
        return

    for uploaded_file in uploaded_files:

        try:

            file_bytes = uploaded_file.getvalue()

            if not file_bytes:
                continue

            project_name = pretty_project_name(
                uploaded_file.name
            )

            unique_hash = file_hash(
                uploaded_file
            )

            # If same file already exists, don't reload
            existing = st.session_state.projects.get(
                project_name
            )

            if (
                existing
                and existing.get("hash") == unique_hash
            ):
                continue

            data = load_excel_file(
                file_bytes
            )

            st.session_state.projects[
                project_name
            ] = {
                "file_name": uploaded_file.name,
                "file_bytes": file_bytes,
                "hash": unique_hash,
                "data": data
            }

        except Exception as e:

            st.error(
                f"Could not load {uploaded_file.name}: {e}"
            )


def get_dynamic_filter_columns(df, excluded_columns):
    """
    Detect useful categorical columns for filters.
    """

    filters = []

    priority_words = [
        "supplier",
        "vendor",
        "status",
        "category",
        "department",
        "section",
        "location",
        "item",
        "material",
        "type",
        "discipline",
        "company"
    ]

    for col in df.columns:

        if col in excluded_columns:
            continue

        series = df[col]

        if pd.api.types.is_numeric_dtype(series):
            continue

        non_null = series.dropna()

        if non_null.empty:
            continue

        unique_count = non_null.astype(str).nunique()

        # Only useful categorical columns
        if unique_count <= 40:

            col_lower = str(col).lower()

            priority = any(
                word in col_lower
                for word in priority_words
            )

            if priority:
                filters.insert(0, col)
            else:
                filters.append(col)

    # Remove duplicates
    result = []

    for col in filters:

        if col not in result:
            result.append(col)

    return result[:12]


# =========================================================
# HEADER
# =========================================================

st.markdown(
    f"""
    <div class="company-title">
        {COMPANY}
    </div>
    <div class="company-subtitle">
        Construction Projects & Purchase Requisitions Dashboard
    </div>
    """,
    unsafe_allow_html=True
)

# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        "## 📁 Project Files"
    )

    st.caption(
        "Each Excel file represents one project."
    )

    uploaded_files = st.file_uploader(
        "Upload Excel Project Files",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="project_uploader"
    )

    if uploaded_files:

        register_uploaded_files(
            uploaded_files
        )

    st.divider()

    if st.session_state.projects:

        st.markdown(
            f"**Registered Projects: {len(st.session_state.projects)}**"
        )

        for project_name in sorted(
            st.session_state.projects.keys()
        ):

            data = st.session_state.projects[
                project_name
            ]["data"]

            rows = len(data)

            st.markdown(
                f"""
                <div class="project-card">
                    <div class="project-name">
                        🏗️ {project_name}
                    </div>
                    <div style="font-size:12px;color:{GRAY};">
                        {rows:,} records
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

    else:

        st.info(
            "Upload one or more Excel files to start."
        )

    st.divider()

    if st.button(
        "🗑️ Clear All Projects",
        use_container_width=True
    ):

        st.session_state.projects = {}

        st.rerun()

    if st.button(
        "🔄 Refresh Dashboard",
        use_container_width=True
    ):

        st.rerun()

# =========================================================
# NO PROJECT
# =========================================================

if not st.session_state.projects:

    st.markdown(
        """
        <div class="info-box">
            <b>Welcome.</b><br>
            Upload Excel files from the sidebar.
            Each Excel file will automatically become a Project.
        </div>
        """,
        unsafe_allow_html=True
    )

    st.stop()

# =========================================================
# PROJECT FILTER
# =========================================================

st.markdown(
    '<div class="section-title">🔎 Filters</div>',
    unsafe_allow_html=True
)

project_names = sorted(
    st.session_state.projects.keys()
)

project_options = [
    "All Projects"
] + project_names

selected_project = st.selectbox(
    "Project",
    project_options,
    key="project_filter"
)

# =========================================================
# COMBINE SELECTED PROJECTS
# =========================================================

if selected_project == "All Projects":

    project_frames = []

    for project_name in project_names:

        project_data = st.session_state.projects[
            project_name
        ]["data"].copy()

        if project_data.empty:
            continue

        project_data.insert(
            0,
            "Project",
            project_name
        )

        project_frames.append(
            project_data
        )

    if project_frames:

        df = pd.concat(
            project_frames,
            ignore_index=True,
            sort=False
        )

    else:

        df = pd.DataFrame()

else:

    df = st.session_state.projects[
        selected_project
    ]["data"].copy()

    if not df.empty:

        df.insert(
            0,
            "Project",
            selected_project
        )

# =========================================================
# CHECK DATA
# =========================================================

if df.empty:

    st.warning(
        "The selected project does not contain readable data."
    )

    st.stop()

# =========================================================
# PR DETECTION
# =========================================================

pr_column = find_pr_column(
    df
)

# =========================================================
# PR FILTER
# =========================================================

selected_pr = "All PRs"

if pr_column:

    pr_values = (
        df[pr_column]
        .dropna()
        .astype(str)
        .str.strip()
    )

    pr_values = pr_values[
        pr_values != ""
    ]

    unique_prs = sorted(
        pr_values.unique().tolist()
    )

    pr_options = [
        "All PRs"
    ] + unique_prs

    selected_pr = st.selectbox(
        f"PR ({pr_column})",
        pr_options,
        key="pr_filter"
    )

else:

    st.info(
        "No PR column was detected in this project."
    )

# =========================================================
# APPLY PR FILTER
# =========================================================

filtered_df = df.copy()

if (
    pr_column
    and selected_pr != "All PRs"
):

    filtered_df = filtered_df[
        filtered_df[pr_column]
        .astype(str)
        .str.strip()
        == selected_pr
    ].copy()

# =========================================================
# DYNAMIC FILTERS
# =========================================================

excluded_columns = {
    "Project",
    "_Source Sheet"
}

if pr_column:
    excluded_columns.add(
        pr_column
    )

dynamic_filters = get_dynamic_filter_columns(
    filtered_df,
    excluded_columns
)

if dynamic_filters:

    st.markdown(
        "### Additional Filters"
    )

    filter_columns = st.columns(
        min(4, len(dynamic_filters))
    )

    for index, column in enumerate(
        dynamic_filters
    ):

        unique_values = (
            filtered_df[column]
            .dropna()
            .astype(str)
            .str.strip()
        )

        unique_values = unique_values[
            unique_values != ""
        ]

        values = sorted(
            unique_values.unique().tolist()
        )

        if not values:
            continue

        with filter_columns[
            index % len(filter_columns)
        ]:

            selected_value = st.selectbox(
                column,
                ["All"] + values,
                key=f"dynamic_filter_{column}"
            )

            if selected_value != "All":

                filtered_df = filtered_df[
                    filtered_df[column]
                    .astype(str)
                    .str.strip()
                    == selected_value
                ].copy()

# =========================================================
# SEARCH
# =========================================================

search_text = st.text_input(
    "🔍 Search",
    placeholder="Search in the displayed data..."
)

if search_text:

    search_text = search_text.strip()

    if search_text:

        mask = pd.Series(
            False,
            index=filtered_df.index
        )

        for col in filtered_df.columns:

            try:

                mask = (
                    mask
                    |
                    filtered_df[col]
                    .astype(str)
                    .str.contains(
                        search_text,
                        case=False,
                        na=False,
                        regex=False
                    )
                )

            except Exception:
                pass

        filtered_df = filtered_df[
            mask
        ].copy()

# =========================================================
# KPI COLUMN DETECTION
# =========================================================

budget_column = find_numeric_column(
    filtered_df,
    [
        "Budget",
        "Budget Amount",
        "Original Budget",
        "Approved Budget",
        "Estimated Budget"
    ]
)

planned_column = find_numeric_column(
    filtered_df,
    [
        "Planned",
        "Planned Amount",
        "Planned Cost",
        "Planned Value"
    ]
)

actual_column = find_numeric_column(
    filtered_df,
    [
        "Actual",
        "Actual Amount",
        "Actual Cost",
        "Actual Value",
        "Spent",
        "Spend"
    ]
)

quantity_column = find_numeric_column(
    filtered_df,
    [
        "Quantity",
        "Qty",
        "Required Quantity",
        "Requested Quantity"
    ]
)

# =========================================================
# KPI CALCULATIONS
# =========================================================

records_count = len(filtered_df)

if pr_column:

    prs_count = (
        filtered_df[pr_column]
        .dropna()
        .astype(str)
        .str.strip()
        .replace("", np.nan)
        .nunique()
    )

else:

    prs_count = 0

budget_value = 0
planned_value = 0
actual_value = 0
quantity_value = 0

if budget_column:

    budget_value = numeric_series(
        filtered_df,
        budget_column
    ).sum()

if planned_column:

    planned_value = numeric_series(
        filtered_df,
        planned_column
    ).sum()

if actual_column:

    actual_value = numeric_series(
        filtered_df,
        actual_column
    ).sum()

if quantity_column:

    quantity_value = numeric_series(
        filtered_df,
        quantity_column
    ).sum()

# Remaining
remaining_value = 0

if budget_column and actual_column:

    remaining_value = (
        budget_value
        - actual_value
    )

elif planned_column and actual_column:

    remaining_value = (
        planned_value
        - actual_value
    )

# =========================================================
# KPI SECTION
# =========================================================

st.markdown(
    '<div class="section-title">📊 KPIs</div>',
    unsafe_allow_html=True
)

kpi_columns = st.columns(7)

with kpi_columns[0]:

    show_kpi(
        "Records",
        format_number(records_count)
    )

with kpi_columns[1]:

    show_kpi(
        "PRs",
        format_number(prs_count)
    )

with kpi_columns[2]:

    show_kpi(
        "Budget",
        format_currency(budget_value)
    )

with kpi_columns[3]:

    show_kpi(
        "Planned",
        format_currency(planned_value)
    )

with kpi_columns[4]:

    show_kpi(
        "Actual",
        format_currency(actual_value)
    )

with kpi_columns[5]:

    show_kpi(
        "Remaining",
        format_currency(remaining_value)
    )

with kpi_columns[6]:

    show_kpi(
        "Quantity",
        format_number(quantity_value)
    )

# =========================================================
# CURRENT FILTER SUMMARY
# =========================================================

st.markdown(
    '<div class="section-title">📌 Current Selection</div>',
    unsafe_allow_html=True
)

summary_col1, summary_col2, summary_col3 = st.columns(3)

with summary_col1:

    if selected_project == "All Projects":
        st.metric(
            "Project",
            "All Projects"
        )
    else:
        st.metric(
            "Project",
            selected_project
        )

with summary_col2:

    st.metric(
        "PR",
        selected_pr
    )

with summary_col3:

    st.metric(
        "Displayed Records",
        f"{len(filtered_df):,}"
    )

# =========================================================
# PR SUMMARY
# =========================================================

if pr_column and selected_pr == "All PRs":

    st.markdown(
        '<div class="section-title">📋 PR Summary</div>',
        unsafe_allow_html=True
    )

    summary_df = filtered_df.copy()

    summary_df["_PR_Display"] = (
        summary_df[pr_column]
        .astype(str)
        .str.strip()
    )

    summary_df = summary_df[
        summary_df["_PR_Display"]
        .notna()
    ]

    group_columns = [
        "_PR_Display"
    ]

    aggregation = {}

    if budget_column:
        aggregation[
            budget_column
        ] = "sum"

    if planned_column:
        aggregation[
            planned_column
        ] = "sum"

    if actual_column:
        aggregation[
            actual_column
        ] = "sum"

    if quantity_column:
        aggregation[
            quantity_column
        ] = "sum"

    aggregation[
        "_PR_Display"
    ] = "size"

    if aggregation:

        pr_summary = (
            summary_df
            .groupby("_PR_Display")
            .agg(aggregation)
            .reset_index()
        )

        # Remove duplicate column generated by aggregation
        if "_PR_Display" in pr_summary.columns:

            pr_summary = pr_summary.rename(
                columns={
                    "_PR_Display": "PR"
                }
            )

        if "_PR_Display" in pr_summary.columns:

            pr_summary = pr_summary.drop(
                columns=["_PR_Display"]
            )

        # Rename count if generated
        if "_PR_Display" in pr_summary.columns:
            pr_summary = pr_summary.rename(
                columns={
                    "_PR_Display": "Records"
                }
            )

        # Better clean summary
        pr_summary = pr_summary.rename(
            columns={
                budget_column: "Budget"
                if budget_column
                else budget_column,
                planned_column: "Planned"
                if planned_column
                else planned_column,
                actual_column: "Actual"
                if actual_column
                else actual_column,
                quantity_column: "Quantity"
                if quantity_column
                else quantity_column
            }
        )

        # Recalculate PR records
        record_counts = (
            summary_df
            .groupby("_PR_Display")
            .size()
            .reset_index(
                name="Records"
            )
            .rename(
                columns={
                    "_PR_Display": "PR"
                }
            )
        )

        pr_summary = (
            pr_summary
            .drop(
                columns=["Records"],
                errors="ignore"
            )
            .merge(
                record_counts,
                on="PR",
                how="left"
            )
        )

        st.dataframe(
            clean_dataframe_for_display(
                pr_summary
            ),
            use_container_width=True,
            hide_index=True
        )

# =========================================================
# MAIN DATA TABLE
# =========================================================

st.markdown(
    '<div class="section-title">📄 Project Data</div>',
    unsafe_allow_html=True
)

display_df = clean_dataframe_for_display(
    filtered_df
)

st.dataframe(
    display_df,
    use_container_width=True,
    height=550,
    hide_index=True
)

# =========================================================
# DOWNLOAD
# =========================================================

st.markdown(
    '<div class="section-title">⬇️ Export</div>',
    unsafe_allow_html=True
)

csv_data = filtered_df.to_csv(
    index=False
).encode(
    "utf-8-sig"
)

st.download_button(
    label="⬇️ Download Current Data as CSV",
    data=csv_data,
    file_name="filtered_project_data.csv",
    mime="text/csv",
    use_container_width=False
)

# =========================================================
# FOOTER
# =========================================================

st.markdown(
    f"""
    <div style="
        margin-top:30px;
        padding-top:15px;
        border-top:1px solid {BORDER};
        text-align:center;
        color:{GRAY};
        font-size:12px;
    ">
        {COMPANY} | Construction Projects & PR Dashboard
    </div>
    """,
    unsafe_allow_html=True
)
