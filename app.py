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
    page_title="Hassan Allam Construction | Projects Dashboard",
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
        min-height: 105px;
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

    </style>
    """,
    unsafe_allow_html=True
)

# =========================================================
# SESSION STATE
# =========================================================

if "projects" not in st.session_state:
    st.session_state.projects = {}

if "selected_project" not in st.session_state:
    st.session_state.selected_project = None

if "selected_sheet" not in st.session_state:
    st.session_state.selected_sheet = None


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def clean_column_name(column):
    if column is None:
        return ""

    column = str(column).strip()
    column = re.sub(r"\s+", " ", column)

    return column


def make_unique_columns(columns):

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

            result.append(
                f"{col}_{counter[col]}"
            )

    return result


def pretty_project_name(filename):

    name = filename.rsplit(".", 1)[0]

    name = re.sub(
        r"[_\-]+",
        " ",
        name
    )

    name = re.sub(
        r"\s+",
        " ",
        name
    ).strip()

    return name


def file_hash(uploaded_file):

    content = uploaded_file.getvalue()

    return hashlib.md5(content).hexdigest()


def detect_header_row(raw_df):

    if raw_df.empty:
        return 0

    max_rows_to_check = min(
        len(raw_df),
        15
    )

    best_row = 0
    best_score = -1

    for i in range(max_rows_to_check):

        row = raw_df.iloc[i]

        values = [
            str(x).strip()
            for x in row.tolist()
            if pd.notna(x)
            and str(x).strip()
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


def read_excel_sheet(
    excel_bytes,
    sheet_name
):

    try:

        raw = pd.read_excel(
            BytesIO(excel_bytes),
            sheet_name=sheet_name,
            header=None
        )

        if raw.empty:
            return pd.DataFrame()

        header_row = detect_header_row(
            raw
        )

        headers = raw.iloc[
            header_row
        ].tolist()

        headers = make_unique_columns(
            headers
        )

        df = raw.iloc[
            header_row + 1:
        ].copy()

        df.columns = headers

        df = df.dropna(
            how="all"
        )

        df = df.dropna(
            axis=1,
            how="all"
        )

        df = df.reset_index(
            drop=True
        )

        df.columns = make_unique_columns(
            df.columns
        )

        return df

    except Exception:

        return pd.DataFrame()


def load_excel_file(excel_bytes):

    sheets = {}

    try:

        excel = pd.ExcelFile(
            BytesIO(excel_bytes)
        )

        for sheet_name in excel.sheet_names:

            df = read_excel_sheet(
                excel_bytes,
                sheet_name
            )

            if not df.empty:

                sheets[sheet_name] = df

    except Exception as e:

        st.error(
            f"Error reading Excel file: {e}"
        )

    return sheets


def find_pr_column(df):

    if df is None or df.empty:
        return None

    columns = list(df.columns)

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

    # Exact match first
    for candidate in exact_candidates:

        for col in columns:

            if (
                str(col)
                .strip()
                .lower()
                ==
                candidate.lower()
            ):

                return col

    # Additional matching
    for col in columns:

        col_text = (
            str(col)
            .strip()
            .lower()
        )

        if re.fullmatch(
            r"pr",
            col_text
        ):
            return col

        if re.search(
            r"\bpr\s*(number|no|no\.)\b",
            col_text
        ):
            return col

        if "purchase requisition" in col_text:
            return col

        if "requisition number" in col_text:
            return col

        if "requisition no" in col_text:
            return col

    return None


def numeric_series(
    df,
    column
):

    if (
        column is None
        or column not in df.columns
    ):

        return pd.Series(
            dtype=float
        )

    series = df[column]

    if pd.api.types.is_numeric_dtype(
        series
    ):

        return pd.to_numeric(
            series,
            errors="coerce"
        )

    cleaned = (
        series
        .astype(str)
        .str.replace(
            ",",
            "",
            regex=False
        )
        .str.replace(
            "EGP",
            "",
            regex=False
        )
        .str.replace(
            "$",
            "",
            regex=False
        )
        .str.replace(
            "€",
            "",
            regex=False
        )
        .str.strip()
    )

    return pd.to_numeric(
        cleaned,
        errors="coerce"
    )


def find_numeric_column(
    df,
    candidates
):

    if df.empty:
        return None

    columns = list(df.columns)

    # Exact match
    for candidate in candidates:

        for col in columns:

            if (
                str(col)
                .strip()
                .lower()
                ==
                candidate.lower()
            ):

                test = numeric_series(
                    df,
                    col
                )

                if test.notna().sum() > 0:
                    return col

    # Partial match
    for candidate in candidates:

        candidate_lower = candidate.lower()

        for col in columns:

            col_lower = str(col).lower()

            if candidate_lower in col_lower:

                test = numeric_series(
                    df,
                    col
                )

                if test.notna().sum() > 0:
                    return col

    return None


def format_number(value):

    try:

        if pd.isna(value):
            return "0"

        value = float(value)

        if value.is_integer():
            return f"{int(value):,}"

        return f"{value:,.2f}"

    except Exception:

        return "0"


def format_currency(value):

    try:

        if pd.isna(value):
            return "0"

        return f"{float(value):,.2f}"

    except Exception:

        return "0"


def show_kpi(
    title,
    value
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

        </div>
        """,
        unsafe_allow_html=True
    )


def clean_dataframe_for_display(df):

    result = df.copy()

    for col in result.columns:

        if pd.api.types.is_datetime64_any_dtype(
            result[col]
        ):

            result[col] = (
                result[col]
                .dt.strftime(
                    "%Y-%m-%d"
                )
            )

    return result


def register_uploaded_files(
    uploaded_files
):

    if not uploaded_files:
        return

    for uploaded_file in uploaded_files:

        try:

            file_bytes = (
                uploaded_file.getvalue()
            )

            if not file_bytes:
                continue

            project_name = (
                pretty_project_name(
                    uploaded_file.name
                )
            )

            unique_hash = file_hash(
                uploaded_file
            )

            existing = (
                st.session_state.projects.get(
                    project_name
                )
            )

            # Same file already loaded
            if (
                existing
                and existing.get("hash")
                == unique_hash
            ):

                continue

            sheets = load_excel_file(
                file_bytes
            )

            st.session_state.projects[
                project_name
            ] = {

                "file_name":
                    uploaded_file.name,

                "file_bytes":
                    file_bytes,

                "hash":
                    unique_hash,

                "sheets":
                    sheets
            }

        except Exception as e:

            st.error(
                f"Could not load "
                f"{uploaded_file.name}: {e}"
            )


# =========================================================
# HEADER
# =========================================================

st.markdown(
    f"""
    <div class="company-title">
        {COMPANY}
    </div>

    <div class="company-subtitle">
        Construction Projects Dashboard
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
        "Upload Excel Files",
        type=[
            "xlsx",
            "xls"
        ],
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
            f"**Projects: "
            f"{len(st.session_state.projects)}**"
        )

        for project_name in sorted(
            st.session_state.projects.keys()
        ):

            project_info = (
                st.session_state.projects[
                    project_name
                ]
            )

            sheet_count = len(
                project_info.get(
                    "sheets",
                    {}
                )
            )

            st.markdown(
                f"""
                <div style="
                    padding:8px;
                    border-bottom:
                    1px solid {BORDER};
                ">

                <b>🏗️ {project_name}</b>

                <br>

                <span style="
                    color:{GRAY};
                    font-size:12px;
                ">
                    {sheet_count} Sheets
                </span>

                </div>
                """,
                unsafe_allow_html=True
            )

    else:

        st.info(
            "Upload Excel files."
        )

    st.divider()

    if st.button(
        "🗑️ Clear All Projects",
        use_container_width=True
    ):

        st.session_state.projects = {}

        st.session_state.selected_project = None

        st.session_state.selected_sheet = None

        st.rerun()

    if st.button(
        "🔄 Refresh Dashboard",
        use_container_width=True
    ):

        st.rerun()


# =========================================================
# CHECK PROJECTS
# =========================================================

project_names = sorted(
    st.session_state.projects.keys()
)

if not project_names:

    st.info(
        "Please upload at least one Excel file."
    )

    st.stop()


# =========================================================
# SAFE PROJECT SELECTION
# =========================================================

if (
    st.session_state.selected_project
    not in project_names
):

    st.session_state.selected_project = (
        project_names[0]
    )

st.markdown(
    '<div class="section-title">🏗️ Select Project</div>',
    unsafe_allow_html=True
)

selected_project = st.selectbox(
    "Project",
    project_names,
    index=project_names.index(
        st.session_state.selected_project
    ),
    key="project_selector"
)

st.session_state.selected_project = (
    selected_project
)


# =========================================================
# SAFE PROJECT ACCESS
# =========================================================

if (
    selected_project
    not in st.session_state.projects
):

    st.warning(
        "Selected project is not available."
    )

    st.session_state.selected_project = (
        project_names[0]
    )

    st.rerun()


project_info = st.session_state.projects.get(
    selected_project
)

if not project_info:

    st.warning(
        "Project data could not be loaded."
    )

    st.stop()


# =========================================================
# GET SHEETS
# =========================================================

project_sheets = project_info.get(
    "sheets",
    {}
)

if not project_sheets:

    st.warning(
        "No readable sheets were found "
        "in this project."
    )

    st.stop()


sheet_names = list(
    project_sheets.keys()
)


# =========================================================
# SAFE SHEET SELECTION
# =========================================================

if (
    st.session_state.selected_sheet
    not in sheet_names
):

    st.session_state.selected_sheet = (
        sheet_names[0]
    )


# =========================================================
# SHEET BUTTONS
# =========================================================

st.markdown(
    '<div class="section-title">📑 Project Sheets</div>',
    unsafe_allow_html=True
)

buttons_per_row = 4

for start in range(
    0,
    len(sheet_names),
    buttons_per_row
):

    row_sheets = sheet_names[
        start:
        start + buttons_per_row
    ]

    columns = st.columns(
        len(row_sheets)
    )

    for index, sheet_name in enumerate(
        row_sheets
    ):

        with columns[index]:

            is_selected = (
                st.session_state.selected_sheet
                == sheet_name
            )

            if is_selected:

                label = (
                    f"✅ {sheet_name}"
                )

            else:

                label = (
                    f"📄 {sheet_name}"
                )

            if st.button(
                label,
                key=(
                    "sheet_button_"
                    + str(hash(selected_project))
                    + "_"
                    + str(hash(sheet_name))
                ),
                use_container_width=True
            ):

                st.session_state.selected_sheet = (
                    sheet_name
                )

                st.rerun()


# =========================================================
# SELECTED SHEET
# =========================================================

selected_sheet = (
    st.session_state.selected_sheet
)

if selected_sheet not in project_sheets:

    selected_sheet = sheet_names[0]

    st.session_state.selected_sheet = (
        selected_sheet
    )


df = project_sheets[
    selected_sheet
].copy()


# =========================================================
# SHEET INFO
# =========================================================

st.markdown(
    f"""
    <div style="
        background:{LIGHT};
        padding:12px 15px;
        border-radius:8px;
        margin-top:15px;
        margin-bottom:15px;
    ">

        <b>Project:</b>
        {selected_project}

        &nbsp;&nbsp; | &nbsp;&nbsp;

        <b>Sheet:</b>
        {selected_sheet}

        &nbsp;&nbsp; | &nbsp;&nbsp;

        <b>Records:</b>
        {len(df):,}

    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# PR FILTER
# =========================================================

pr_column = find_pr_column(
    df
)

selected_pr = "All PRs"

if pr_column:

    st.markdown(
        '<div class="section-title">🔎 PR Filter</div>',
        unsafe_allow_html=True
    )

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
        key=(
            "pr_filter_"
            + str(hash(selected_project))
            + "_"
            + str(hash(selected_sheet))
        )
    )

else:

    st.info(
        "This sheet does not contain a PR column."
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
        filtered_df[
            pr_column
        ]
        .astype(str)
        .str.strip()
        ==
        selected_pr
    ].copy()


# =========================================================
# SEARCH
# =========================================================

search_text = st.text_input(
    "🔍 Search",
    placeholder=(
        "Search in the selected sheet..."
    ),
    key=(
        "search_"
        + str(hash(selected_project))
        + "_"
        + str(hash(selected_sheet))
    )
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
# NUMERIC COLUMN DETECTION
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

records_count = len(
    filtered_df
)

if pr_column:

    prs_count = (
        filtered_df[
            pr_column
        ]
        .dropna()
        .astype(str)
        .str.strip()
        .replace(
            "",
            np.nan
        )
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


remaining_value = 0

if (
    budget_column
    and actual_column
):

    remaining_value = (
        budget_value
        - actual_value
    )

elif (
    planned_column
    and actual_column
):

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

kpi_columns = st.columns(6)

with kpi_columns[0]:

    show_kpi(
        "Records",
        format_number(
            records_count
        )
    )

with kpi_columns[1]:

    show_kpi(
        "PRs",
        format_number(
            prs_count
        )
    )

with kpi_columns[2]:

    show_kpi(
        "Budget",
        format_currency(
            budget_value
        )
    )

with kpi_columns[3]:

    show_kpi(
        "Planned",
        format_currency(
            planned_value
        )
    )

with kpi_columns[4]:

    show_kpi(
        "Actual",
        format_currency(
            actual_value
        )
    )

with kpi_columns[5]:

    show_kpi(
        "Quantity",
        format_number(
            quantity_value
        )
    )


# =========================================================
# CURRENT SELECTION
# =========================================================

st.markdown(
    '<div class="section-title">📌 Current Selection</div>',
    unsafe_allow_html=True
)

col1, col2, col3 = st.columns(3)

with col1:

    st.metric(
        "Project",
        selected_project
    )

with col2:

    st.metric(
        "Sheet",
        selected_sheet
    )

with col3:

    st.metric(
        "PR",
        selected_pr
    )


# =========================================================
# PR SUMMARY
# =========================================================

if (
    pr_column
    and selected_pr == "All PRs"
    and not filtered_df.empty
):

    st.markdown(
        '<div class="section-title">📋 PR Summary</div>',
        unsafe_allow_html=True
    )

    summary = (
        filtered_df
        .groupby(
            pr_column,
            dropna=False
        )
        .size()
        .reset_index(
            name="Records"
        )
    )

    summary = summary.rename(
        columns={
            pr_column: "PR"
        }
    )

    st.dataframe(
        summary,
        use_container_width=True,
        hide_index=True
    )


# =========================================================
# MAIN TABLE
# =========================================================

st.markdown(
    f"""
    <div class="section-title">
        📄 {selected_sheet} Data
    </div>
    """,
    unsafe_allow_html=True
)

if filtered_df.empty:

    st.warning(
        "No data matches the current filters."
    )

else:

    display_df = (
        clean_dataframe_for_display(
            filtered_df
        )
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
    label="⬇️ Download Current Sheet",
    data=csv_data,
    file_name=(
        f"{selected_project}_"
        f"{selected_sheet}.csv"
    ),
    mime="text/csv"
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
        {COMPANY} |
        Construction Projects Dashboard
    </div>
    """,
    unsafe_allow_html=True
)
