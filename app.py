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
    page_title="Hassan Allam Construction",
    page_icon="🏗️",
    layout="wide"
)


# =========================================================
# CSS
# =========================================================

st.markdown("""
<style>

.main-title {
    font-size: 30px;
    font-weight: 700;
    margin-bottom: 5px;
}

.sub-title {
    color: #666;
    font-size: 15px;
    margin-bottom: 25px;
}

.sheet-button {
    margin-bottom: 10px;
}

</style>
""", unsafe_allow_html=True)


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
# FUNCTIONS
# =========================================================

def clean_column_name(column):

    if column is None:
        return ""

    column = str(column).strip()

    column = re.sub(
        r"\s+",
        " ",
        column
    )

    return column


def make_unique_columns(columns):

    result = []
    counts = {}

    for col in columns:

        col = clean_column_name(col)

        if not col:
            col = "Unnamed"

        if col not in counts:

            counts[col] = 0
            result.append(col)

        else:

            counts[col] += 1

            result.append(
                f"{col}_{counts[col]}"
            )

    return result


def project_name_from_file(filename):

    name = filename.rsplit(
        ".",
        1
    )[0]

    name = re.sub(
        r"[_\-]+",
        " ",
        name
    )

    name = re.sub(
        r"\s+",
        " ",
        name
    )

    return name.strip()


def get_file_hash(uploaded_file):

    return hashlib.md5(
        uploaded_file.getvalue()
    ).hexdigest()


# =========================================================
# READ EXCEL
# =========================================================

def read_excel_file(uploaded_file):

    file_bytes = uploaded_file.getvalue()

    excel = pd.ExcelFile(
        BytesIO(file_bytes)
    )

    sheets = {}

    for sheet_name in excel.sheet_names:

        try:

            df = pd.read_excel(
                BytesIO(file_bytes),
                sheet_name=sheet_name
            )

            # Remove completely empty rows
            df = df.dropna(
                how="all"
            )

            # Remove completely empty columns
            df = df.dropna(
                axis=1,
                how="all"
            )

            df.columns = make_unique_columns(
                df.columns
            )

            df = df.reset_index(
                drop=True
            )

            sheets[sheet_name] = df

        except Exception as e:

            st.warning(
                f"Could not read sheet "
                f"{sheet_name}: {e}"
            )

    return sheets


# =========================================================
# REGISTER PROJECT
# =========================================================

def register_file(uploaded_file):

    project_name = project_name_from_file(
        uploaded_file.name
    )

    file_hash = get_file_hash(
        uploaded_file
    )

    # Don't reload same file
    if project_name in st.session_state.projects:

        old_hash = (
            st.session_state
            .projects[project_name]
            .get("hash")
        )

        if old_hash == file_hash:
            return

    sheets = read_excel_file(
        uploaded_file
    )

    st.session_state.projects[
        project_name
    ] = {

        "file_name": uploaded_file.name,

        "hash": file_hash,

        "sheets": sheets
    }


# =========================================================
# HEADER
# =========================================================

st.markdown(
    '<div class="main-title">'
    'Hassan Allam Construction'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="sub-title">'
    'Construction Projects Dashboard'
    '</div>',
    unsafe_allow_html=True
)


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.header("📁 Projects")

    uploaded_files = st.file_uploader(
        "Upload Excel Files",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    # -----------------------------------------------------
    # Register uploaded files
    # -----------------------------------------------------

    if uploaded_files:

        for uploaded_file in uploaded_files:

            register_file(
                uploaded_file
            )


    # -----------------------------------------------------
    # Existing projects
    # -----------------------------------------------------

    project_names = sorted(
        st.session_state
        .projects
        .keys()
    )

    if project_names:

        st.write(
            f"Projects: {len(project_names)}"
        )

    else:

        st.info(
            "Upload an Excel file."
        )


    # -----------------------------------------------------
    # Clear projects
    # -----------------------------------------------------

    if st.button(
        "🗑️ Clear Projects",
        use_container_width=True
    ):

        st.session_state.projects = {}

        st.session_state.selected_project = None

        st.session_state.selected_sheet = None

        st.rerun()


# =========================================================
# NO PROJECT
# =========================================================

project_names = sorted(
    st.session_state
    .projects
    .keys()
)

if not project_names:

    st.info(
        "Please upload an Excel file to start."
    )

    st.stop()


# =========================================================
# PROJECT SELECTOR
# =========================================================

st.subheader("🏗️ Project")

if (
    st.session_state.selected_project
    not in project_names
):

    st.session_state.selected_project = (
        project_names[0]
    )


selected_project = st.selectbox(
    "Select Project",
    project_names,
    index=project_names.index(
        st.session_state.selected_project
    )
)

st.session_state.selected_project = (
    selected_project
)


# =========================================================
# GET SELECTED PROJECT
# =========================================================

project_data = (
    st.session_state
    .projects
    .get(
        selected_project
    )
)

if project_data is None:

    st.error(
        "Project not found."
    )

    st.stop()


sheets = project_data.get(
    "sheets",
    {}
)

if not sheets:

    st.warning(
        "No sheets found in this project."
    )

    st.stop()


sheet_names = list(
    sheets.keys()
)


# =========================================================
# DEFAULT SHEET
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

st.subheader(
    "📑 Project Sheets"
)

st.caption(
    "Select a sheet to display its data."
)


# 4 buttons per row

buttons_per_row = 4


for row_start in range(
    0,
    len(sheet_names),
    buttons_per_row
):

    row_sheets = sheet_names[
        row_start:
        row_start + buttons_per_row
    ]

    cols = st.columns(
        len(row_sheets)
    )

    for i, sheet_name in enumerate(
        row_sheets
    ):

        with cols[i]:

            # ------------------------------------------------
            # Highlight selected sheet
            # ------------------------------------------------

            if (
                st.session_state.selected_sheet
                == sheet_name
            ):

                label = (
                    f"✅ {sheet_name}"
                )

            else:

                label = (
                    f"📄 {sheet_name}"
                )


            # ------------------------------------------------
            # BUTTON
            # ------------------------------------------------

            if st.button(
                label,
                key=(
                    f"sheet_{row_start}_{i}"
                ),
                use_container_width=True
            ):

                # THIS is the important part:
                # Button selects its own sheet

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


# =========================================================
# GET DATA OF SELECTED SHEET ONLY
# =========================================================

selected_df = sheets.get(
    selected_sheet
)


if selected_df is None:

    st.error(
        "Could not find selected sheet data."
    )

    st.stop()


# Make a copy
df = selected_df.copy()


# =========================================================
# SHEET INFORMATION
# =========================================================

st.divider()

st.subheader(
    f"📄 {selected_sheet}"
)


info1, info2, info3 = st.columns(3)


with info1:

    st.metric(
        "Project",
        selected_project
    )


with info2:

    st.metric(
        "Sheet",
        selected_sheet
    )


with info3:

    st.metric(
        "Records",
        f"{len(df):,}"
    )


# =========================================================
# PR DETECTION
# =========================================================

def is_pr_sheet(sheet_name):

    return bool(
        re.match(
            r"^PR[-_\s]?\d+$",
            str(sheet_name).strip(),
            re.IGNORECASE
        )
    )


sheet_is_pr = is_pr_sheet(
    selected_sheet
)


# =========================================================
# PR
# =========================================================

if sheet_is_pr:

    st.info(
        f"🔎 PR: **{selected_sheet}**"
    )


# =========================================================
# KPI
# =========================================================

st.subheader(
    "📊 Sheet KPIs"
)


# Number of rows
records = len(df)


# Number of columns
columns_count = len(
    df.columns
)


# Numeric columns
numeric_columns = df.select_dtypes(
    include=np.number
).columns


numeric_count = len(
    numeric_columns
)


# Empty cells
blank_cells = int(
    df.isna().sum().sum()
)


k1, k2, k3, k4 = st.columns(4)


with k1:

    st.metric(
        "Records",
        f"{records:,}"
    )


with k2:

    st.metric(
        "Columns",
        f"{columns_count:,}"
    )


with k3:

    st.metric(
        "Numeric Columns",
        f"{numeric_count:,}"
    )


with k4:

    st.metric(
        "Blank Cells",
        f"{blank_cells:,}"
    )


# =========================================================
# SEARCH
# =========================================================

st.subheader(
    "🔍 Search"
)

search = st.text_input(
    "Search inside this sheet",
    placeholder=(
        "Type any value to search..."
    )
)


filtered_df = df.copy()


if search:

    search = search.strip()

    if search:

        mask = pd.Series(
            False,
            index=filtered_df.index
        )

        for column in filtered_df.columns:

            mask = (
                mask
                |
                filtered_df[column]
                .astype(str)
                .str.contains(
                    search,
                    case=False,
                    na=False,
                    regex=False
                )
            )

        filtered_df = filtered_df[
            mask
        ]


# =========================================================
# DATA TABLE
# =========================================================

st.subheader(
    f"📋 Data in {selected_sheet}"
)


if filtered_df.empty:

    st.warning(
        "No records found."
    )

else:

    st.dataframe(
        filtered_df,
        use_container_width=True,
        height=600,
        hide_index=True
    )


# =========================================================
# DOWNLOAD
# =========================================================

st.subheader(
    "⬇️ Export"
)

csv = filtered_df.to_csv(
    index=False
).encode(
    "utf-8-sig"
)

st.download_button(
    label="Download Selected Sheet",
    data=csv,
    file_name=(
        f"{selected_project}_"
        f"{selected_sheet}.csv"
    ),
    mime="text/csv",
    use_container_width=True
)
