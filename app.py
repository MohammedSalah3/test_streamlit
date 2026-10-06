import os
from pathlib import Path
from io import BytesIO

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


# =========================================================
# CONFIGURATION
# =========================================================

DATA_FOLDER = Path("data")


# =========================================================
# CUSTOM CSS
# =========================================================

st.markdown(
    """
    <style>
        .main {
            padding-top: 1rem;
        }

        .project-title {
            font-size: 30px;
            font-weight: 700;
            margin-bottom: 5px;
        }

        .project-subtitle {
            color: #6b7280;
            font-size: 15px;
            margin-bottom: 20px;
        }

        div[data-testid="stMetric"] {
            border: 1px solid #e5e7eb;
            border-radius: 10px;
            padding: 12px;
            background-color: #ffffff;
        }

        .sheet-title {
            font-size: 22px;
            font-weight: 700;
            margin-top: 10px;
            margin-bottom: 10px;
        }

        .file-info {
            color: #6b7280;
            font-size: 14px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# FUNCTIONS
# =========================================================

@st.cache_data(show_spinner=False)
def get_excel_files():
    """
    Get all Excel files from the data folder.
    """
    if not DATA_FOLDER.exists():
        return []

    files = []

    for file in DATA_FOLDER.iterdir():
        if file.is_file() and file.suffix.lower() in [".xlsx", ".xls"]:
            files.append(file)

    return sorted(files, key=lambda x: x.name.lower())


@st.cache_data(show_spinner=False)
def load_excel_file(file_path):
    """
    Read all sheets from an Excel file.
    Returns:
        dict: {sheet_name: dataframe}
    """

    try:
        excel = pd.ExcelFile(file_path)

        sheets = {}

        for sheet_name in excel.sheet_names:
            try:
                df = pd.read_excel(
                    file_path,
                    sheet_name=sheet_name
                )

                sheets[sheet_name] = df

            except Exception:
                # If a sheet cannot be read, keep an empty dataframe
                sheets[sheet_name] = pd.DataFrame()

        return sheets

    except Exception as e:
        st.error(f"Error reading file: {file_path.name}")
        st.error(str(e))
        return {}


def clean_dataframe(df):
    """
    Basic cleaning for displaying data.
    """

    if df is None:
        return pd.DataFrame()

    df = df.copy()

    # Remove completely empty rows
    df = df.dropna(how="all")

    # Remove completely empty columns
    df = df.dropna(axis=1, how="all")

    return df


def filter_dataframe(df, search_text):
    """
    Search all columns in the selected dataframe.
    """

    if not search_text:
        return df

    search_text = search_text.lower()

    mask = df.astype(str).apply(
        lambda column: column.str.lower().str.contains(
            search_text,
            na=False
        )
    ).any(axis=1)

    return df[mask]


def safe_key(text):
    """
    Create a safe Streamlit widget key.
    """

    return (
        str(text)
        .replace(" ", "_")
        .replace("/", "_")
        .replace("\\", "_")
        .replace(".", "_")
        .replace("-", "_")
    )


# =========================================================
# LOAD PROJECTS
# =========================================================

excel_files = get_excel_files()


# =========================================================
# CHECK DATA FOLDER
# =========================================================

if not DATA_FOLDER.exists():

    st.error("The 'data' folder does not exist.")

    st.info(
        """
        Create a folder named `data` in the same repository as `app.py`.

        Example:

        data/
        ├── Project1.xlsx
        ├── Project2.xlsx
        └── Project3.xlsx
        """
    )

    st.stop()


if not excel_files:

    st.warning("No Excel files were found inside the `data` folder.")

    st.info(
        """
        Upload your Excel files to the GitHub repository inside:

        `data/`

        Supported formats:

        - `.xlsx`
        - `.xls`
        """
    )

    st.stop()


# =========================================================
# CREATE PROJECT LIST
# =========================================================

project_names = [
    file.stem
    for file in excel_files
]


project_files = {
    file.stem: file
    for file in excel_files
}


# =========================================================
# HEADER
# =========================================================

st.markdown(
    '<div class="project-title">🏗️ Hassan Allam Construction</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="project-subtitle">'
    'Construction Projects & Pipeline Dashboard'
    '</div>',
    unsafe_allow_html=True
)


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.header("📁 Projects")

    selected_project = st.selectbox(
        "Select Project",
        project_names,
        index=0
    )

    st.divider()

    selected_file = project_files[selected_project]

    st.write("**Excel File**")

    st.caption(selected_file.name)

    st.divider()

    st.write("**Available Projects**")

    st.write(f"Total Projects: **{len(project_names)}**")


# =========================================================
# LOAD SELECTED PROJECT
# =========================================================

selected_file = project_files[selected_project]

sheets = load_excel_file(selected_file)


if not sheets:

    st.error(
        f"No sheets could be loaded from `{selected_file.name}`."
    )

    st.stop()


sheet_names = list(sheets.keys())


# =========================================================
# PROJECT INFORMATION
# =========================================================

st.info(
    f"📁 **Project:** {selected_project}  |  "
    f"📄 **Excel File:** {selected_file.name}  |  "
    f"📑 **Sheets:** {len(sheet_names)}"
)


# =========================================================
# SHEET BUTTONS
# =========================================================

st.markdown(
    "### 📑 Project Sheets"
)


# Number of buttons per row
BUTTONS_PER_ROW = 4


selected_sheet = st.session_state.get(
    "selected_sheet",
    sheet_names[0]
)


for start in range(0, len(sheet_names), BUTTONS_PER_ROW):

    row_sheets = sheet_names[
        start:start + BUTTONS_PER_ROW
    ]

    columns = st.columns(BUTTONS_PER_ROW)

    for i, sheet_name in enumerate(row_sheets):

        with columns[i]:

            button_label = f"📄 {sheet_name}"

            if st.button(
                button_label,
                key=f"sheet_{safe_key(selected_project)}_{safe_key(sheet_name)}",
                use_container_width=True
            ):
                st.session_state.selected_sheet = sheet_name
                st.rerun()


# =========================================================
# VALIDATE SELECTED SHEET
# =========================================================

if selected_sheet not in sheet_names:

    selected_sheet = sheet_names[0]
    st.session_state.selected_sheet = selected_sheet


# =========================================================
# GET SELECTED SHEET DATA
# =========================================================

df = sheets.get(selected_sheet)

df = clean_dataframe(df)


# =========================================================
# SHEET TITLE
# =========================================================

st.markdown(
    f'<div class="sheet-title">📄 {selected_sheet}</div>',
    unsafe_allow_html=True
)

st.caption(
    f"Source: {selected_file.name}"
)


# =========================================================
# KPI SECTION
# =========================================================

total_records = len(df)
total_columns = len(df.columns)

numeric_columns = len(
    df.select_dtypes(include="number").columns
)

blank_cells = int(
    df.isna().sum().sum()
)


col1, col2, col3, col4 = st.columns(4)


with col1:
    st.metric(
        "Records",
        f"{total_records:,}"
    )


with col2:
    st.metric(
        "Columns",
        f"{total_columns:,}"
    )


with col3:
    st.metric(
        "Numeric Columns",
        f"{numeric_columns:,}"
    )


with col4:
    st.metric(
        "Blank Cells",
        f"{blank_cells:,}"
    )


# =========================================================
# SEARCH
# =========================================================

st.markdown("### 🔎 Search")

search_text = st.text_input(
    "Search in current sheet",
    placeholder="Type a value to search..."
)


filtered_df = filter_dataframe(
    df,
    search_text
)


# =========================================================
# SEARCH RESULT INFO
# =========================================================

if search_text:

    st.info(
        f"Showing **{len(filtered_df):,}** records "
        f"matching **{search_text}**."
    )


# =========================================================
# DATA TABLE
# =========================================================

st.markdown("### 📊 Sheet Data")


if filtered_df.empty:

    st.warning(
        "No data found for the selected search."
    )

else:

    st.dataframe(
        filtered_df,
        use_container_width=True,
        height=600,
        hide_index=True
    )


# =========================================================
# DOWNLOAD CURRENT SHEET
# =========================================================

st.markdown("### ⬇️ Download")


csv_data = filtered_df.to_csv(
    index=False
).encode("utf-8-sig")


download_name = (
    f"{selected_project}_{selected_sheet}.csv"
)


st.download_button(
    label="⬇️ Download Current Sheet",
    data=csv_data,
    file_name=download_name,
    mime="text/csv",
    use_container_width=False
)


# =========================================================
# FOOTER
# =========================================================

st.divider()

st.caption(
    "Hassan Allam Construction | "
    "Projects & Pipeline Dashboard"
)
