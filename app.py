from pathlib import Path

import pandas as pd
import streamlit as st
import openpyxl


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Hassan Allam Construction | Projects Dashboard",
    page_icon="🏗️",
    layout="wide"
)


# =========================================================
# CONFIG
# =========================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_FOLDER = BASE_DIR / "data"


# =========================================================
# CSS
# =========================================================

st.markdown(
    """
    <style>

        .project-title {
            font-size: 30px;
            font-weight: 700;
        }

        .project-subtitle {
            color: #6b7280;
            margin-bottom: 20px;
        }

    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# FIND EXCEL FILES
# =========================================================

def get_excel_files():

    if not DATA_FOLDER.exists():
        return []

    files = []

    for file in DATA_FOLDER.iterdir():

        if (
            file.is_file()
            and file.suffix.lower() in [".xlsx", ".xls"]
        ):
            files.append(file)

    return sorted(
        files,
        key=lambda x: x.name.lower()
    )


# =========================================================
# FIND DATA RANGES
# =========================================================

def find_ranges(ws):

    """
    Find separate blocks of data inside the worksheet.

    Example:

    A1:D20   -> first table
    A23:F40  -> second table
    H1:K15   -> third table
    """

    max_row = ws.max_row
    max_col = ws.max_column

    # -----------------------------------------------------
    # Find rows that contain data
    # -----------------------------------------------------

    rows_with_data = []

    for row_number in range(
        1,
        max_row + 1
    ):

        has_data = False

        for col_number in range(
            1,
            max_col + 1
        ):

            value = ws.cell(
                row_number,
                col_number
            ).value

            if value is not None:

                has_data = True
                break

        if has_data:

            rows_with_data.append(
                row_number
            )


    if not rows_with_data:

        return []


    # -----------------------------------------------------
    # Group consecutive rows
    # -----------------------------------------------------

    row_groups = []

    start_row = rows_with_data[0]
    previous_row = rows_with_data[0]

    for row_number in rows_with_data[1:]:

        if row_number == previous_row + 1:

            previous_row = row_number

        else:

            row_groups.append(
                (
                    start_row,
                    previous_row
                )
            )

            start_row = row_number
            previous_row = row_number


    row_groups.append(
        (
            start_row,
            previous_row
        )
    )


    # -----------------------------------------------------
    # Find column blocks
    # -----------------------------------------------------

    ranges = []

    for row_start, row_end in row_groups:

        columns_with_data = []

        for col_number in range(
            1,
            max_col + 1
        ):

            has_data = False

            for row_number in range(
                row_start,
                row_end + 1
            ):

                value = ws.cell(
                    row_number,
                    col_number
                ).value

                if value is not None:

                    has_data = True
                    break

            if has_data:

                columns_with_data.append(
                    col_number
                )


        if not columns_with_data:

            continue


        # -------------------------------------------------
        # Group consecutive columns
        # -------------------------------------------------

        col_groups = []

        start_col = columns_with_data[0]
        previous_col = columns_with_data[0]

        for col_number in columns_with_data[1:]:

            if col_number == previous_col + 1:

                previous_col = col_number

            else:

                col_groups.append(
                    (
                        start_col,
                        previous_col
                    )
                )

                start_col = col_number
                previous_col = col_number


        col_groups.append(
            (
                start_col,
                previous_col
            )
        )


        # -------------------------------------------------
        # Save ranges
        # -------------------------------------------------

        for col_start, col_end in col_groups:

            ranges.append(
                {
                    "min_row": row_start,
                    "max_row": row_end,
                    "min_col": col_start,
                    "max_col": col_end
                }
            )


    return ranges


# =========================================================
# RANGE TO DATAFRAME
# =========================================================

def range_to_dataframe(
    ws,
    min_row,
    max_row,
    min_col,
    max_col
):

    values = []

    # -----------------------------------------------------
    # Read cells
    # -----------------------------------------------------

    for row in ws.iter_rows(
        min_row=min_row,
        max_row=max_row,
        min_col=min_col,
        max_col=max_col,
        values_only=True
    ):

        values.append(
            list(row)
        )


    if not values:

        return pd.DataFrame()


    # -----------------------------------------------------
    # Remove completely empty rows
    # -----------------------------------------------------

    values = [
        row
        for row in values
        if any(
            value is not None
            for value in row
        )
    ]


    if not values:

        return pd.DataFrame()


    # =====================================================
    # FIRST ROW = COLUMN NAMES
    # =====================================================

    header = values[0]

    data = values[1:]


    # =====================================================
    # CLEAN COLUMN NAMES
    # =====================================================

    columns = []

    for index, column in enumerate(header):

        # Empty header
        if (
            column is None
            or str(column).strip() == ""
        ):

            column_name = (
                f"Column {index + 1}"
            )

        else:

            column_name = str(
                column
            ).strip()


        # -------------------------------------------------
        # Make duplicate names unique
        # -------------------------------------------------

        original_name = column_name

        counter = 2

        while column_name in columns:

            column_name = (
                f"{original_name}_{counter}"
            )

            counter += 1

        columns.append(
            column_name
        )


    # =====================================================
    # CREATE DATAFRAME
    # =====================================================

    if data:

        df = pd.DataFrame(
            data,
            columns=columns
        )

    else:

        df = pd.DataFrame(
            columns=columns
        )


    # =====================================================
    # REMOVE EMPTY ROWS
    # =====================================================

    df = df.dropna(
        how="all"
    )


    # =====================================================
    # REMOVE EMPTY COLUMNS
    # =====================================================

    df = df.dropna(
        axis=1,
        how="all"
    )


    return df


# =========================================================
# LOAD EXCEL FILE
# =========================================================

@st.cache_data(show_spinner=False)
def load_excel_file(file_path):

    try:

        workbook = openpyxl.load_workbook(
            file_path,
            data_only=True
        )

        project_data = {}

        for ws in workbook.worksheets:

            ranges = find_ranges(ws)

            project_data[ws.title] = {
                "ranges": ranges
            }

        return project_data


    except Exception as e:

        st.error(
            f"Error reading {file_path.name}: {e}"
        )

        return {}


# =========================================================
# HEADER
# =========================================================

st.markdown(
    '<div class="project-title">'
    '🏗️ Hassan Allam Construction'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="project-subtitle">'
    'Construction Projects & Pipeline Dashboard'
    '</div>',
    unsafe_allow_html=True
)


# =========================================================
# GET EXCEL FILES
# =========================================================

excel_files = get_excel_files()


# =========================================================
# CHECK DATA FOLDER
# =========================================================

if not DATA_FOLDER.exists():

    st.error(
        "❌ The data folder does not exist."
    )

    st.write(
        "Expected location:"
    )

    st.code(
        str(DATA_FOLDER)
    )

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


# =========================================================
# CHECK EXCEL FILES
# =========================================================

if not excel_files:

    st.error(
        "❌ No Excel files found inside the data folder."
    )

    st.write(
        "The application is looking here:"
    )

    st.code(
        str(DATA_FOLDER)
    )

    items = list(
        DATA_FOLDER.iterdir()
    )

    if items:

        st.write(
            "Files/folders currently found:"
        )

        for item in items:

            st.write(
                f"- {item.name}"
            )

    else:

        st.warning(
            "The data folder is completely empty."
        )

    st.info(
        """
        Upload your Excel file inside the GitHub
        `data` folder.

        Example:

        data/
        └── Project.xlsx
        """
    )

    st.stop()


# =========================================================
# PROJECTS
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
# SIDEBAR
# =========================================================

with st.sidebar:

    st.header(
        "📁 Projects"
    )

    selected_project = st.selectbox(
        "Select Project",
        project_names
    )

    st.divider()

    st.write(
        "Excel files found:"
    )

    st.write(
        f"**{len(excel_files)}**"
    )

    st.divider()

    for file in excel_files:

        st.caption(
            f"📄 {file.name}"
        )


# =========================================================
# LOAD SELECTED PROJECT
# =========================================================

selected_file = project_files[
    selected_project
]

project_data = load_excel_file(
    selected_file
)


# =========================================================
# CHECK SHEETS
# =========================================================

if not project_data:

    st.error(
        "❌ No sheets were found in this Excel file."
    )

    st.stop()


sheet_names = list(
    project_data.keys()
)


# =========================================================
# SELECTED SHEET STATE
# =========================================================

if (
    "selected_sheet"
    not in st.session_state
):

    st.session_state.selected_sheet = (
        sheet_names[0]
    )


if (
    st.session_state.selected_sheet
    not in sheet_names
):

    st.session_state.selected_sheet = (
        sheet_names[0]
    )


# =========================================================
# PROJECT INFO
# =========================================================

st.info(
    f"📁 Project: **{selected_project}**  |  "
    f"📄 File: **{selected_file.name}**  |  "
    f"📑 Sheets: **{len(sheet_names)}**"
)


# =========================================================
# SHEET BUTTONS
# =========================================================

st.subheader(
    "📑 Project Sheets"
)

BUTTONS_PER_ROW = 4


for start in range(
    0,
    len(sheet_names),
    BUTTONS_PER_ROW
):

    row = sheet_names[
        start:start + BUTTONS_PER_ROW
    ]

    cols = st.columns(
        BUTTONS_PER_ROW
    )

    for i, sheet_name in enumerate(row):

        with cols[i]:

            if st.button(
                f"📄 {sheet_name}",
                key=(
                    f"sheet_{start}_"
                    f"{i}_{sheet_name}"
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


# =========================================================
# SHEET TITLE
# =========================================================

st.subheader(
    f"📄 {selected_sheet}"
)

st.caption(
    f"Source: {selected_file.name}"
)


# =========================================================
# OPEN WORKBOOK
# =========================================================

workbook = openpyxl.load_workbook(
    selected_file,
    data_only=True
)

ws = workbook[
    selected_sheet
]


# =========================================================
# FIND DATA RANGES
# =========================================================

ranges = find_ranges(
    ws
)


# =========================================================
# DISPLAY DATA
# =========================================================

if not ranges:

    st.info(
        "This sheet does not contain any data."
    )

else:

    for item in ranges:

        min_row = item["min_row"]
        max_row = item["max_row"]
        min_col = item["min_col"]
        max_col = item["max_col"]


        # -------------------------------------------------
        # Convert Range to DataFrame
        # -------------------------------------------------

        df = range_to_dataframe(
            ws,
            min_row,
            max_row,
            min_col,
            max_col
        )


        # -------------------------------------------------
        # Skip empty data
        # -------------------------------------------------

        if df.empty:

            continue


        # -------------------------------------------------
        # DISPLAY
        #
        # FIRST ROW = COLUMN NAMES
        # REST = DATA
        # -------------------------------------------------

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
            height=600
        )


        # Small spacing
        st.markdown(
            "<br>",
            unsafe_allow_html=True
        )
