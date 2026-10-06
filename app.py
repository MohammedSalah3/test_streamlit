from pathlib import Path
from io import BytesIO

import pandas as pd
import openpyxl
import streamlit as st


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Hassan Allam Construction | Projects Dashboard",
    page_icon="🏗️",
    layout="wide",
)


# =========================================================
# PATHS
# =========================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_FOLDER = BASE_DIR / "data"


# =========================================================
# CSS
# =========================================================

st.markdown(
    """
    <style>

    .main {
        padding-top: 1rem;
    }

    .title {
        font-size: 30px;
        font-weight: 700;
        margin-bottom: 5px;
    }

    .subtitle {
        color: #6b7280;
        margin-bottom: 20px;
    }

    .section-title {
        font-size: 21px;
        font-weight: 700;
        margin-top: 20px;
        margin-bottom: 10px;
    }

    div[data-testid="stMetric"] {
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        padding: 12px;
        background: white;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# GET EXCEL FILES
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
# READ WORKBOOK
# =========================================================

@st.cache_data(show_spinner=False)
def read_workbook(file_path):

    """
    Read workbook using openpyxl.

    Returns:

    {
        sheet_name: {
            "dataframe": dataframe,
            "excel_tables": [...],
            "ranges": [...]
        }
    }
    """

    workbook = openpyxl.load_workbook(
        file_path,
        data_only=True
    )

    result = {}

    for ws in workbook.worksheets:

        # -------------------------------------------------
        # Read complete worksheet
        # -------------------------------------------------

        data = list(
            ws.values
        )

        if not data:
            result[ws.title] = {
                "dataframe": pd.DataFrame(),
                "excel_tables": [],
                "ranges": []
            }
            continue

        # Convert worksheet to dataframe
        df = pd.DataFrame(data)

        # Remove completely empty rows
        df = df.dropna(
            how="all"
        )

        # Remove completely empty columns
        df = df.dropna(
            axis=1,
            how="all"
        )

        # -------------------------------------------------
        # Excel Tables
        # -------------------------------------------------

        excel_tables = []

        for table_name, table in ws.tables.items():

            excel_tables.append(
                {
                    "name": table_name,
                    "range": table.ref
                }
            )

        # -------------------------------------------------
        # Detect normal data ranges
        # -------------------------------------------------

        ranges = detect_data_ranges(
            ws
        )

        result[ws.title] = {
            "dataframe": df,
            "excel_tables": excel_tables,
            "ranges": ranges
        }

    return result


# =========================================================
# DETECT DATA RANGES
# =========================================================

def detect_data_ranges(ws):

    """
    Detect independent blocks of cells.

    Example:

    A1:D8

    F1:J10

    A15:H25

    will be detected as 3 separate ranges.
    """

    min_row = ws.min_row
    max_row = ws.max_row
    min_col = ws.min_column
    max_col = ws.max_column

    if (
        min_row is None
        or max_row is None
        or min_col is None
        or max_col is None
    ):
        return []

    # -----------------------------------------------------
    # Find rows containing data
    # -----------------------------------------------------

    occupied_rows = []

    for r in range(
        min_row,
        max_row + 1
    ):

        has_data = False

        for c in range(
            min_col,
            max_col + 1
        ):

            value = ws.cell(
                r,
                c
            ).value

            if value is not None:

                has_data = True
                break

        if has_data:

            occupied_rows.append(r)

    if not occupied_rows:
        return []

    # -----------------------------------------------------
    # Split rows into groups
    # -----------------------------------------------------

    row_groups = []

    start = occupied_rows[0]
    previous = occupied_rows[0]

    for row in occupied_rows[1:]:

        if row == previous + 1:

            previous = row

        else:

            row_groups.append(
                (
                    start,
                    previous
                )
            )

            start = row
            previous = row

    row_groups.append(
        (
            start,
            previous
        )
    )

    # -----------------------------------------------------
    # Detect columns inside each row group
    # -----------------------------------------------------

    ranges = []

    for row_start, row_end in row_groups:

        occupied_columns = []

        for c in range(
            min_col,
            max_col + 1
        ):

            has_data = False

            for r in range(
                row_start,
                row_end + 1
            ):

                value = ws.cell(
                    r,
                    c
                ).value

                if value is not None:

                    has_data = True
                    break

            if has_data:

                occupied_columns.append(c)

        if not occupied_columns:
            continue

        # -------------------------------------------------
        # Split columns into groups
        # -------------------------------------------------

        column_groups = []

        start = occupied_columns[0]
        previous = occupied_columns[0]

        for col in occupied_columns[1:]:

            if col == previous + 1:

                previous = col

            else:

                column_groups.append(
                    (
                        start,
                        previous
                    )
                )

                start = col
                previous = col

        column_groups.append(
            (
                start,
                previous
            )
        )

        # -------------------------------------------------
        # Create ranges
        # -------------------------------------------------

        for col_start, col_end in column_groups:

            # Ignore single empty-looking cells
            if (
                row_end - row_start < 1
                and col_end - col_start < 1
            ):
                continue

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
    # Assume first row is header
    # -----------------------------------------------------

    header = values[0]

    data = values[1:]

    # Create unique column names
    columns = []

    for i, value in enumerate(header):

        if value is None:
            value = f"Column {i + 1}"

        value = str(value)

        # Avoid duplicate columns
        original = value
        counter = 2

        while value in columns:

            value = (
                f"{original}_{counter}"
            )

            counter += 1

        columns.append(value)

    df = pd.DataFrame(
        data,
        columns=columns
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

    return df


# =========================================================
# SAFE KEY
# =========================================================

def safe_key(text):

    return (
        str(text)
        .replace(" ", "_")
        .replace("-", "_")
        .replace(".", "_")
        .replace("/", "_")
        .replace("\\", "_")
        .replace(":", "_")
    )


# =========================================================
# DATA FOLDER CHECK
# =========================================================

excel_files = get_excel_files()


if not DATA_FOLDER.exists():

    st.error(
        "❌ The `data` folder does not exist."
    )

    st.code(
        str(DATA_FOLDER)
    )

    st.stop()


if not excel_files:

    st.error(
        "❌ No Excel files found in the `data` folder."
    )

    st.info(
        """
        Upload your Excel files into:

        data/

        Example:

        data/
        ├── Project1.xlsx
        ├── Project2.xlsx
        └── Project3.xlsx
        """
    )

    st.stop()


# =========================================================
# HEADER
# =========================================================

st.markdown(
    '<div class="title">'
    '🏗️ Hassan Allam Construction'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    'Construction Projects & Pipeline Dashboard'
    '</div>',
    unsafe_allow_html=True
)


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

    st.header("📁 Projects")

    selected_project = st.selectbox(
        "Select Project",
        project_names
    )

    st.divider()

    st.write(
        f"**Projects:** {len(project_names)}"
    )

    st.write(
        f"**Excel Files:** {len(excel_files)}"
    )


# =========================================================
# LOAD SELECTED PROJECT
# =========================================================

selected_file = project_files[
    selected_project
]

workbook_data = read_workbook(
    selected_file
)


if not workbook_data:

    st.error(
        "Unable to read the selected Excel file."
    )

    st.stop()


# =========================================================
# SHEETS
# =========================================================

sheet_names = list(
    workbook_data.keys()
)


# =========================================================
# SELECTED SHEET
# =========================================================

if "selected_sheet" not in st.session_state:

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
    f"📁 **Project:** {selected_project}  |  "
    f"📄 **File:** {selected_file.name}  |  "
    f"📑 **Sheets:** {len(sheet_names)}"
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
                    f"sheet_"
                    f"{safe_key(selected_project)}_"
                    f"{safe_key(sheet_name)}"
                ),
                use_container_width=True
            ):

                st.session_state.selected_sheet = (
                    sheet_name
                )

                st.rerun()


# =========================================================
# SELECTED SHEET DATA
# =========================================================

selected_sheet = (
    st.session_state.selected_sheet
)

sheet_info = workbook_data[
    selected_sheet
]


# =========================================================
# OPENPYXL WORKSHEET
# =========================================================

workbook = openpyxl.load_workbook(
    selected_file,
    data_only=True
)

ws = workbook[
    selected_sheet
]


# =========================================================
# SHEET HEADER
# =========================================================

st.subheader(
    f"📄 {selected_sheet}"
)

st.caption(
    f"Source: {selected_file.name}"
)


# =========================================================
# EXCEL TABLES
# =========================================================

excel_tables = sheet_info[
    "excel_tables"
]


# =========================================================
# DETECTED RANGES
# =========================================================

ranges = sheet_info[
    "ranges"
]


# =========================================================
# SEARCH
# =========================================================

search_text = st.text_input(
    "🔎 Search in current sheet",
    placeholder="Search..."
)


# =========================================================
# EXCEL TABLES SECTION
# =========================================================

if excel_tables:

    st.markdown(
        "### 📋 Excel Tables"
    )

    for table_index, table_info in enumerate(
        excel_tables,
        start=1
    ):

        table_name = table_info[
            "name"
        ]

        table_range = table_info[
            "range"
        ]

        st.markdown(
            f"#### 📋 {table_name}"
        )

        st.caption(
            f"Range: {table_range}"
        )

        # Parse range
        min_cell, max_cell = (
            table_range.split(":")
        )

        from openpyxl.utils.cell import (
            range_boundaries
        )

        (
            min_col,
            min_row,
            max_col,
            max_row
        ) = range_boundaries(
            table_range
        )

        df_table = range_to_dataframe(
            ws,
            min_row,
            max_row,
            min_col,
            max_col
        )

        # Search
        if search_text:

            mask = (
                df_table
                .astype(str)
                .apply(
                    lambda col:
                    col.str.contains(
                        search_text,
                        case=False,
                        na=False
                    )
                )
                .any(axis=1)
            )

            df_table = df_table[
                mask
            ]

        # Metrics
        c1, c2 = st.columns(2)

        with c1:

            st.metric(
                "Records",
                f"{len(df_table):,}"
            )

        with c2:

            st.metric(
                "Columns",
                f"{len(df_table.columns):,}"
            )

        st.dataframe(
            df_table,
            use_container_width=True,
            hide_index=True
        )

        st.divider()


# =========================================================
# NORMAL RANGES SECTION
# =========================================================

if ranges:

    st.markdown(
        "### 📦 Detected Ranges"
    )

    # Track Excel Table ranges
    excel_table_ranges = set()

    for table_info in excel_tables:

        excel_table_ranges.add(
            table_info["range"]
        )

    range_number = 0

    for r in ranges:

        range_number += 1

        min_row = r[
            "min_row"
        ]

        max_row = r[
            "max_row"
        ]

        min_col = r[
            "min_col"
        ]

        max_col = r[
            "max_col"
        ]

        from openpyxl.utils import (
            get_column_letter
        )

        range_address = (
            f"{get_column_letter(min_col)}"
            f"{min_row}:"
            f"{get_column_letter(max_col)}"
            f"{max_row}"
        )

        # -------------------------------------------------
        # Skip if this range exactly matches an Excel Table
        # -------------------------------------------------

        if range_address in excel_table_ranges:
            continue

        df_range = range_to_dataframe(
            ws,
            min_row,
            max_row,
            min_col,
            max_col
        )

        if df_range.empty:
            continue

        # Ignore very small single-cell blocks
        if (
            len(df_range) == 0
            and len(df_range.columns) <= 1
        ):
            continue

        st.markdown(
            f"#### 📦 Range {range_number}"
        )

        st.caption(
            f"Excel Range: {range_address}"
        )

        # Search
        display_range = df_range.copy()

        if search_text:

            mask = (
                display_range
                .astype(str)
                .apply(
                    lambda col:
                    col.str.contains(
                        search_text,
                        case=False,
                        na=False
                    )
                )
                .any(axis=1)
            )

            display_range = (
                display_range[mask]
            )

        # Metrics
        c1, c2 = st.columns(2)

        with c1:

            st.metric(
                "Records",
                f"{len(display_range):,}"
            )

        with c2:

            st.metric(
                "Columns",
                f"{len(display_range.columns):,}"
            )

        st.dataframe(
            display_range,
            use_container_width=True,
            hide_index=True
        )

        st.divider()


# =========================================================
# NO DATA
# =========================================================

if not excel_tables and not ranges:

    st.warning(
        "No tables or data ranges were detected "
        "in this sheet."
    )


# =========================================================
# SHEET SUMMARY
# =========================================================

st.markdown(
    "### 📊 Sheet Summary"
)

c1, c2, c3 = st.columns(3)

with c1:

    st.metric(
        "Excel Tables",
        len(excel_tables)
    )

with c2:

    st.metric(
        "Detected Ranges",
        len(ranges)
    )

with c3:

    st.metric(
        "Total Sheets",
        len(sheet_names)
    )


# =========================================================
# FOOTER
# =========================================================

st.divider()

st.caption(
    "Hassan Allam Construction | "
    "Projects & Pipeline Dashboard"
)
