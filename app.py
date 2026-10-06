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
    Detect separate data blocks inside the worksheet.

    Example:

    A1:D10   -> Table 1

    A13:C20  -> Table 2

    F1:H8    -> Table 3
    """

    max_row = ws.max_row
    max_col = ws.max_column

    # -----------------------------------------------------
    # Find rows containing data
    # -----------------------------------------------------

    rows_with_data = []

    for r in range(1, max_row + 1):

        has_data = False

        for c in range(1, max_col + 1):

            value = ws.cell(r, c).value

            if value is not None:
                has_data = True
                break

        if has_data:
            rows_with_data.append(r)

    if not rows_with_data:
        return []


    # -----------------------------------------------------
    # Group consecutive rows
    # -----------------------------------------------------

    row_groups = []

    start = rows_with_data[0]
    previous = rows_with_data[0]

    for row in rows_with_data[1:]:

        if row == previous + 1:

            previous = row

        else:

            row_groups.append(
                (start, previous)
            )

            start = row
            previous = row

    row_groups.append(
        (start, previous)
    )


    # -----------------------------------------------------
    # Find column blocks inside each row group
    # -----------------------------------------------------

    ranges = []

    for row_start, row_end in row_groups:

        columns_with_data = []

        for c in range(1, max_col + 1):

            has_data = False

            for r in range(
                row_start,
                row_end + 1
            ):

                if ws.cell(r, c).value is not None:

                    has_data = True
                    break

            if has_data:
                columns_with_data.append(c)


        if not columns_with_data:
            continue


        # -------------------------------------------------
        # Group consecutive columns
        # -------------------------------------------------

        col_groups = []

        start_col = columns_with_data[0]
        previous_col = columns_with_data[0]

        for col in columns_with_data[1:]:

            if col == previous_col + 1:

                previous_col = col

            else:

                col_groups.append(
                    (start_col, previous_col)
                )

                start_col = col
                previous_col = col

        col_groups.append(
            (start_col, previous_col)
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


    # Remove completely empty rows

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


    # -----------------------------------------------------
    # First row = header
    # -----------------------------------------------------

    header = values[0]

    data = values[1:]


    # -----------------------------------------------------
    # Create unique column names
    # -----------------------------------------------------

    columns = []

    for index, column in enumerate(header):

        if column is None or str(column).strip() == "":

            column_name = f"Column {index + 1}"

        else:

            column_name = str(column)


        # Make column names unique

        original_name = column_name
        counter = 2

        while column_name in columns:

            column_name = (
                f"{original_name} {counter}"
            )

            counter += 1

        columns.append(column_name)


    # -----------------------------------------------------
    # Create DataFrame
    # -----------------------------------------------------

    if data:

        df = pd.DataFrame(
            data,
            columns=columns
        )

    else:

        df = pd.DataFrame(
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
# LOAD EXCEL WORKBOOK
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

            tables = []

            # -------------------------------------------------
            # Excel Tables
            # -------------------------------------------------

            for table_name, table in ws.tables.items():

                tables.append(
                    {
                        "name": table_name,
                        "range": table.ref
                    }
                )


            project_data[ws.title] = {
                "ranges": ranges,
                "tables": tables
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
# GET FILES
# =========================================================

excel_files = get_excel_files()


# =========================================================
# CHECK DATA FOLDER
# =========================================================

if not DATA_FOLDER.exists():

    st.error(
        "❌ The data folder does not exist."
    )

    st.code(
        str(DATA_FOLDER)
    )

    st.stop()


# =========================================================
# CHECK EXCEL FILES
# =========================================================

if not excel_files:

    st.error(
        "❌ No Excel files found inside the data folder."
    )

    st.code(
        str(DATA_FOLDER)
    )

    items = list(
        DATA_FOLDER.iterdir()
    )

    if items:

        st.write("Files/folders found:")

        for item in items:

            st.write(
                f"- {item.name}"
            )

    else:

        st.warning(
            "The data folder is empty."
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

    st.header("📁 Projects")

    selected_project = st.selectbox(
        "Select Project",
        project_names
    )

    st.divider()

    st.write("Excel files found:")

    st.write(
        f"**{len(excel_files)}**"
    )

    st.divider()

    for file in excel_files:

        st.caption(
            f"📄 {file.name}"
        )


# =========================================================
# LOAD PROJECT
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
        "❌ No sheets were found."
    )

    st.stop()


sheet_names = list(
    project_data.keys()
)


# =========================================================
# SHEET SELECTION
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

sheet_data = project_data[
    selected_sheet
]


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
# DISPLAY DATA ONLY
# =========================================================

st.subheader(
    "📊 Sheet Data"
)


# =========================================================
# EXCEL TABLES
# =========================================================

if sheet_data["tables"]:

    st.markdown(
        "### 📋 Excel Tables"
    )

    for table in sheet_data["tables"]:

        table_name = table["name"]
        table_range = table["range"]

        st.markdown(
            f"#### {table_name}"
        )

        # Parse Excel range

        from openpyxl.utils.cell import range_boundaries

        min_col, min_row, max_col, max_row = (
            range_boundaries(table_range)
        )

        df_table = range_to_dataframe(
            # Need worksheet
            openpyxl.load_workbook(
                selected_file,
                data_only=True
            )[selected_sheet],

            min_row,
            max_row,
            min_col,
            max_col
        )

        if not df_table.empty:

            st.dataframe(
                df_table,
                use_container_width=True,
                hide_index=True
            )


# =========================================================
# OTHER RANGES
# =========================================================

st.markdown(
    "### 📦 Other Data Ranges"
)


workbook = openpyxl.load_workbook(
    selected_file,
    data_only=True
)

ws = workbook[selected_sheet]


ranges = sheet_data["ranges"]


# ---------------------------------------------------------
# Remove ranges already represented by Excel Tables
# ---------------------------------------------------------

table_ranges = set(
    table["range"]
    for table in sheet_data["tables"]
)


range_number = 1


for item in ranges:

    min_row = item["min_row"]
    max_row = item["max_row"]
    min_col = item["min_col"]
    max_col = item["max_col"]


    # Excel range notation

    from openpyxl.utils import get_column_letter

    excel_range = (
        f"{get_column_letter(min_col)}{min_row}:"
        f"{get_column_letter(max_col)}{max_row}"
    )


    # Skip if already an Excel Table

    if excel_range in table_ranges:

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


    st.markdown(
        f"#### Range {range_number}"
    )

    st.caption(
        f"Excel Range: `{excel_range}`"
    )

    st.dataframe(
        df_range,
        use_container_width=True,
        hide_index=True
    )

    range_number += 1


# =========================================================
# DOWNLOAD
# =========================================================

st.divider()

st.subheader(
    "⬇️ Download"
)


# Create combined data for download

all_download_data = []

for item in ranges:

    df_range = range_to_dataframe(
        ws,
        item["min_row"],
        item["max_row"],
        item["min_col"],
        item["max_col"]
    )

    if not df_range.empty:

        all_download_data.append(
            df_range
        )


if all_download_data:

    download_df = pd.concat(
        all_download_data,
        ignore_index=True
    )

    csv = download_df.to_csv(
        index=False
    ).encode(
        "utf-8-sig"
    )

    st.download_button(
        "⬇️ Download Sheet Data",
        csv,
        file_name=f"{selected_sheet}.csv",
        mime="text/csv"
    )
