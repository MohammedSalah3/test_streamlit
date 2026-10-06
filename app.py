from pathlib import Path
import pandas as pd
import streamlit as st

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

        if file.is_file() and file.suffix.lower() in [".xlsx", ".xls"]:
            files.append(file)

    return sorted(files, key=lambda x: x.name.lower())


# =========================================================
# READ EXCEL
# =========================================================

@st.cache_data(show_spinner=False)
def load_excel_file(file_path):

    try:

        excel = pd.ExcelFile(file_path)

        sheets = {}

        for sheet_name in excel.sheet_names:

            df = pd.read_excel(
                file_path,
                sheet_name=sheet_name
            )

            # Remove completely empty rows
            df = df.dropna(how="all")

            # Remove completely empty columns
            df = df.dropna(axis=1, how="all")

            sheets[sheet_name] = df

        return sheets

    except Exception as e:

        st.error(
            f"Error reading {file_path.name}: {e}"
        )

        return {}


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
# GET FILES
# =========================================================

excel_files = get_excel_files()


# =========================================================
# CHECK DATA FOLDER
# =========================================================

if not DATA_FOLDER.exists():

    st.error("❌ The data folder does not exist.")

    st.write(
        f"Expected location:"
    )

    st.code(str(DATA_FOLDER))

    st.info(
        """
        Your GitHub structure must be:

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

    st.error("❌ No Excel files found inside the data folder.")

    st.write("The application is looking here:")

    st.code(str(DATA_FOLDER))

    st.write("Files/folders currently found:")

    items = list(DATA_FOLDER.iterdir())

    if items:

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
        Upload your Excel file inside the GitHub `data` folder.

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

selected_file = project_files[selected_project]

sheets = load_excel_file(selected_file)


# =========================================================
# CHECK SHEETS
# =========================================================

if not sheets:

    st.error(
        "❌ No sheets were found in this Excel file."
    )

    st.stop()


sheet_names = list(sheets.keys())


# =========================================================
# SHEET SELECTION
# =========================================================

if "selected_sheet" not in st.session_state:

    st.session_state.selected_sheet = sheet_names[0]


if st.session_state.selected_sheet not in sheet_names:

    st.session_state.selected_sheet = sheet_names[0]


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

st.subheader("📑 Project Sheets")

BUTTONS_PER_ROW = 4

for start in range(
    0,
    len(sheet_names),
    BUTTONS_PER_ROW
):

    row = sheet_names[
        start:start + BUTTONS_PER_ROW
    ]

    cols = st.columns(BUTTONS_PER_ROW)

    for i, sheet_name in enumerate(row):

        with cols[i]:

            if st.button(
                f"📄 {sheet_name}",
                key=f"sheet_{start}_{i}_{sheet_name}",
                use_container_width=True
            ):

                st.session_state.selected_sheet = sheet_name

                st.rerun()


# =========================================================
# SELECTED SHEET
# =========================================================

selected_sheet = st.session_state.selected_sheet

df = sheets[selected_sheet]


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
# KPIs
# =========================================================

records = len(df)

columns = len(df.columns)

numeric_columns = len(
    df.select_dtypes(
        include="number"
    ).columns
)

blank_cells = int(
    df.isna().sum().sum()
)


c1, c2, c3, c4 = st.columns(4)

with c1:

    st.metric(
        "Records",
        f"{records:,}"
    )

with c2:

    st.metric(
        "Columns",
        f"{columns:,}"
    )

with c3:

    st.metric(
        "Numeric Columns",
        f"{numeric_columns:,}"
    )

with c4:

    st.metric(
        "Blank Cells",
        f"{blank_cells:,}"
    )


# =========================================================
# SEARCH
# =========================================================

st.subheader("🔎 Search")

search = st.text_input(
    "Search current sheet",
    placeholder="Search..."
)


filtered_df = df.copy()


if search:

    search_lower = search.lower()

    mask = (
        filtered_df
        .astype(str)
        .apply(
            lambda col:
            col.str.lower().str.contains(
                search_lower,
                na=False
            )
        )
        .any(axis=1)
    )

    filtered_df = filtered_df[mask]

    st.info(
        f"Found {len(filtered_df):,} records."
    )


# =========================================================
# TABLE
# =========================================================

st.subheader("📊 Data")

st.dataframe(
    filtered_df,
    use_container_width=True,
    height=600,
    hide_index=True
)


# =========================================================
# DOWNLOAD
# =========================================================

csv = filtered_df.to_csv(
    index=False
).encode("utf-8-sig")


st.download_button(
    "⬇️ Download Current Sheet",
    csv,
    file_name=f"{selected_sheet}.csv",
    mime="text/csv"
)
