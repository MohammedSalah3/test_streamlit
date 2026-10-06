import streamlit as st
import pandas as pd
import numpy as np

from io import BytesIO
import re
import json

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Hassan Allam Construction",
    page_icon="🏗️",
    layout="wide"
)


# =========================================================
# GOOGLE DRIVE CONFIG
# =========================================================

DRIVE_FOLDER_ID = "1V8-f6hDj7dEX6RImfDZi11FJvl6V1cvn"

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly"
]


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
# CSS
# =========================================================

st.markdown(
    """
    <style>

    .main-title {
        font-size: 30px;
        font-weight: 700;
        margin-bottom: 5px;
    }

    .sub-title {
        color: #666666;
        font-size: 15px;
        margin-bottom: 25px;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# GOOGLE DRIVE AUTHENTICATION
# =========================================================

@st.cache_resource
def get_drive_service():

    try:

        # ---------------------------------------------
        # Streamlit Secrets
        # ---------------------------------------------

        service_account_info = dict(
            st.secrets["gcp_service_account"]
        )

        credentials = (
            service_account
            .Credentials
            .from_service_account_info(
                service_account_info,
                scopes=SCOPES
            )
        )

        service = build(
            "drive",
            "v3",
            credentials=credentials
        )

        return service

    except Exception as e:

        st.error(
            "Google Drive authentication failed."
        )

        st.code(
            str(e)
        )

        return None


# =========================================================
# LIST EXCEL FILES FROM GOOGLE DRIVE
# =========================================================

def get_drive_excel_files(service):

    files = []

    page_token = None

    query = (
        f"'{DRIVE_FOLDER_ID}' in parents "
        "and trashed = false "
        "and ("
        "mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' "
        "or "
        "mimeType = 'application/vnd.ms-excel'"
        ")"
    )

    try:

        while True:

            response = (
                service.files()
                .list(
                    q=query,
                    spaces="drive",
                    fields=(
                        "nextPageToken,"
                        "files(id,name,mimeType,"
                        "modifiedTime,size)"
                    ),
                    pageToken=page_token,
                    pageSize=1000,
                    orderBy="name"
                )
                .execute()
            )

            files.extend(
                response.get(
                    "files",
                    []
                )
            )

            page_token = response.get(
                "nextPageToken"
            )

            if not page_token:
                break

        return files

    except Exception as e:

        st.error(
            "Could not read files from Google Drive."
        )

        st.code(
            str(e)
        )

        return []


# =========================================================
# DOWNLOAD FILE FROM GOOGLE DRIVE
# =========================================================

def download_drive_file(
    service,
    file_id
):

    try:

        request = (
            service.files()
            .get(
                fileId=file_id,
                alt="media"
            )
        )

        file_buffer = BytesIO()

        downloader = MediaIoBaseDownload(
            file_buffer,
            request
        )

        done = False

        while not done:

            status, done = (
                downloader.next_chunk()
            )

        file_buffer.seek(0)

        return file_buffer.getvalue()

    except Exception as e:

        st.error(
            f"Could not download file: {e}"
        )

        return None


# =========================================================
# CLEAN PROJECT NAME
# =========================================================

def project_name_from_file(
    filename
):

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


# =========================================================
# CLEAN COLUMN NAMES
# =========================================================

def clean_column_name(
    column
):

    if column is None:
        return ""

    column = str(
        column
    ).strip()

    column = re.sub(
        r"\s+",
        " ",
        column
    )

    return column


def make_unique_columns(
    columns
):

    result = []

    counts = {}

    for column in columns:

        column = clean_column_name(
            column
        )

        if not column:

            column = "Unnamed"

        if column not in counts:

            counts[column] = 0

            result.append(
                column
            )

        else:

            counts[column] += 1

            result.append(
                f"{column}_{counts[column]}"
            )

    return result


# =========================================================
# READ EXCEL FROM BYTES
# =========================================================

def read_excel_bytes(
    file_bytes
):

    sheets = {}

    try:

        excel = pd.ExcelFile(
            BytesIO(file_bytes)
        )

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

                df.columns = (
                    make_unique_columns(
                        df.columns
                    )
                )

                df = df.reset_index(
                    drop=True
                )

                sheets[
                    sheet_name
                ] = df

            except Exception as e:

                st.warning(
                    f"Could not read "
                    f"sheet '{sheet_name}': "
                    f"{e}"
                )

        return sheets

    except Exception as e:

        st.error(
            f"Could not read Excel file: {e}"
        )

        return {}


# =========================================================
# LOAD PROJECTS FROM DRIVE
# =========================================================

def load_projects_from_drive(
    service,
    drive_files
):

    projects = {}

    progress = st.progress(
        0
    )

    total_files = len(
        drive_files
    )

    for index, file_info in enumerate(
        drive_files
    ):

        file_id = file_info[
            "id"
        ]

        file_name = file_info[
            "name"
        ]

        # ---------------------------------------------
        # Download
        # ---------------------------------------------

        file_bytes = download_drive_file(
            service,
            file_id
        )

        if file_bytes is None:
            continue

        # ---------------------------------------------
        # Project name
        # ---------------------------------------------

        project_name = (
            project_name_from_file(
                file_name
            )
        )

        # ---------------------------------------------
        # Read all sheets
        # ---------------------------------------------

        sheets = read_excel_bytes(
            file_bytes
        )

        # ---------------------------------------------
        # Store project
        # ---------------------------------------------

        projects[
            project_name
        ] = {

            "file_id": file_id,

            "file_name": file_name,

            "modified_time":
                file_info.get(
                    "modifiedTime",
                    ""
                ),

            "sheets": sheets
        }

        # ---------------------------------------------
        # Progress
        # ---------------------------------------------

        if total_files > 0:

            progress.progress(
                (index + 1)
                / total_files
            )

    progress.empty()

    return projects


# =========================================================
# PR SHEET DETECTION
# =========================================================

def is_pr_sheet(
    sheet_name
):

    pattern = (
        r"^PR[-_\s]?\d+$"
    )

    return bool(
        re.match(
            pattern,
            str(sheet_name).strip(),
            re.IGNORECASE
        )
    )


# =========================================================
# FIND PR COLUMN
# =========================================================

def find_pr_column(
    df
):

    if df.empty:
        return None

    candidates = [

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

    columns = list(
        df.columns
    )

    # Exact match
    for candidate in candidates:

        for column in columns:

            if (
                str(column)
                .strip()
                .lower()
                ==
                candidate.lower()
            ):

                return column

    # Partial match
    for column in columns:

        text = (
            str(column)
            .strip()
            .lower()
        )

        if re.fullmatch(
            r"pr",
            text
        ):

            return column

        if (
            "purchase requisition"
            in text
        ):

            return column

        if (
            "requisition number"
            in text
        ):

            return column

        if (
            "requisition no"
            in text
        ):

            return column

    return None


# =========================================================
# MAIN HEADER
# =========================================================

st.markdown(
    """
    <div class="main-title">
        Hassan Allam Construction
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div class="sub-title">
        Construction Projects Dashboard
    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# CONNECT TO GOOGLE DRIVE
# =========================================================

drive_service = get_drive_service()

if drive_service is None:

    st.stop()


# =========================================================
# DRIVE CONTROL
# =========================================================

col1, col2 = st.columns(
    [5, 1]
)

with col1:

    st.info(
        "📁 Projects are loaded automatically "
        "from the Google Drive folder."
    )

with col2:

    refresh = st.button(
        "🔄 Refresh",
        use_container_width=True
    )


# =========================================================
# LOAD DRIVE FILES
# =========================================================

drive_files = get_drive_excel_files(
    drive_service
)


if not drive_files:

    st.warning(
        "No Excel files were found "
        "inside the Google Drive folder."
    )

    st.stop()


# =========================================================
# LOAD PROJECTS
# =========================================================

with st.spinner(
    "Loading projects from Google Drive..."
):

    projects = load_projects_from_drive(
        drive_service,
        drive_files
    )


st.session_state.projects = projects


# =========================================================
# PROJECT LIST
# =========================================================

project_names = sorted(
    st.session_state
    .projects
    .keys()
)


if not project_names:

    st.warning(
        "No projects could be loaded."
    )

    st.stop()


# =========================================================
# PROJECT SELECTOR
# =========================================================

st.subheader(
    "🏗️ Project"
)


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
# SELECT PROJECT
# =========================================================

project = (
    st.session_state
    .projects
    .get(
        selected_project
    )
)


if project is None:

    st.error(
        "Project could not be found."
    )

    st.stop()


sheets = project.get(
    "sheets",
    {}
)


if not sheets:

    st.warning(
        "This project does not contain "
        "readable Excel sheets."
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
    "Each button displays the data stored "
    "inside that specific Excel sheet."
)


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

    columns = st.columns(
        len(row_sheets)
    )

    for i, sheet_name in enumerate(
        row_sheets
    ):

        with columns[i]:

            selected = (
                st.session_state
                .selected_sheet
                ==
                sheet_name
            )

            if selected:

                button_text = (
                    f"✅ {sheet_name}"
                )

            else:

                button_text = (
                    f"📄 {sheet_name}"
                )

            if st.button(
                button_text,
                key=(
                    f"sheet_button_"
                    f"{row_start}_{i}"
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
    st.session_state
    .selected_sheet
)


# =========================================================
# GET EXACT SHEET DATA
# =========================================================

df = sheets.get(
    selected_sheet
)


if df is None:

    st.error(
        "The selected sheet could not be loaded."
    )

    st.stop()


# IMPORTANT:
# This DataFrame belongs ONLY to the selected sheet.
df = df.copy()


# =========================================================
# SHEET INFORMATION
# =========================================================

st.divider()

st.subheader(
    f"📄 {selected_sheet}"
)


info1, info2, info3 = st.columns(
    3
)


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
# PR INFORMATION
# =========================================================

if is_pr_sheet(
    selected_sheet
):

    st.info(
        f"🔎 Current PR: **{selected_sheet}**"
    )

else:

    pr_column = find_pr_column(
        df
    )

    if pr_column:

        st.info(
            f"🔎 PR Column: **{pr_column}**"
        )


# =========================================================
# SEARCH
# =========================================================

st.subheader(
    "🔍 Search"
)

search_text = st.text_input(
    "Search inside the selected sheet",
    placeholder=(
        "Enter any value..."
    ),
    key=(
        f"search_{selected_sheet}"
    )
)


filtered_df = df.copy()


if search_text:

    search_text = (
        search_text
        .strip()
    )

    if search_text:

        mask = pd.Series(
            False,
            index=filtered_df.index
        )

        for column in filtered_df.columns:

            mask = (
                mask
                |
                filtered_df[
                    column
                ]
                .astype(str)
                .str.contains(
                    search_text,
                    case=False,
                    na=False,
                    regex=False
                )
            )

        filtered_df = (
            filtered_df[
                mask
            ]
            .copy()
        )


# =========================================================
# DATA TABLE
# =========================================================

st.subheader(
    f"📋 Data stored in {selected_sheet}"
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
# SIMPLE SHEET KPIs
# =========================================================

st.subheader(
    "📊 Sheet Summary"
)


numeric_columns = (
    filtered_df
    .select_dtypes(
        include=np.number
    )
    .columns
)


blank_cells = int(
    filtered_df
    .isna()
    .sum()
    .sum()
)


k1, k2, k3, k4 = st.columns(
    4
)


with k1:

    st.metric(
        "Records",
        f"{len(filtered_df):,}"
    )


with k2:

    st.metric(
        "Columns",
        f"{len(filtered_df.columns):,}"
    )


with k3:

    st.metric(
        "Numeric Columns",
        f"{len(numeric_columns):,}"
    )


with k4:

    st.metric(
        "Blank Cells",
        f"{blank_cells:,}"
    )


# =========================================================
# DOWNLOAD
# =========================================================

st.subheader(
    "⬇️ Export"
)


csv_data = (
    filtered_df
    .to_csv(
        index=False
    )
    .encode(
        "utf-8-sig"
    )
)


st.download_button(
    "⬇️ Download Current Sheet",
    data=csv_data,
    file_name=(
        f"{selected_project}_"
        f"{selected_sheet}.csv"
    ),
    mime="text/csv",
    use_container_width=True
)
