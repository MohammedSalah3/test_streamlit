"""
Construction Projects Dashboard  (Streamlit MVP - single file)
Excel -> Pandas -> Dynamic sheet buttons -> Reusable dashboard
"""
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Construction Projects Dashboard", page_icon="🏗️", layout="wide")

# ----------------------------------------------------------------------------
# 1) CONFIGURATION  -> غيّر أسماء الأعمدة من هنا فقط
#    كل مفتاح له قائمة أسماء محتملة (مش حساسة لحالة الحروف)
# ----------------------------------------------------------------------------
COLUMN_MAP = {
    "project":       ["Project", "Project Name"],
    "item":          ["Item", "Material", "Description"],
    "location":      ["Location", "Site", "Zone"],
    "required_qty":  ["Required Qty", "Required Quantity", "Quantity", "Qty"],
    "delivered_qty": ["Delivered Qty", "Delivered Quantity"],
    "remaining_qty": ["Remaining Qty", "Remaining Quantity"],
    "required_date": ["Required Date"],
    "expected_date": ["Expected Date"],
    "actual_date":   ["Actual Date", "Delivery Date"],
    "status":        ["Status"],
}
PRIMARY = "#1f4e9c"
STATUS_COLORS = {"Delivered": "#2e9e5b", "Pending / On Track": "#1f6fd1",
                 "Delayed": "#d9534f", "Unknown": "#9aa5b1"}

# ----------------------------------------------------------------------------
# 2) STYLE
# ----------------------------------------------------------------------------
st.markdown(f"""
<style>
.stApp {{ background:#ffffff; }}
h1,h2,h3 {{ color:{PRIMARY}; }}
.kpi {{ background:#f4f8ff; border:1px solid #d6e4fb; border-left:5px solid {PRIMARY};
        border-radius:12px; padding:14px 16px; margin-bottom:10px; }}
.kpi .l {{ font-size:13px; color:#5b6b82; }}
.kpi .v {{ font-size:26px; font-weight:700; color:{PRIMARY}; }}
div.stButton > button {{ border-radius:10px; border:1px solid {PRIMARY}; color:{PRIMARY};
        background:#fff; font-weight:600; width:100%; padding:.6rem 0; }}
div.stButton > button:hover {{ background:{PRIMARY}; color:#fff; border-color:{PRIMARY}; }}
.hero {{ background:linear-gradient(120deg,#1f4e9c,#3b82d6); color:#fff; padding:34px;
        border-radius:16px; margin-bottom:22px; }}
.hero h1 {{ color:#fff; margin:0; }} .hero p {{ margin:6px 0 0; opacity:.9; font-size:18px; }}
</style>""", unsafe_allow_html=True)


# ----------------------------------------------------------------------------
# 3) DATA LAYER  (لو هتغيّر المصدر لاحقاً SQL/OneDrive غيّر الدالتين دول بس)
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes) -> list:
    """يرجع أسماء كل الشيتات من الملف."""
    return pd.ExcelFile(BytesIO(file_bytes)).sheet_names


@st.cache_data(show_spinner=False)
def load_excel_data(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    """يقرأ شيت واحد. الباقي في التطبيق ما يعرفش إن المصدر Excel."""
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name)
    df = df.dropna(how="all").dropna(axis=1, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    return df


# ----------------------------------------------------------------------------
# 4) HELPERS
# ----------------------------------------------------------------------------
def find_col(df, key):
    """يدوّر على العمود المناسب حسب COLUMN_MAP."""
    lower = {c.lower(): c for c in df.columns}
    for name in COLUMN_MAP.get(key, []):
        if name.lower() in lower:
            return lower[name.lower()]
    return None


def kpi(col, label, value):
    col.markdown(f'<div class="kpi"><div class="l">{label}</div><div class="v">{value}</div></div>',
                 unsafe_allow_html=True)


def kpi_row(items):
    """items = [(label, value), ...] ويتوزعوا على صفوف من 3."""
    for i in range(0, len(items), 3):
        cols = st.columns(3)
        for c, (l, v) in zip(cols, items[i:i + 3]):
            kpi(c, l, v)


def fmt(x):
    return f"{x:,.0f}" if float(x).is_integer() else f"{x:,.2f}"


def to_num(s):
    return pd.to_numeric(s, errors="coerce")


# ----------------------------------------------------------------------------
# 5) DELIVERY LOGIC
# ----------------------------------------------------------------------------
def compute_delivery_status(df, c):
    """منطق الحالة - سهل التعديل."""
    def status(row):
        if c["actual_date"] and pd.notna(row[c["actual_date"]]):
            return "Delivered"
        exp = row[c["expected_date"]] if c["expected_date"] else pd.NaT
        req = row[c["required_date"]] if c["required_date"] else pd.NaT
        if pd.notna(exp) and pd.notna(req) and exp > req:
            return "Delayed"
        if pd.notna(exp):
            return "Pending / On Track"
        return "Unknown"
    return df.apply(status, axis=1)


def show_delivery_dashboard(df, sheet_name):
    c = {k: find_col(df, k) for k in COLUMN_MAP}
    df = df.copy()
    for k in ("required_date", "expected_date", "actual_date"):
        if c[k]:
            df[c[k]] = pd.to_datetime(df[c[k]], errors="coerce")
    for k in ("required_qty", "delivered_qty", "remaining_qty"):
        if c[k]:
            df[c[k]] = to_num(df[c[k]])

    # --- تحذيرات الأعمدة الناقصة
    for k, msg in [("expected_date", "Expected delivery KPIs cannot be calculated."),
                   ("required_date", "Delay detection cannot be calculated."),
                   ("required_qty", "Remaining quantity cannot be calculated."),
                   ("project", "Project filter/chart is unavailable.")]:
        if not c[k]:
            st.warning(f"⚠ '{COLUMN_MAP[k][0]}' column was not found. {msg}")

    # --- Status
    if c["actual_date"] or c["expected_date"]:
        df["Computed Status"] = compute_delivery_status(df, c)
        status_col = "Computed Status"
    elif c["status"]:
        status_col = c["status"]
    else:
        status_col = None

    # --- Remaining
    if c["required_qty"] and c["delivered_qty"]:
        df["Remaining Qty (calc)"] = (df[c["required_qty"]] - df[c["delivered_qty"]].fillna(0)).clip(lower=0)
        rem_col = "Remaining Qty (calc)"
    else:
        rem_col = c["remaining_qty"]

    # --- Filters
    st.subheader("🔎 Filters")
    fcols = st.columns(5)
    for fc, (label, col) in zip(fcols, [("Project", c["project"]), ("Status", status_col),
                                        ("Item", c["item"]), ("Location", c["location"])]):
        if col:
            opts = sorted(df[col].dropna().astype(str).unique())
            sel = fc.multiselect(label, opts, key=f"f_{sheet_name}_{label}")
            if sel:
                df = df[df[col].astype(str).isin(sel)]
    if c["expected_date"] and df[c["expected_date"]].notna().any():
        dmin, dmax = df[c["expected_date"]].min().date(), df[c["expected_date"]].max().date()
        if dmin < dmax:
            rng = fcols[4].date_input("Expected Date", (dmin, dmax), key=f"d_{sheet_name}")
            if isinstance(rng, tuple) and len(rng) == 2:
                m = df[c["expected_date"]].dt.date.between(*rng) | df[c["expected_date"]].isna()
                df = df[m]

    # --- KPIs
    today = pd.Timestamp.today().normalize()
    week_start = today - pd.Timedelta(days=today.weekday())
    items = [("Total Deliveries", len(df))]
    if status_col:
        vc = df[status_col].astype(str).str.lower()
        items += [("Delivered", int(vc.str.contains("deliver").sum())),
                  ("Pending", int(vc.str.contains("pending|track").sum())),
                  ("Delayed", int(vc.str.contains("delay").sum()))]
    if rem_col:
        items.append(("Remaining Quantity", fmt(df[rem_col].sum())))
    if c["expected_date"]:
        e = df[c["expected_date"]]
        items.append(("Expected This Week", int(e.between(week_start, week_start + pd.Timedelta(days=6)).sum())))
        items.append(("Expected This Month", int(((e.dt.month == today.month) & (e.dt.year == today.year)).sum())))
    kpi_row(items)

    # --- Table
    st.subheader("📋 Delivery Table")
    show_cols = [x for x in [c["project"], c["item"], c["location"], c["required_qty"], c["delivered_qty"],
                             rem_col, c["required_date"], c["expected_date"], c["actual_date"], status_col] if x]
    st.dataframe(df[show_cols] if show_cols else df, width="stretch", hide_index=True)

    # --- Charts (فقط لو الأعمدة موجودة)
    st.subheader("📊 Charts")
    g1, g2 = st.columns(2)
    if status_col:
        cnt = df[status_col].value_counts().reset_index()
        cnt.columns = ["Status", "Count"]
        g1.plotly_chart(px.pie(cnt, names="Status", values="Count", hole=.45, title="Deliveries by Status",
                               color="Status", color_discrete_map=STATUS_COLORS), width="stretch")
    if c["expected_date"] and df[c["expected_date"]].notna().any():
        by = df.groupby(df[c["expected_date"]].dt.date).size().reset_index(name="Deliveries")
        g2.plotly_chart(px.bar(by, x=c["expected_date"], y="Deliveries", title="Expected Deliveries by Date",
                               color_discrete_sequence=[PRIMARY]), width="stretch")
    g3, g4 = st.columns(2)
    if c["delivered_qty"] and rem_col:
        tot = pd.DataFrame({"Type": ["Delivered", "Remaining"],
                            "Qty": [df[c["delivered_qty"]].sum(), df[rem_col].sum()]})
        g3.plotly_chart(px.bar(tot, x="Type", y="Qty", color="Type", title="Delivered vs Remaining Quantity",
                               color_discrete_sequence=["#2e9e5b", "#d9534f"]), width="stretch")
    if c["project"]:
        bp = df.groupby(c["project"]).size().reset_index(name="Deliveries")
        g4.plotly_chart(px.bar(bp, x=c["project"], y="Deliveries", title="Deliveries by Project",
                               color_discrete_sequence=[PRIMARY]), width="stretch")


# ----------------------------------------------------------------------------
# 6) GENERIC DASHBOARD  (يشتغل على أي شيت مهما كانت أعمدته)
# ----------------------------------------------------------------------------
def show_generic_dashboard(df, sheet_name):
    num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    cat_cols = [c for c in df.columns if df[c].dtype == object and 2 <= df[c].nunique() <= 20]

    # Filters
    st.subheader("🔎 Filters")
    search = st.text_input("Search", key=f"s_{sheet_name}")
    if search:
        mask = df.astype(str).apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)
        df = df[mask]
    fcols = st.columns(3)
    for fc, col in zip(fcols, cat_cols[:3]):
        sel = fc.multiselect(col, sorted(df[col].dropna().astype(str).unique()), key=f"g_{sheet_name}_{col}")
        if sel:
            df = df[df[col].astype(str).isin(sel)]

    # KPIs
    items = [("Total Records", len(df)), ("Columns", len(df.columns))]
    items += [(f"Total {c}", fmt(df[c].sum())) for c in num_cols[:3]]
    st_col = find_col(df, "status")
    if st_col:
        for label, word in [("Completed", "complet|deliver|done"), ("Pending", "pend"), ("Delayed", "delay")]:
            items.append((label, int(df[st_col].astype(str).str.lower().str.contains(word).sum())))
    kpi_row(items[:6])

    st.subheader("📋 Data")
    st.dataframe(df, width="stretch", hide_index=True)

    if cat_cols or num_cols:
        st.subheader("📊 Charts")
    ch = st.columns(2)
    if cat_cols:
        vc = df[cat_cols[0]].value_counts().reset_index()
        vc.columns = [cat_cols[0], "Count"]
        ch[0].plotly_chart(px.bar(vc, x=cat_cols[0], y="Count", title=f"Records by {cat_cols[0]}",
                                  color_discrete_sequence=[PRIMARY]), width="stretch")
    if cat_cols and num_cols:
        agg = df.groupby(cat_cols[0])[num_cols[0]].sum().reset_index()
        ch[1].plotly_chart(px.pie(agg, names=cat_cols[0], values=num_cols[0], hole=.45,
                                  title=f"{num_cols[0]} by {cat_cols[0]}"), width="stretch")


def show_dashboard(df, sheet_name):
    """القالب الموحد: كل شيت يعدي من هنا."""
    st.title(f"{sheet_name} Dashboard")
    if df.empty:
        st.warning("This sheet is empty.")
        return
    is_delivery = "deliver" in sheet_name.lower() or (
        find_col(df, "expected_date") and find_col(df, "required_date"))
    (show_delivery_dashboard if is_delivery else show_generic_dashboard)(df, sheet_name)


# ----------------------------------------------------------------------------
# 7) NAVIGATION / STATE
# ----------------------------------------------------------------------------
ss = st.session_state
ss.setdefault("page", "home")
ss.setdefault("file_bytes", None)
ss.setdefault("file_name", None)
ss.setdefault("sheet", None)


def go_home():
    ss.page = "home"


def open_sheet(name):
    ss.sheet, ss.page = name, "sheet"


def open_named(wanted):
    """للأزرار الثابتة Projects/Deliveries: يدوّر على شيت بنفس الاسم."""
    names = get_sheet_names(ss.file_bytes) if ss.file_bytes else []
    match = next((n for n in names if n.strip().lower() == wanted.lower()), None)
    if match:
        open_sheet(match)
    else:
        st.warning(f"No sheet named '{wanted}' in the uploaded file.")


with st.sidebar:
    st.header("🏗️ Navigation")
    up = st.file_uploader("Upload Excel File", type=["xlsx"])
    if up is not None and (up.name != ss.file_name or up.size != len(ss.file_bytes or b"")):
        ss.file_bytes, ss.file_name, ss.sheet, ss.page = up.getvalue(), up.name, None, "home"
        st.cache_data.clear()
    st.button("🏠 Home", on_click=go_home, width="stretch")
    if ss.file_bytes:
        st.button("📁 Projects", on_click=open_named, args=("Projects",), width="stretch")
        st.button("🚚 Deliveries", on_click=open_named, args=("Deliveries",), width="stretch")
        if ss.sheet:
            st.button(f"📄 Current: {ss.sheet}", on_click=open_sheet, args=(ss.sheet,), width="stretch")
        if st.button("🔄 Refresh Data", width="stretch"):
            st.cache_data.clear()
            st.rerun()

# ----------------------------------------------------------------------------
# 8) PAGES
# ----------------------------------------------------------------------------
try:
    if ss.page == "home" or not ss.file_bytes:
        st.markdown('<div class="hero"><h1>Construction Projects Dashboard</h1>'
                    '<p>Monitor projects, deliveries, materials and construction activities.</p></div>',
                    unsafe_allow_html=True)
        if not ss.file_bytes:
            st.info("⬅ Upload an Excel (.xlsx) file from the sidebar to start.")
        else:
            st.success(f"Loaded: {ss.file_name}")
            b1, b2 = st.columns(2)
            b1.button("📁 Projects", on_click=open_named, args=("Projects",), key="h_p")
            b2.button("🚚 Deliveries", on_click=open_named, args=("Deliveries",), key="h_d")

            # ---- الأزرار الديناميكية: زر لكل شيت
            names = get_sheet_names(ss.file_bytes)
            st.subheader(f"Excel Sheets ({len(names)})")
            per_row = 4
            for i in range(0, len(names), per_row):
                cols = st.columns(per_row)
                for j, name in enumerate(names[i:i + per_row]):
                    cols[j].button(name, key=f"sheet_btn_{i + j}", on_click=open_sheet, args=(name,))
    else:
        st.caption(f"📂 {ss.file_name}")
        df = load_excel_data(ss.file_bytes, ss.sheet)
        show_dashboard(df, ss.sheet)
except Exception as e:  # ملف تالف أو شيت غير صالح
    st.error(f"Could not read the Excel file or sheet: {e}")
    st.button("Back to Home", on_click=go_home)
