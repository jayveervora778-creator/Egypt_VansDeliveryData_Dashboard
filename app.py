
import os, io, tempfile, datetime
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.io as pio
from pivottablejs import pivot_ui
import streamlit.components.v1 as components
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="Vans Interactive Dashboard", layout="wide")
APP_TITLE = "📊 Vans Data Interactive Dashboard"

# ---------- Auth (optional) ----------
PASSWORD = os.getenv("STREAMLIT_DASH_PASSWORD", "")
if PASSWORD:
    def login():
        with st.form("login", clear_on_submit=False):
            pwd = st.text_input("Password", type="password")
            ok = st.form_submit_button("Enter")
        return ok, pwd
    st.title(APP_TITLE)
    ok, pwd = login()
    if not ok or pwd != PASSWORD:
        st.info("Enter the password to access the dashboard.")
        st.stop()

st.title(APP_TITLE)
st.caption("Slice, dice, visualize, and export Vans dataset.")

DEFAULT_PATH = "Vans data for dashboard.xlsx"

def _flatten_columns(columns):
    flattened = []
    for col in columns:
        if isinstance(col, tuple) or isinstance(col, list):
            parts = [str(x) for x in col if str(x) != "nan"]
            flattened.append(" - ".join(parts).strip())
        else:
            flattened.append(str(col).strip())
    return flattened

@st.cache_data(show_spinner=False)
def load_excel_any(path_or_bytes):
    xls = pd.ExcelFile(path_or_bytes)
    frames = []
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(path_or_bytes, sheet_name=sheet, header=[0,1,2])
            df.columns = _flatten_columns(df.columns.values)
        except Exception:
            df = pd.read_excel(path_or_bytes, sheet_name=sheet, header=0)
            df.columns = _flatten_columns(df.columns.values)
        df.columns = [c.replace("Unnamed: ", "").strip() for c in df.columns]
        frames.append(df)
    return pd.concat(frames, ignore_index=True)

choice = st.radio("Data source:", ["Use included file", "Upload your own"], horizontal=True)

if choice == "Use included file":
    if not os.path.exists(DEFAULT_PATH):
        st.error("Included file not found.")
        st.stop()
    df_all = load_excel_any(DEFAULT_PATH)
else:
    up = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    if up is None:
        st.stop()
    else:
        data = io.BytesIO(up.read())
        df_all = load_excel_any(data)

# Clean cols
df_all.columns = [str(c).strip() for c in df_all.columns]
df_view = df_all.copy()

# ---------- Global Filters ----------
with st.sidebar:
    st.header("🔎 Global Filters")
    for c in ["Company", "Employment Status", "Areas Covered"]:
        if c in df_all.columns:
            vals = sorted([v for v in df_all[c].dropna().unique()])
            sel = st.multiselect(f"{c} filter", options=vals, default=vals)
            df_view = df_view[df_view[c].isin(sel)]
    if "Age (Years)" in df_all.columns:
        min_age, max_age = int(df_all["Age (Years)"].min()), int(df_all["Age (Years)"].max())
        age_range = st.slider("Age Range", min_age, max_age, (min_age, max_age))
        df_view = df_view[df_view["Age (Years)"].between(age_range[0], age_range[1])]

# ---------- KPI Metrics ----------
st.subheader("📌 Key Metrics")
col1, col2, col3, col4 = st.columns(4)
if "Deliveries per day" in df_view.columns:
    col1.metric("Avg Deliveries/day", round(df_view["Deliveries per day"].mean(),1))
if "Medical Insurance" in df_view.columns:
    pct = (df_view["Medical Insurance"].eq("Yes").mean()*100)
    col2.metric("% with Medical Insurance", f"{pct:.1f}%")
if "Net Income (Gross - All Expenses) (EGP)" in df_view.columns:
    col3.metric("Avg Net Income", f"{df_view['Net Income (Gross - All Expenses) (EGP)'].mean():,.0f} EGP")
if "Employment Status" in df_view.columns:
    mode = df_view["Employment Status"].mode()[0] if not df_view["Employment Status"].empty else "N/A"
    col4.metric("Most Common Employment", mode)

# ---------- Preset Pivots ----------
st.subheader("🔖 Quick Pivot Presets")
preset = st.selectbox("Select a preset:", [
    "None",
    "Employment Status × Benefits",
    "Company × Avg Deliveries",
    "Areas Covered × Avg Net Income",
    "Company × Bicycle Ownership",
    "Employment Status × Overtime Pay",
    "Company × Ramadan Incentives"
])
if preset != "None":
    st.info(f"Preset '{preset}' loaded. Adjust further in the pivot below.")

# ---------- Pivot ----------
st.subheader("📊 Pivot Table")
try:
    pivot_ui(df_view, outfile_path="pivottable.html")
    with open("pivottable.html", "r", encoding="utf-8") as f:
        html = f.read()
    components.html(html, height=800, scrolling=True)
except Exception as e:
    st.error(f"Pivot error: {e}")

# ---------- Visual Analytics ----------
st.subheader("📈 Visual Analytics")

charts = []

if "Employment Status" in df_view.columns:
    fig = px.pie(df_view, names="Employment Status", title="Employment Status Share")
    st.plotly_chart(fig, use_container_width=True)
    charts.append(("pie.png", fig))

if "Company" in df_view.columns and "Deliveries per day" in df_view.columns:
    fig = px.bar(df_view, x="Company", y="Deliveries per day",
                 title="Deliveries per Day by Company", barmode="group")
    st.plotly_chart(fig, use_container_width=True)
    charts.append(("deliveries.png", fig))

if "Age (Years)" in df_view.columns:
    fig = px.histogram(df_view, x="Age (Years)", nbins=10, title="Age Distribution")
    st.plotly_chart(fig, use_container_width=True)
    charts.append(("age.png", fig))

if "Net Income (Gross - All Expenses) (EGP)" in df_view.columns and "Employment Status" in df_view.columns:
    fig = px.box(df_view, x="Employment Status",
                 y="Net Income (Gross - All Expenses) (EGP)", title="Net Income by Employment Status")
    st.plotly_chart(fig, use_container_width=True)
    charts.append(("income.png", fig))

if "Fuel Expenses (EGP)" in df_view.columns and "Company" in df_view.columns:
    df_exp = df_view.groupby("Company")[["Fuel Expenses (EGP)", "Maintenance Costs (EGP)",
                                         "Financing/Lease (EGP)", "Other Expenses (licenses, permits, fines, etc....)"]].mean().reset_index()
    df_exp = df_exp.melt(id_vars="Company", var_name="Expense Type", value_name="Avg Expense")
    fig = px.bar(df_exp, x="Company", y="Avg Expense", color="Expense Type", barmode="stack",
                 title="Average Expenses by Company")
    st.plotly_chart(fig, use_container_width=True)
    charts.append(("expenses.png", fig))

# ---------- PDF Export ----------
st.subheader("📑 Export")
if st.button("⬇️ Export Dashboard to PDF"):
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = os.path.join(tmpdir, "dashboard_snapshot.pdf")
        doc = SimpleDocTemplate(pdf_path, pagesize=A4)
        styles = getSampleStyleSheet()
        flow = []

        flow.append(Paragraph("Vans Data Dashboard Snapshot", styles['Title']))
        flow.append(Paragraph(datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), styles['Normal']))
        flow.append(Spacer(1,12))

        # KPIs
        if "Deliveries per day" in df_view.columns:
            flow.append(Paragraph(f"Avg Deliveries/day: {round(df_view['Deliveries per day'].mean(),1)}", styles['Normal']))
        if "Medical Insurance" in df_view.columns:
            pct = df_view["Medical Insurance"].eq("Yes").mean()*100
            flow.append(Paragraph(f"% with Medical Insurance: {pct:.1f}%", styles['Normal']))
        if "Net Income (Gross - All Expenses) (EGP)" in df_view.columns:
            flow.append(Paragraph(f"Avg Net Income: {df_view['Net Income (Gross - All Expenses) (EGP)'].mean():,.0f} EGP", styles['Normal']))
        flow.append(Spacer(1,12))

        # Save and add charts
        for fname, fig in charts:
            outpath = os.path.join(tmpdir, fname)
            pio.write_image(fig, outpath, format="png")
            flow.append(Image(outpath, width=400, height=250))
            flow.append(Spacer(1,12))

        doc.build(flow)
        with open(pdf_path,"rb") as f:
            st.download_button("Download PDF", data=f, file_name="dashboard_snapshot.pdf", mime="application/pdf")
