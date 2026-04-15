import streamlit as st
import pandas as pd
from datetime import timedelta, datetime
from io import BytesIO

st.set_page_config(page_title="Comité Paritaire QC", layout="wide")
st.title("Comité Paritaire Québec - Weekly Report")

# =========================
# Helpers
# =========================
def clean_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df

def normalize_work_type(value):
    v = str(value).strip().upper()

    if "REGULAR" in v:
        return "Regular"
    elif "SUPPL" in v:
        return "Suppl."
    elif "CONGE TRAVAIL" in v or "CONGÉ TRAVAIL" in v:
        return "Congé Travaillé"
    elif v == "CONGE" or v == "CONGÉ" or "CONGE " in v or "CONGÉ " in v:
        return "Congé"
    elif "MALAD" in v:
        return "Maladie"
    else:
        return "Other"

def assign_committee_week(date_value, start_date, num_weeks=6):
    for i in range(num_weeks):
        week_start = start_date + timedelta(days=i * 7)
        week_end = week_start + timedelta(days=6)
        if week_start <= date_value <= week_end:
            return week_start, week_end
    return None, None

def load_uploaded_files(uploaded_files):
    dataframes = []

    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file, sheet_name="data")
            df = clean_columns(df)
            df["source_file"] = uploaded_file.name
            dataframes.append(df)
        except Exception as e:
            st.warning(f"No se pudo leer {uploaded_file.name}: {e}")

    if dataframes:
        return pd.concat(dataframes, ignore_index=True)
    return pd.DataFrame()

def to_excel_report(detail_df, summary_df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        detail_df.to_excel(writer, index=False, sheet_name="Data_Filtered")
        summary_df.to_excel(writer, index=False, sheet_name="Summary")

    output.seek(0)
    return output

# =========================
# Upload section
# =========================
st.subheader("1. Upload Excel files")
uploaded_files = st.file_uploader(
    "Upload one or more Excel files",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("Upload your Excel source files to begin.")
    st.stop()

df = load_uploaded_files(uploaded_files)

if df.empty:
    st.error("No valid data could be loaded from sheet 'data'.")
    st.stop()

# =========================
# Column mapping
# =========================
column_map = {
    "Date": "date",
    "Province": "province",
    "name employee": "employee",
    "total hours worked (numb)": "hours",
    "total hours worked (numb.)": "hours",
    "total hours worked (numb...)": "hours",
    "total hours worked": "hours",
    "Total to pay": "total_pay",
    "Type of work": "type_of_work",
    "Vendor Company": "vendor_company"
}

df = df.rename(columns=column_map)

required_cols = ["date", "province", "employee", "hours", "total_pay", "type_of_work", "vendor_company"]
missing = [c for c in required_cols if c not in df.columns]

if missing:
    st.error(f"Missing required columns: {missing}")
    st.stop()

# =========================
# Data cleaning
# =========================
df["date"] = pd.to_datetime(df["date"], errors="coerce")
df["province"] = df["province"].astype(str).str.strip().str.upper()
df["employee"] = df["employee"].astype(str).str.strip()
df["vendor_company"] = df["vendor_company"].astype(str).str.strip()
df["type_of_work"] = df["type_of_work"].astype(str).str.strip()
df["hours"] = pd.to_numeric(df["hours"], errors="coerce").fillna(0)
df["total_pay"] = pd.to_numeric(df["total_pay"], errors="coerce").fillna(0)

df = df.dropna(subset=["date"])

# only QC
df = df[df["province"] == "QC"].copy()

# normalize work type
df["work_class"] = df["type_of_work"].apply(normalize_work_type)

# =========================
# Sidebar filters
# =========================
st.sidebar.header("Filters")

default_start = datetime(2026, 1, 4).date()
start_date = st.sidebar.date_input("First committee week start date", value=default_start)
num_weeks = st.sidebar.number_input("Number of weeks", min_value=1, max_value=12, value=4)

vendors = sorted([v for v in df["vendor_company"].dropna().unique().tolist() if v])
selected_vendors = st.sidebar.multiselect("Vendor Company", vendors, default=vendors)

employees = sorted([e for e in df["employee"].dropna().unique().tolist() if e])
selected_employees = st.sidebar.multiselect("Employee", employees, default=employees)

work_types = ["Regular", "Suppl.", "Congé", "Congé Travaillé", "Maladie", "Other"]
selected_work_types = st.sidebar.multiselect("Work Class", work_types, default=work_types)

# =========================
# Week assignment
# =========================
start_date_dt = pd.to_datetime(start_date)

df[["week_start", "week_end"]] = df["date"].apply(
    lambda x: pd.Series(assign_committee_week(x, start_date_dt, num_weeks))
)

df = df[df["week_start"].notna()].copy()
df = df[df["vendor_company"].isin(selected_vendors)]
df = df[df["employee"].isin(selected_employees)]
df = df[df["work_class"].isin(selected_work_types)]

df["week_label"] = df["week_end"].dt.strftime("%Y-%m-%d")

# =========================
# Summary
# =========================
summary = (
    df.groupby(["vendor_company", "employee", "week_label", "work_class"], dropna=False)
      .agg(
          total_hours=("hours", "sum"),
          total_pay=("total_pay", "sum")
      )
      .reset_index()
      .sort_values(["vendor_company", "employee", "week_label", "work_class"])
)

# =========================
# Screen output
# =========================
st.subheader("2. Filtered source data")
st.dataframe(df, use_container_width=True)

st.subheader("3. Weekly summary")
st.dataframe(summary, use_container_width=True)

# KPIs
col1, col2, col3 = st.columns(3)
col1.metric("Rows", len(df))
col2.metric("Total Hours", f"{df['hours'].sum():,.2f}")
col3.metric("Total Pay", f"${df['total_pay'].sum():,.2f}")

# =========================
# Download
# =========================
excel_file = to_excel_report(df, summary)

st.download_button(
    label="Download Excel Report",
    data=excel_file,
    file_name="comite_paritaire_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
