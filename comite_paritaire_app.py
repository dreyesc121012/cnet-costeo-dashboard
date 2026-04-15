import streamlit as st
import pandas as pd
import requests
import msal
import base64
from datetime import timedelta, datetime
from io import BytesIO

st.set_page_config(page_title="Comité Paritaire QC", layout="wide")
st.title("Comité Paritaire Québec - Weekly Report")

# =========================
# Secrets / Config
# =========================
CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
TENANT_ID = st.secrets["TENANT_ID"]
ONEDRIVE_FOLDER_URL = st.secrets["ONEDRIVE_FOLDER_URL"]

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

# =========================
# Microsoft Graph helpers
# =========================
def get_access_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=SCOPES)

    if "access_token" not in result:
        raise Exception(f"Could not obtain access token: {result}")

    return result["access_token"]


def encode_share_url(url: str) -> str:
    encoded = base64.b64encode(url.encode("utf-8")).decode("utf-8")
    encoded = encoded.rstrip("=").replace("/", "_").replace("+", "-")
    return f"u!{encoded}"


def resolve_folder_from_share_url(token: str, folder_url: str) -> dict:
    share_id = encode_share_url(folder_url)
    endpoint = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem"
    headers = {"Authorization": f"Bearer {token}"}

    response = requests.get(endpoint, headers=headers, timeout=60)
    response.raise_for_status()
    return response.json()


def list_folder_files(token: str, drive_id: str, item_id: str) -> list[dict]:
    endpoint = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children"
    headers = {"Authorization": f"Bearer {token}"}

    response = requests.get(endpoint, headers=headers, timeout=60)
    response.raise_for_status()
    data = response.json()

    files = []
    for item in data.get("value", []):
        name = item.get("name", "")
        if name.lower().endswith((".xlsx", ".xlsm")):
            files.append({
                "name": name,
                "id": item["id"],
                "drive_id": drive_id,
            })

    files = sorted(files, key=lambda x: x["name"].lower())
    return files


def download_file_bytes(token: str, drive_id: str, item_id: str) -> BytesIO:
    endpoint = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    headers = {"Authorization": f"Bearer {token}"}

    response = requests.get(endpoint, headers=headers, timeout=120)
    response.raise_for_status()
    return BytesIO(response.content)

# =========================
# Data helpers
# =========================
def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    return df


def normalize_work_type(value: str) -> str:
    v = str(value).strip().upper()

    if "REGULAR" in v:
        return "Regular"
    if "SUPPL" in v:
        return "Suppl."
    if "CONGE TRAVAIL" in v or "CONGÉ TRAVAIL" in v:
        return "Congé Travaillé"
    if v == "CONGE" or v == "CONGÉ" or "CONGE " in v or "CONGÉ " in v:
        return "Congé"
    if "MALAD" in v:
        return "Maladie"
    return "Other"


def assign_committee_week(date_value: pd.Timestamp, start_date: pd.Timestamp, num_weeks: int = 6):
    for i in range(num_weeks):
        week_start = start_date + timedelta(days=i * 7)
        week_end = week_start + timedelta(days=6)
        if week_start <= date_value <= week_end:
            return week_start, week_end
    return None, None


def load_selected_cloud_files(token: str, selected_files: list[dict]) -> pd.DataFrame:
    dataframes = []

    for file_info in selected_files:
        try:
            file_bytes = download_file_bytes(token, file_info["drive_id"], file_info["id"])
            df = pd.read_excel(file_bytes, sheet_name="data")
            df = clean_columns(df)
            df["source_file"] = file_info["name"]
            dataframes.append(df)
        except Exception as e:
            st.warning(f"Could not read {file_info['name']}: {e}")

    if dataframes:
        return pd.concat(dataframes, ignore_index=True)

    return pd.DataFrame()


def to_excel_report(detail_df: pd.DataFrame, summary_df: pd.DataFrame) -> BytesIO:
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        detail_df.to_excel(writer, index=False, sheet_name="Data_Filtered")
        summary_df.to_excel(writer, index=False, sheet_name="Summary")

    output.seek(0)
    return output

# =========================
# Connect to cloud folder
# =========================
st.subheader("1. Read Excel files from OneDrive folder")

try:
    token = get_access_token()
    folder_info = resolve_folder_from_share_url(token, ONEDRIVE_FOLDER_URL)

    drive_id = folder_info["parentReference"]["driveId"]
    item_id = folder_info["id"]

    cloud_files = list_folder_files(token, drive_id, item_id)

    if not cloud_files:
        st.warning("No Excel files were found in the OneDrive folder.")
        st.stop()

except Exception as e:
    st.error(f"Error connecting to OneDrive folder: {e}")
    st.stop()

file_names = [f["name"] for f in cloud_files]

selected_names = st.multiselect(
    "Select files from OneDrive folder",
    file_names,
    default=file_names
)

if not selected_names:
    st.info("Select at least one Excel file.")
    st.stop()

selected_files = [f for f in cloud_files if f["name"] in selected_names]

with st.expander("Files found in OneDrive folder", expanded=False):
    st.write(file_names)

df = load_selected_cloud_files(token, selected_files)

if df.empty:
    st.error("No valid data could be loaded from the selected files.")
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
    "Vendor Company": "vendor_company",
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
df = df[df["province"] == "QC"].copy()
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
# Assign weeks and filter
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
# Output
# =========================
st.subheader("2. Filtered source data")
st.dataframe(df, use_container_width=True)

st.subheader("3. Weekly summary")
st.dataframe(summary, use_container_width=True)

col1, col2, col3 = st.columns(3)
col1.metric("Rows", len(df))
col2.metric("Total Hours", f"{df['hours'].sum():,.2f}")
col3.metric("Total Pay", f"${df['total_pay'].sum():,.2f}")

excel_file = to_excel_report(df, summary)

st.download_button(
    label="Download Excel Report",
    data=excel_file,
    file_name="comite_paritaire_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
