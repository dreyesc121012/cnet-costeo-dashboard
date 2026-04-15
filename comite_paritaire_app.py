import base64
from io import BytesIO
from datetime import timedelta, datetime

import pandas as pd
import requests
import streamlit as st
import msal

# ============================================================
# PAGE
# ============================================================
st.set_page_config(page_title="Comité Paritaire QC", layout="wide")
st.title("Comité Paritaire Québec - Weekly Report")

# ============================================================
# CONFIG (SECRETS)
# Required in Streamlit Secrets:
#
# CLIENT_ID = "..."
# CLIENT_SECRET = "..."
# TENANT_ID = "..."
# REDIRECT_URI = "https://comite-paritaire.streamlit.app"
# ONEDRIVE_FOLDER_URL = "https://...shared folder link of 2026 root folder..."
# ALLOWED_DOMAIN = "groupcastillo.com"
# DOMAIN_HINT = "groupcastillo.com"
# LOGIN_HINT = ""
# ============================================================
CLIENT_ID = str(st.secrets["CLIENT_ID"]).strip()
CLIENT_SECRET = str(st.secrets["CLIENT_SECRET"]).strip()
TENANT_ID = str(st.secrets["TENANT_ID"]).strip()
REDIRECT_URI = str(st.secrets["REDIRECT_URI"]).strip().rstrip("/")
ONEDRIVE_FOLDER_URL = str(st.secrets["ONEDRIVE_FOLDER_URL"]).strip()
ALLOWED_DOMAIN = str(st.secrets.get("ALLOWED_DOMAIN", "")).strip().lower()
DOMAIN_HINT = str(st.secrets.get("DOMAIN_HINT", "")).strip()
LOGIN_HINT = str(st.secrets.get("LOGIN_HINT", "")).strip()

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.Read.All", "Sites.Read.All"]

# ============================================================
# QUERY PARAM HELPERS
# ============================================================
def get_query_params_compat() -> dict:
    try:
        qp = st.query_params
        out = {}
        for k in qp.keys():
            v = qp.get(k)
            if isinstance(v, list):
                out[k] = v[0] if v else ""
            else:
                out[k] = str(v) if v is not None else ""
        return out
    except Exception:
        try:
            qp = st.experimental_get_query_params()
            out = {}
            for k, v in qp.items():
                if isinstance(v, list):
                    out[k] = v[0] if v else ""
                else:
                    out[k] = str(v) if v is not None else ""
            return out
        except Exception:
            return {}

def clear_query_params_compat():
    try:
        st.query_params.clear()
    except Exception:
        try:
            st.experimental_set_query_params()
        except Exception:
            pass

# ============================================================
# MSAL
# ============================================================
def get_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        token_cache=None,
    )

# ============================================================
# GRAPH HELPERS
# ============================================================
def graph_get(url: str, access_token: str) -> requests.Response:
    return requests.get(
        url,
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=60,
    )

def graph_get_json(url: str, access_token: str) -> dict:
    r = graph_get(url, access_token)
    if r.status_code != 200:
        raise RuntimeError(f"Graph error {r.status_code}\n{r.text}")
    return r.json()

def get_me(access_token: str) -> dict:
    r = requests.get(
        "https://graph.microsoft.com/v1.0/me",
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=60,
    )
    if r.status_code != 200:
        raise RuntimeError(f"Graph /me error {r.status_code}\n{r.text}")
    return r.json()

def get_user_email(me: dict) -> str:
    return (me.get("mail") or me.get("userPrincipalName") or "").strip().lower()

def is_allowed_user(me: dict) -> bool:
    email = get_user_email(me)
    if not ALLOWED_DOMAIN:
        return False
    return email.endswith(f"@{ALLOWED_DOMAIN}")

def make_share_id(shared_url: str) -> str:
    b = base64.b64encode(shared_url.encode("utf-8")).decode("utf-8")
    b = b.rstrip("=").replace("/", "_").replace("+", "-")
    return "u!" + b

def resolve_shared_link(access_token: str, shared_url: str) -> dict:
    share_id = make_share_id(shared_url)
    meta_url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem"
    meta = graph_get(meta_url, access_token)
    if meta.status_code != 200:
        raise RuntimeError(
            f"Error resolving shared link: {meta.status_code}\n{meta.text}\n\n"
            "TIP: Use SharePoint/OneDrive → Share → Copy link (within your organization)."
        )
    return meta.json()

def download_item_bytes(access_token: str, drive_id: str, item_id: str) -> bytes:
    content_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    r = graph_get(content_url, access_token)
    if r.status_code != 200:
        raise RuntimeError(f"Error downloading file: {r.status_code}\n{r.text}")
    return r.content

def list_children_all(access_token: str, drive_id: str, folder_item_id: str):
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_item_id}/children?$top=200"
    all_items = []
    while url:
        data = graph_get_json(url, access_token)
        all_items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return all_items

def is_excel_name(name: str) -> bool:
    n = (name or "").lower()
    return n.endswith(".xlsx") or n.endswith(".xlsm") or n.endswith(".xls")

def is_folder_item(item: dict) -> bool:
    return "folder" in item

# ============================================================
# DATA HELPERS
# ============================================================
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

def safe_text_series(s: pd.Series) -> pd.Series:
    out = s.astype(str).str.replace("\u00A0", " ", regex=False).str.strip()
    return out.replace(
        {
            "nan": "",
            "None": "",
            "none": "",
            "NULL": "",
            "null": "",
            "<NA>": "",
        }
    ).fillna("")

def load_selected_excel_files(access_token: str, drive_id: str, selected_files: list[dict], month_name_map: dict) -> pd.DataFrame:
    dfs = []

    for file_info in selected_files:
        try:
            file_bytes = download_item_bytes(access_token, drive_id, file_info["id"])
            excel_file = pd.ExcelFile(BytesIO(file_bytes))

            sheet_to_use = None
            for s in excel_file.sheet_names:
                if s.strip().lower() == "data":
                    sheet_to_use = s
                    break

            if sheet_to_use is None:
                st.warning(
                    f"Could not read {file_info['name']}: sheet 'Data' not found. "
                    f"Available sheets: {excel_file.sheet_names}"
                )
                continue

            df = excel_file.parse(sheet_to_use)
            df = clean_columns(df)
            df["source_file"] = file_info["name"]
            df["source_month_folder"] = month_name_map.get(file_info["id"], "")
            dfs.append(df)

        except Exception as e:
            st.warning(f"Could not read {file_info['name']}: {e}")

    if dfs:
        return pd.concat(dfs, ignore_index=True)

    return pd.DataFrame()

def to_excel_report(detail_df: pd.DataFrame, summary_df: pd.DataFrame) -> BytesIO:
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        detail_df.to_excel(writer, index=False, sheet_name="Data_Filtered")
        summary_df.to_excel(writer, index=False, sheet_name="Summary")

    output.seek(0)
    return output

# ============================================================
# AUTH FLOW
# ============================================================
if not ALLOWED_DOMAIN:
    st.error("Missing ALLOWED_DOMAIN in Streamlit secrets.")
    st.stop()

app = get_msal_app()
params = get_query_params_compat()

if "token_result" not in st.session_state:
    code = params.get("code")

    if code:
        result = app.acquire_token_by_authorization_code(
            code=code,
            scopes=SCOPES,
            redirect_uri=REDIRECT_URI,
        )

        if "access_token" in result:
            st.session_state.token_result = result
            clear_query_params_compat()
            st.rerun()
        else:
            st.error("Could not obtain access token.")
            st.code(str(result))
            st.stop()

    st.warning("You are not signed in to Microsoft 365 / SharePoint.")

    extra_qp = {}
    if DOMAIN_HINT:
        extra_qp["domain_hint"] = DOMAIN_HINT
    if LOGIN_HINT:
        extra_qp["login_hint"] = LOGIN_HINT

    auth_url = app.get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
        prompt="select_account",
        response_mode="query",
        extra_query_parameters=extra_qp,
    )

    st.link_button("🔐 Sign in with Microsoft (Company)", auth_url)
    st.caption(f"Redirect URI used: {REDIRECT_URI}")
    st.stop()

token_result = st.session_state.token_result
access_token = token_result.get("access_token", "")

if not access_token:
    st.error("No access token found. Please sign in again.")
    st.session_state.pop("token_result", None)
    st.stop()

try:
    me = get_me(access_token)
    signed_in_email = get_user_email(me)
except Exception as e:
    st.error("Could not validate signed-in user.")
    st.code(str(e))
    st.session_state.pop("token_result", None)
    st.stop()

if not is_allowed_user(me):
    st.error("Access denied. This dashboard is restricted to company users only.")
    st.write("Signed in as:", signed_in_email if signed_in_email else "(unknown user)")
    st.session_state.pop("token_result", None)
    st.stop()

# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.success(f"Logged in as {signed_in_email}")
st.success(f"✅ Signed in as {signed_in_email}")

if st.button("🚪 Sign out"):
    st.session_state.pop("token_result", None)
    clear_query_params_compat()
    st.rerun()

st.sidebar.header("📁 SharePoint Source")
st.sidebar.caption("Select month folder(s), then choose Excel files inside them.")

# ============================================================
# RESOLVE ROOT FOLDER (2026)
# ============================================================
try:
    meta = resolve_shared_link(access_token, ONEDRIVE_FOLDER_URL)
except Exception as e:
    st.error("Could not resolve the SharePoint/OneDrive folder link.")
    st.code(str(e))
    st.stop()

drive_id = meta["parentReference"]["driveId"]
root_item_id = meta["id"]

if "folder" not in meta:
    st.error("ONEDRIVE_FOLDER_URL must be a folder link.")
    st.stop()

# ============================================================
# LIST MONTH FOLDERS
# ============================================================
try:
    root_children = list_children_all(access_token, drive_id, root_item_id)
except Exception as e:
    st.error("Could not list folders from the root folder.")
    st.code(str(e))
    st.stop()

month_folders = [x for x in root_children if is_folder_item(x)]
month_folders.sort(key=lambda x: (x.get("name") or "").lower())

if not month_folders:
    st.warning("No month folders were found inside the selected root folder.")
    st.stop()

month_folder_names = [f["name"] for f in month_folders]

selected_month_names = st.sidebar.multiselect(
    "Select month folder(s)",
    month_folder_names,
    default=month_folder_names[:1] if month_folder_names else [],
)

if not selected_month_names:
    st.info("Please select at least one month folder.")
    st.stop()

selected_month_folders = [f for f in month_folders if f["name"] in selected_month_names]

# ============================================================
# LIST EXCEL FILES INSIDE SELECTED MONTH FOLDERS
# ============================================================
all_excel_files = []
month_name_map = {}

for folder_info in selected_month_folders:
    folder_id = folder_info["id"]
    folder_name = folder_info["name"]

    try:
        children = list_children_all(access_token, drive_id, folder_id)
    except Exception as e:
        st.warning(f"Could not list files inside '{folder_name}': {e}")
        continue

    excels = [x for x in children if is_excel_name(x.get("name", ""))]
    excels.sort(key=lambda x: (x.get("name") or "").lower())

    for item in excels:
        item_copy = dict(item)
        display_name = f"{folder_name} | {item_copy['name']}"
        item_copy["display_name"] = display_name
        all_excel_files.append(item_copy)
        month_name_map[item_copy["id"]] = folder_name

all_excel_files.sort(key=lambda x: x["display_name"].lower())

if not all_excel_files:
    st.warning("No Excel files found inside the selected month folder(s).")
    st.stop()

excel_display_names = [f["display_name"] for f in all_excel_files]

selected_excel_display_names = st.sidebar.multiselect(
    "Select Excel file(s)",
    excel_display_names,
    default=excel_display_names,
)

if not selected_excel_display_names:
    st.info("Please select at least one Excel file.")
    st.stop()

selected_excel_files = [f for f in all_excel_files if f["display_name"] in selected_excel_display_names]

with st.expander("Selected source overview", expanded=False):
    st.write("Month folders:", selected_month_names)
    st.write("Excel files:", selected_excel_display_names)

# ============================================================
# LOAD EXCEL DATA
# ============================================================
df = load_selected_excel_files(access_token, drive_id, selected_excel_files, month_name_map)

if df.empty:
    st.error("No valid data could be loaded from the selected Excel files.")
    st.stop()

# ============================================================
# COLUMN MAPPING
# Adjust here if your source headers vary
# ============================================================
column_map = {
    # DATE
    "date": "date",
    "Date": "date",

    # EMPLOYEE
    "employee": "employee",
    "Employee": "employee",
    "name employee": "employee",
    "Name Employee": "employee",
    "Name Employee & vendor company": "employee",

    # PROVINCE
    "province": "province",
    "Province": "province",

    # HOURS (🔥 FIX AQUÍ)
    "total hours worked (number)": "hours",
    "total hours worked (number )": "hours",
    "total hours worked (number)": "hours",
    "total hours worked (number)": "hours",
    "total hours worked (number)": "hours",
    "total hours worked (number)": "hours",
    "total hours worked (number)": "hours",
    "total hours worked (number)": "hours",
    "total hours worked (number)": "hours",
    "total hours worked (number)": "hours",
    "total hours worked": "hours",
    "Total Hours Worked": "hours",

    # PAY
    "total_pay": "total_pay",
    "Total Pay": "total_pay",
    "Total to pay": "total_pay",

    # TYPE OF WORK
    "type_of_work": "type_of_work",
    "Type of work": "type_of_work",
    "Type Of Work": "type_of_work",
    "Category": "type_of_work",

    # VENDOR
    "vendor_company": "vendor_company",
    "Vendor Company": "vendor_company",
    "Vendor company": "vendor_company",
    "Building & vendor company": "vendor_company",
}

df = df.rename(columns=column_map)

required_cols = ["date", "province", "employee", "hours", "total_pay", "type_of_work", "vendor_company"]
missing = [c for c in required_cols if c not in df.columns]

if missing:
    st.error(f"Missing required columns: {missing}")
    st.write("Detected columns:", list(df.columns))
    st.stop()

# ============================================================
# DATA CLEANING
# ============================================================
df["date"] = pd.to_datetime(df["date"], errors="coerce")
df["province"] = safe_text_series(df["province"]).str.upper()
df["employee"] = safe_text_series(df["employee"])
df["vendor_company"] = safe_text_series(df["vendor_company"])
df["type_of_work"] = safe_text_series(df["type_of_work"])
df["hours"] = pd.to_numeric(df["hours"], errors="coerce").fillna(0)
df["total_pay"] = pd.to_numeric(df["total_pay"], errors="coerce").fillna(0)

df = df.dropna(subset=["date"]).copy()
df = df[df["province"] == "QC"].copy()
df["work_class"] = df["type_of_work"].apply(normalize_work_type)

# ============================================================
# FILTERS
# ============================================================
st.sidebar.header("🔎 Committee Filters")

default_start = datetime(2026, 1, 4).date()
start_date = st.sidebar.date_input("First committee week start date", value=default_start)
num_weeks = st.sidebar.number_input("Number of weeks", min_value=1, max_value=12, value=4)

vendors = sorted([v for v in df["vendor_company"].dropna().unique().tolist() if v])
selected_vendors = st.sidebar.multiselect("Vendor Company", vendors, default=vendors)

employees = sorted([e for e in df["employee"].dropna().unique().tolist() if e])
selected_employees = st.sidebar.multiselect("Employee", employees, default=employees)

work_types = ["Regular", "Suppl.", "Congé", "Congé Travaillé", "Maladie", "Other"]
selected_work_types = st.sidebar.multiselect("Work Class", work_types, default=work_types)

# ============================================================
# ASSIGN WEEKS
# ============================================================
start_date_dt = pd.to_datetime(start_date)

df[["week_start", "week_end"]] = df["date"].apply(
    lambda x: pd.Series(assign_committee_week(x, start_date_dt, num_weeks))
)

df = df[df["week_start"].notna()].copy()
df = df[df["vendor_company"].isin(selected_vendors)]
df = df[df["employee"].isin(selected_employees)]
df = df[df["work_class"].isin(selected_work_types)]

df["week_label"] = df["week_end"].dt.strftime("%Y-%m-%d")

# ============================================================
# SUMMARY
# ============================================================
summary = (
    df.groupby(
        ["source_month_folder", "source_file", "vendor_company", "employee", "week_label", "work_class"],
        dropna=False,
    )
    .agg(
        total_hours=("hours", "sum"),
        total_pay=("total_pay", "sum"),
    )
    .reset_index()
    .sort_values(
        ["source_month_folder", "source_file", "vendor_company", "employee", "week_label", "work_class"]
    )
)

# ============================================================
# OUTPUT
# ============================================================
col1, col2, col3, col4 = st.columns(4)
col1.metric("Rows", len(df))
col2.metric("Total Hours", f"{df['hours'].sum():,.2f}")
col3.metric("Total Pay", f"${df['total_pay'].sum():,.2f}")
col4.metric("Files Used", f"{df['source_file'].nunique():,}")

st.subheader("Filtered source data")
st.dataframe(df, use_container_width=True)

st.subheader("Weekly summary")
st.dataframe(summary, use_container_width=True)

excel_file = to_excel_report(df, summary)

st.download_button(
    label="Download Excel Report",
    data=excel_file,
    file_name="comite_paritaire_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ============================================================
# DIAGNOSTICS
# ============================================================
with st.expander("Diagnostics", expanded=False):
    st.write("Redirect URI used:", REDIRECT_URI)
    st.write("Allowed domain:", ALLOWED_DOMAIN)
    st.write("Signed in as:", signed_in_email)
    st.write("Root folder resolved name:", meta.get("name"))
    st.write("Selected month folders:", selected_month_names)
    st.write("Selected excel display names:", selected_excel_display_names)
    st.write("Detected columns:", list(df.columns))
