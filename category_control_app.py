import streamlit as st
import pandas as pd
import requests
import msal
import io

st.set_page_config(page_title="Category Control Dashboard", layout="wide")

# =========================================================
# SECRETS
# =========================================================

TENANT_ID = st.secrets["TENANT_ID"]
CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]

ONEDRIVE_FOLDER_URL = st.secrets["ONEDRIVE_FOLDER_URL"]

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

# =========================================================
# AUTHENTICATION
# =========================================================

@st.cache_resource
def get_access_token():
    
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )

    token = app.acquire_token_for_client(scopes=SCOPES)

    if "access_token" in token:
        return token["access_token"]
    else:
        st.error("Error getting access token")
        st.stop()

# =========================================================
# GET FILE LIST FROM SHAREPOINT
# =========================================================

def get_files():

    token = get_access_token()

    headers = {
        "Authorization": f"Bearer {token}"
    }

    graph_url = f"https://graph.microsoft.com/v1.0/sites/root:/{ONEDRIVE_FOLDER_URL}:/children"

    r = requests.get(graph_url, headers=headers)

    if r.status_code != 200:
        st.error("Error accessing SharePoint folder")
        st.stop()

    data = r.json()

    files = []

    for item in data["value"]:
        if item["name"].endswith(".xlsx"):
            files.append({
                "name": item["name"],
                "download": item["@microsoft.graph.downloadUrl"]
            })

    return files


# =========================================================
# LOAD EXCEL
# =========================================================

def load_excel(download_url):

    r = requests.get(download_url)

    excel_file = io.BytesIO(r.content)

    df = pd.read_excel(
        excel_file,
        sheet_name="2025 Summary PAYMENTS"
    )

    return df


# =========================================================
# UI
# =========================================================

st.title("📊 Supplier Payments Category Control")

files = get_files()

file_names = [f["name"] for f in files]

selected = st.selectbox(
    "Select Excel file",
    file_names
)

selected_file = next(
    f for f in files if f["name"] == selected
)

df = load_excel(selected_file["download"])

# =========================================================
# CLEAN DATA
# =========================================================

df = df[[
    "Building Address",
    "Category",
    "Amount without taxes"
]]

df = df.dropna()

# =========================================================
# METRICS
# =========================================================

col1, col2 = st.columns(2)

col1.metric(
    "Total Expenses",
    f"${df['Amount without taxes'].sum():,.2f}"
)

col2.metric(
    "Transactions",
    len(df)
)

# =========================================================
# BUILDING SUMMARY
# =========================================================

st.subheader("Expenses by Building")

building = (
    df.groupby("Building Address")["Amount without taxes"]
    .sum()
    .sort_values(ascending=False)
)

st.bar_chart(building)

# =========================================================
# CATEGORY SUMMARY
# =========================================================

st.subheader("Expenses by Category")

category = (
    df.groupby("Category")["Amount without taxes"]
    .sum()
    .sort_values(ascending=False)
)

st.bar_chart(category)

# =========================================================
# TABLE
# =========================================================

st.subheader("Detailed Data")

st.dataframe(df, use_container_width=True)
