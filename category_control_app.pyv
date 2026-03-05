import base64
from io import BytesIO
import pandas as pd
import streamlit as st
import requests
import msal

st.set_page_config(page_title="Invoice Category Control", layout="wide")

# -----------------------------
# SECRETS
# -----------------------------
TENANT_ID = st.secrets["TENANT_ID"]
CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
ONEDRIVE_SHARED_URL = st.secrets["ONEDRIVE_SHARED_URL"]

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

# -----------------------------
# FUNCTIONS
# -----------------------------
def encode_share_url(url):
    b64 = base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
    return "u!" + b64


@st.cache_data(ttl=300)
def download_excel():

    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
    )

    token = app.acquire_token_for_client(scopes=SCOPES)

    headers = {"Authorization": f"Bearer {token['access_token']}"}

    share_id = encode_share_url(ONEDRIVE_SHARED_URL)

    meta_url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem"

    meta = requests.get(meta_url, headers=headers).json()

    drive_id = meta["parentReference"]["driveId"]
    item_id = meta["id"]

    download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"

    file = requests.get(download_url, headers=headers)

    return BytesIO(file.content)


# -----------------------------
# LOAD DATA
# -----------------------------
excel = download_excel()

payments = pd.read_excel(
    excel,
    sheet_name="2025 Summary PAYMENTS"
)

excel.seek(0)

invoicing = pd.read_excel(
    excel,
    sheet_name="Invoicing"
)

# -----------------------------
# CLEAN DATA
# -----------------------------
payments = payments[[
    "Building Address",
    "Category",
    "Amount without taxes"
]]

payments["Amount without taxes"] = pd.to_numeric(
    payments["Amount without taxes"],
    errors="coerce"
)

payments = payments.dropna(subset=["Amount without taxes"])

# -----------------------------
# GROUP DATA
# -----------------------------
summary = payments.groupby(
    ["Building Address", "Category"]
)["Amount without taxes"].sum().reset_index()

# -----------------------------
# UI
# -----------------------------
st.title("📊 Invoice Category Control")

building = st.selectbox(
    "Building",
    ["All"] + sorted(summary["Building Address"].dropna().unique())
)

if building != "All":
    summary = summary[summary["Building Address"] == building]

st.subheader("Category Expenses")

st.dataframe(summary, use_container_width=True)

# -----------------------------
# BUDGET COMPARISON
# -----------------------------
st.subheader("Budget Comparison")

budget = invoicing[[
    "Building Address",
    "Total Labor Budget",
    "Total Supplies Budget",
    "Total Equipment Budget",
    "Total PW Budget"
]]

merged = summary.merge(
    budget,
    on="Building Address",
    how="left"
)

st.dataframe(merged, use_container_width=True)
