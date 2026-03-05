import base64
from io import BytesIO
import re

import pandas as pd
import streamlit as st
import requests
import msal
import plotly.express as px

# ============================================================
# CONFIG (Secrets)
# ============================================================
CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
TENANT_ID = st.secrets["TENANT_ID"]
REDIRECT_URI = st.secrets["REDIRECT_URI"].strip().rstrip("/")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.Read.All"]

SHEET_PAYMENTS = "2025 Summary PAYMENTS"
SHEET_INVOICING = "Invoicing"

st.set_page_config(page_title="Invoices Control", layout="wide")
st.title("📑 Invoice Category Control Dashboard")

# ============================================================
# HELPERS (URL params) - needed for MSAL redirect (?code=...)
# ============================================================
def _get_query_params() -> dict:
    try:
        qp = st.query_params
        out = {}
        for k in qp.keys():
            v = qp.get(k)
            out[k] = v[0] if isinstance(v, list) else str(v)
        return out
    except Exception:
        try:
            qp = st.experimental_get_query_params()
            return {k: (v[0] if isinstance(v, list) and v else str(v)) for k, v in qp.items()}
        except Exception:
            return {}

def _clear_query_params():
    try:
        st.query_params.clear()
    except Exception:
        try:
            st.experimental_set_query_params()
        except Exception:
            pass

# ============================================================
# MSAL APP
# ============================================================
def get_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        token_cache=None,
    )

# ============================================================
# OneDrive shared-link -> bytes (Graph)
# ============================================================
def make_share_id(shared_url: str) -> str:
    b = base64.b64encode(shared_url.encode("utf-8")).decode("utf-8")
    b = b.rstrip("=").replace("/", "_").replace("+", "-")
    return "u!" + b

def graph_get(url: str, access_token: str) -> requests.Response:
    return requests.get(url, headers={"Authorization": f"Bearer {access_token}"}, timeout=60)

def download_excel_bytes_from_shared_link(access_token: str, shared_url: str) -> bytes:
    share_id = make_share_id(shared_url)

    meta_url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem"
    meta = graph_get(meta_url, access_token)
    if meta.status_code != 200:
        raise RuntimeError(f"Error resolving shared link: {meta.status_code}\n{meta.text}")

    meta_json = meta.json()
    item_id = meta_json["id"]
    drive_id = meta_json["parentReference"]["driveId"]

    content_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    file_r = graph_get(content_url, access_token)
    if file_r.status_code != 200:
        raise RuntimeError(f"Error downloading file: {file_r.status_code}\n{file_r.text}")

    return file_r.content

# ============================================================
# Read Excel
# ============================================================
@st.cache_data(ttl=300, show_spinner=False)
def load_sheets_from_bytes(excel_bytes: bytes) -> tuple[pd.DataFrame, pd.DataFrame]:
    xls = pd.ExcelFile(BytesIO(excel_bytes))
    payments = pd.read_excel(xls, sheet_name=SHEET_PAYMENTS)
    invoicing = pd.read_excel(xls, sheet_name=SHEET_INVOICING)
    payments.columns = [str(c).strip() for c in payments.columns]
    invoicing.columns = [str(c).strip() for c in invoicing.columns]
    return payments, invoicing

def _norm(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def find_col(df: pd.DataFrame, name: str):
    t = _norm(name)
    for c in df.columns:
        if _norm(c) == t:
            return c
    for c in df.columns:
        if t in _norm(c):
            return c
    return None

def safe_num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0)

# ============================================================
# LOGIN FLOW
# ============================================================
app = get_msal_app()
qp = _get_query_params()

if "token_result" not in st.session_state:
    # Coming back from Microsoft with ?code=...
    if qp.get("code"):
        result = app.acquire_token_by_authorization_code(
            code=qp["code"],
            scopes=SCOPES,
            redirect_uri=REDIRECT_URI,
        )
        if "access_token" in result:
            st.session_state.token_result = result
            _clear_query_params()
            st.rerun()
        else:
            st.error("Could not obtain access token.")
            st.code(result)
            st.stop()

    # Not logged in yet
    st.warning("You are not signed in to OneDrive/SharePoint.")
    auth_url = app.get_authorization_request_url(scopes=SCOPES, redirect_uri=REDIRECT_URI)
    st.link_button("🔐 Sign in to OneDrive", auth_url)
    st.stop()

access_token = st.session_state.token_result["access_token"]
st.success("✅ Connected to OneDrive/SharePoint")

# ============================================================
# UI: Choose Excel by shared URL
# ============================================================
st.subheader("Excel Source")
shared_url = st.text_input("Paste OneDrive/SharePoint shared link for the Excel file")

if st.button("📥 Load Excel"):
    if not shared_url.strip():
        st.error("Paste a valid shared link first.")
        st.stop()

    st.info("Downloading Excel…")
    excel_bytes = download_excel_bytes_from_shared_link(access_token, shared_url.strip())

    payments_df, invoicing_df = load_sheets_from_bytes(excel_bytes)

    # Build actuals
    c_addr = find_col(payments_df, "Building Address")
    c_cat = find_col(payments_df, "Category")
    c_amt = find_col(payments_df, "Amount without taxes")

    if not (c_addr and c_cat and c_amt):
        st.error("Missing required columns in '2025 Summary PAYMENTS'.")
        st.write("Detected columns:", list(payments_df.columns))
        st.stop()

    tmp = payments_df[[c_addr, c_cat, c_amt]].copy()
    tmp.columns = ["Building Address", "Category", "Actual"]
    tmp["Actual"] = safe_num(tmp["Actual"])

    actuals = tmp.groupby(["Building Address", "Category"], as_index=False)["Actual"].sum()

    st.subheader("Actual Amount by Category (No Taxes)")
    st.dataframe(actuals, use_container_width=True)

    fig = px.bar(
        actuals.sort_values("Actual", ascending=False),
        x="Category",
        y="Actual",
        title="Actual Amount by Category"
    )
    st.plotly_chart(fig, use_container_width=True)

    st.info("Next step: map budgets from 'Invoicing' (Labor Budget + F/G/H) and compare vs actuals.")
