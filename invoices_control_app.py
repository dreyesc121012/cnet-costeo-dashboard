import base64
from io import BytesIO
import re
from typing import List, Dict, Optional, Tuple

import pandas as pd
import streamlit as st
import requests
import msal
import plotly.express as px

# ============================================================
# CONFIG (Secrets)
# ============================================================
CLIENT_ID = str(st.secrets["CLIENT_ID"]).strip()
CLIENT_SECRET = str(st.secrets["CLIENT_SECRET"]).strip()
TENANT_ID = str(st.secrets["TENANT_ID"]).strip()

# MUST match Azure App Registration Redirect URI EXACTLY
REDIRECT_URI = str(st.secrets["REDIRECT_URI"]).strip().rstrip("/")

# Optional: default SharePoint/OneDrive *FOLDER* (recommended) OR *FILE* share link
DEFAULT_SHARED_URL = str(st.secrets.get("ONEDRIVE_SHARED_URL", "")).strip()

# Optional login experience hints
DOMAIN_HINT = str(st.secrets.get("DOMAIN_HINT", "")).strip()
LOGIN_HINT = str(st.secrets.get("LOGIN_HINT", "")).strip()

# Required: allowed corporate domain
ALLOWED_DOMAIN = str(st.secrets.get("ALLOWED_DOMAIN", "")).strip().lower()

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.Read.All"]

SHEET_PAYMENTS = "2025 Summary PAYMENTS"
SHEET_INVOICING = "Invoicing"

st.set_page_config(page_title="Invoices Control", layout="wide")
st.title("📑 Invoice Category Control Dashboard")

# ============================================================
# HELPERS (URL params)
# ============================================================
def _get_query_params() -> dict:
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
    """
    Returns driveItem metadata for a shared link (file OR folder).
    """
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

def list_children_all(access_token: str, drive_id: str, folder_item_id: str) -> List[Dict]:
    """
    Lists ALL children of a folder (handles paging).
    """
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

# ============================================================
# UTILITIES
# ============================================================
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

def get_excel_col_by_letter(df: pd.DataFrame, letter: str):
    letter = letter.strip().upper()
    if not re.fullmatch(r"[A-Z]+", letter):
        return None
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    idx = n - 1
    if idx < 0 or idx >= df.shape[1]:
        return None
    return df.columns[idx]

# ============================================================
# READ EXCEL
# ============================================================
@st.cache_data(ttl=300, show_spinner=False)
def load_sheets_from_bytes(excel_bytes: bytes) -> Tuple[pd.DataFrame, pd.DataFrame]:
    xls = pd.ExcelFile(BytesIO(excel_bytes))
    payments = pd.read_excel(xls, sheet_name=SHEET_PAYMENTS)
    invoicing = pd.read_excel(xls, sheet_name=SHEET_INVOICING)
    payments.columns = [str(c).strip() for c in payments.columns]
    invoicing.columns = [str(c).strip() for c in invoicing.columns]
    return payments, invoicing

@st.cache_data(ttl=300, show_spinner=False)
def cached_folder_excel_list(
    access_token: str,
    drive_id: str,
    folder_id: str,
    shared_url: str,
    marker: str,
):
    _ = shared_url, marker
    children = list_children_all(access_token, drive_id, folder_id)
    excels = [c for c in children if c.get("id") and is_excel_name(c.get("name", ""))]
    excels.sort(key=lambda x: (x.get("name") or "").lower())
    return excels

# ============================================================
# AUTHENTICATION FLOW
# ============================================================
if not ALLOWED_DOMAIN:
    st.error("Missing ALLOWED_DOMAIN in Streamlit secrets.")
    st.stop()

app = get_msal_app()
qp = _get_query_params()

if "token_result" not in st.session_state:
    code = qp.get("code")

    if code:
        result = app.acquire_token_by_authorization_code(
            code=code,
            scopes=SCOPES,
            redirect_uri=REDIRECT_URI,
        )

        if "access_token" in result:
            st.session_state.token_result = result
            try:
                st.query_params.clear()
            except Exception:
                _clear_query_params()
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

    st.markdown(
        f"""
        <a href="{auth_url}" target="_self">
            <button style="
                background-color:#ffffff;
                border:1px solid #d0d0d0;
                border-radius:8px;
                padding:10px 16px;
                font-size:16px;
                cursor:pointer;">
                🔐 Sign in with Microsoft (Company)
            </button>
        </a>
        """,
        unsafe_allow_html=True,
    )

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

st.success(f"✅ Signed in as {signed_in_email}")

if st.button("🚪 Sign out"):
    st.session_state.pop("token_result", None)
    st.session_state.pop("excel_bytes", None)
    st.session_state.pop("selected_item_id", None)
    st.cache_data.clear()
    try:
        st.query_params.clear()
    except Exception:
        _clear_query_params()
    st.rerun()

# ============================================================
# SIDEBAR: Folder/File Source
# ============================================================
st.sidebar.header("📁 SharePoint Source")
st.sidebar.success(f"Logged in as {signed_in_email}")

shared_url = st.sidebar.text_input(
    "Paste SharePoint/OneDrive share link (FOLDER recommended)",
    value=DEFAULT_SHARED_URL,
    help="SharePoint/OneDrive: Share → Copy link (within your organization). Use a FOLDER link to pick any Excel inside.",
).strip()

col_sb1, col_sb2 = st.sidebar.columns(2)
with col_sb1:
    refresh_btn = st.sidebar.button("🔄 Refresh list", use_container_width=True)
with col_sb2:
    if st.sidebar.button("🧹 Clear cache", use_container_width=True):
        st.session_state.pop("excel_bytes", None)
        st.session_state.pop("selected_item_id", None)
        st.cache_data.clear()
        st.rerun()

if not shared_url:
    st.info("👈 Paste a SharePoint/OneDrive share link in the sidebar.")
    st.stop()

# ============================================================
# Resolve link (file or folder)
# ============================================================
try:
    meta = resolve_shared_link(access_token, shared_url)
except Exception as e:
    st.error("Could not resolve the SharePoint link.")
    st.code(str(e))
    st.stop()

drive_id = meta["parentReference"]["driveId"]
root_item_id = meta["id"]
is_folder = "folder" in meta

selected_item_id: Optional[str] = None
selected_name: Optional[str] = None

if is_folder:
    st.sidebar.subheader("📄 Excel files in folder")

    marker = f"{len(access_token)}-{TENANT_ID[-6:]}"
    excels = cached_folder_excel_list(access_token, drive_id, root_item_id, shared_url, marker)

    if not excels:
        st.warning("No Excel files found in this folder.")
        st.stop()

    names = [f["name"] for f in excels]
    default_ix = 0
    prev = st.session_state.get("selected_item_id")
    if prev:
        for i, f in enumerate(excels):
            if f["id"] == prev:
                default_ix = i
                break

    selected_name = st.sidebar.selectbox("Choose an Excel file", names, index=default_ix)
    chosen = next(f for f in excels if f["name"] == selected_name)
    selected_item_id = chosen["id"]

else:
    selected_item_id = root_item_id
    selected_name = meta.get("name", "Selected file")
    st.sidebar.caption(f"Using file: {selected_name}")

# ============================================================
# DOWNLOAD FILE
# ============================================================
needs_download = (
    ("excel_bytes" not in st.session_state)
    or (st.session_state.get("selected_item_id") != selected_item_id)
    or refresh_btn
)

if needs_download:
    try:
        st.info("📥 Downloading Excel from SharePoint/OneDrive...")
        st.session_state.excel_bytes = download_item_bytes(access_token, drive_id, selected_item_id)
        st.session_state.selected_item_id = selected_item_id
    except Exception as e:
        st.error("Could not download the selected Excel file.")
        st.code(str(e))
        st.stop()

excel_bytes = st.session_state.excel_bytes

# ============================================================
# LOAD DATA
# ============================================================
try:
    payments_df, invoicing_df = load_sheets_from_bytes(excel_bytes)
except Exception as e:
    st.error("Could not read the required sheets from the Excel file.")
    st.code(str(e))
    st.stop()

# ============================================================
# BUILD ACTUALS (PAYMENTS)
# ============================================================
c_addr = find_col(payments_df, "Building Address")
c_cat = find_col(payments_df, "Category")
c_amt = find_col(payments_df, "Amount without taxes")

if not (c_addr and c_cat and c_amt):
    st.error("Missing required columns in '2025 Summary PAYMENTS'.")
    st.write("Detected columns:", list(payments_df.columns))
    st.stop()

tmp = payments_df[[c_addr, c_cat, c_amt]].copy()
tmp.columns = ["Building Address", "Category", "Actual"]
tmp["Building Address"] = tmp["Building Address"].astype(str).str.strip()
tmp["Category"] = tmp["Category"].astype(str).str.strip()
tmp["Actual"] = safe_num(tmp["Actual"])

actuals = (
    tmp.groupby(["Building Address", "Category"], as_index=False)["Actual"]
    .sum()
    .sort_values("Actual", ascending=False)
)

# ============================================================
# BUILD BUDGETS (INVOICING)
# ============================================================
inv_addr = find_col(invoicing_df, "Building Address")
if not inv_addr:
    inv_addr = (
        find_col(invoicing_df, "Address")
        or find_col(invoicing_df, "Building")
        or find_col(invoicing_df, "Location")
    )

labor_budget_col = find_col(invoicing_df, "Labor Budget")
col_F = get_excel_col_by_letter(invoicing_df, "F")
col_G = get_excel_col_by_letter(invoicing_df, "G")
col_H = get_excel_col_by_letter(invoicing_df, "H")

budgets_long = []

if inv_addr:
    inv_base = invoicing_df.copy()
    inv_base[inv_addr] = inv_base[inv_addr].astype(str).str.strip()

    if labor_budget_col:
        b = inv_base[[inv_addr, labor_budget_col]].copy()
        b.columns = ["Building Address", "Budget"]
        b["Category"] = "Labor"
        b["Budget"] = safe_num(b["Budget"])
        budgets_long.append(b[["Building Address", "Category", "Budget"]])

    for letter, colname in [("F", col_F), ("G", col_G), ("H", col_H)]:
        if colname:
            b = inv_base[[inv_addr, colname]].copy()
            b.columns = ["Building Address", "Budget"]
            b["Category"] = f"Budget {letter}"
            b["Budget"] = safe_num(b["Budget"])
            budgets_long.append(b[["Building Address", "Category", "Budget"]])

if budgets_long:
    budgets = pd.concat(budgets_long, ignore_index=True)
    budgets = budgets.groupby(["Building Address", "Category"], as_index=False)["Budget"].sum()
else:
    budgets = pd.DataFrame(columns=["Building Address", "Category", "Budget"])

# ============================================================
# COMPARE
# ============================================================
compare = actuals.merge(budgets, on=["Building Address", "Category"], how="left")
compare["Budget"] = compare["Budget"].fillna(0)
compare["Variance"] = compare["Budget"] - compare["Actual"]
compare["% Used"] = compare.apply(lambda r: (r["Actual"] / r["Budget"]) if r["Budget"] else 0.0, axis=1)

# ============================================================
# FILTERS
# ============================================================
st.sidebar.header("🔎 Filters")

all_buildings = sorted(compare["Building Address"].dropna().unique().tolist())
sel_buildings = st.sidebar.multiselect("Building Address", all_buildings, default=[])

all_categories = sorted(compare["Category"].dropna().unique().tolist())
sel_categories = st.sidebar.multiselect("Category", all_categories, default=[])

view = compare.copy()
if sel_buildings:
    view = view[view["Building Address"].isin(sel_buildings)]
if sel_categories:
    view = view[view["Category"].isin(sel_categories)]

# ============================================================
# KPIs
# ============================================================
st.subheader("📌 KPIs")

total_actual = float(view["Actual"].sum())
total_budget = float(view["Budget"].sum())
total_var = total_budget - total_actual

k1, k2, k3 = st.columns(3)
k1.metric("Total Actual (no taxes)", f"${total_actual:,.2f}")
k2.metric("Total Budget (matched categories)", f"${total_budget:,.2f}")
status = "🟢 Under budget" if total_var > 0 else ("🔴 Over budget" if total_var < 0 else "⚪ On budget")
k3.metric("Variance (Budget - Actual)", f"${total_var:,.2f}", status)

# ============================================================
# TABLE
# ============================================================
st.subheader("📋 Actual vs Budget (by Building & Category)")
st.dataframe(
    view.sort_values(["Building Address", "Category"]),
    use_container_width=True
)

# ============================================================
# CHARTS
# ============================================================
st.subheader("📊 Charts")

top = view.groupby("Category", as_index=False).agg(Actual=("Actual", "sum"), Budget=("Budget", "sum"))
top = top.sort_values("Actual", ascending=False).head(20)

fig1 = px.bar(
    top,
    x="Category",
    y="Actual",
    title="Top Categories by Actual (No Taxes)",
)
st.plotly_chart(fig1, use_container_width=True)

top2 = top[top["Budget"] > 0].copy()
if not top2.empty:
    top2_m = top2.melt(
        id_vars=["Category"],
        value_vars=["Actual", "Budget"],
        var_name="Type",
        value_name="Amount",
    )
    fig2 = px.bar(
        top2_m,
        x="Category",
        y="Amount",
        color="Type",
        barmode="group",
        title="Actual vs Budget by Category (where budget exists)",
    )
    st.plotly_chart(fig2, use_container_width=True)
else:
    st.info("No budget columns detected/matched yet (Labor Budget or F/G/H). Budgets are currently 0.")

# ============================================================
# DIAGNOSTICS
# ============================================================
with st.expander("🛠 Diagnostics"):
    st.write("Redirect URI used:", REDIRECT_URI)
    st.write("Allowed domain:", ALLOWED_DOMAIN)
    st.write("Signed in as:", signed_in_email)
    st.write("Selected Excel:", selected_name)
    st.write("Link type:", "FOLDER" if is_folder else "FILE")
    st.write("Payments columns:", list(payments_df.columns))
    st.write("Invoicing columns:", list(invoicing_df.columns))
    st.write("Detected in Payments:", {"Building Address": c_addr, "Category": c_cat, "Amount without taxes": c_amt})
    st.write("Detected in Invoicing:", {"Building Address": inv_addr, "Labor Budget": labor_budget_col, "F": col_F, "G": col_G, "H": col_H})
