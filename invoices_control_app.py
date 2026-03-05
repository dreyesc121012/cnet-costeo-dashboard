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
REDIRECT_URI = str(st.secrets["REDIRECT_URI"]).strip().rstrip("/")

# Optional default Excel link
DEFAULT_SHARED_URL = str(st.secrets.get("ONEDRIVE_SHARED_URL", "")).strip()

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
            # st.query_params values may already be strings
            if isinstance(v, list):
                out[k] = v[0] if v else ""
            else:
                out[k] = str(v) if v is not None else ""
        return out
    except Exception:
        # compatibility
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
        raise RuntimeError(
            f"Error resolving shared link: {meta.status_code}\n{meta.text}\n\n"
            "TIP: Make sure the link was created from Share -> Copy link and is accessible in your tenant."
        )

    meta_json = meta.json()
    item_id = meta_json["id"]
    drive_id = meta_json["parentReference"]["driveId"]

    content_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    file_r = graph_get(content_url, access_token)
    if file_r.status_code != 200:
        raise RuntimeError(f"Error downloading file: {file_r.status_code}\n{file_r.text}")

    return file_r.content


# ============================================================
# Utilities
# ============================================================
def _norm(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def find_col(df: pd.DataFrame, name: str):
    t = _norm(name)
    # exact
    for c in df.columns:
        if _norm(c) == t:
            return c
    # contains
    for c in df.columns:
        if t in _norm(c):
            return c
    return None


def safe_num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0)


def get_excel_col_by_letter(df: pd.DataFrame, letter: str):
    """
    Excel column letter (A=1) -> 0-based pandas index.
    F=6 => index 5, G=7 => 6, H=8 => 7
    """
    letter = letter.strip().upper()
    if not re.fullmatch(r"[A-Z]+", letter):
        return None
    # convert base-26 letters to number
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    idx = n - 1
    if idx < 0 or idx >= df.shape[1]:
        return None
    return df.columns[idx]


# ============================================================
# Read Excel (bytes -> dataframes)
# ============================================================
@st.cache_data(ttl=300, show_spinner=False)
def load_sheets_from_bytes(excel_bytes: bytes) -> tuple[pd.DataFrame, pd.DataFrame]:
    xls = pd.ExcelFile(BytesIO(excel_bytes))
    payments = pd.read_excel(xls, sheet_name=SHEET_PAYMENTS)
    invoicing = pd.read_excel(xls, sheet_name=SHEET_INVOICING)

    payments.columns = [str(c).strip() for c in payments.columns]
    invoicing.columns = [str(c).strip() for c in invoicing.columns]
    return payments, invoicing


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

token_result = st.session_state.token_result
access_token = token_result.get("access_token", "")
if not access_token:
    st.error("No access token found. Please sign in again.")
    st.session_state.pop("token_result", None)
    st.stop()

st.success("✅ Connected to OneDrive/SharePoint")


# ============================================================
# SIDEBAR: Excel Source
# ============================================================
st.sidebar.header("📁 Excel Source")

shared_url = st.sidebar.text_input(
    "Paste OneDrive/SharePoint Excel share link",
    value=DEFAULT_SHARED_URL,
    help="From SharePoint/OneDrive: Share → Copy link (within your organization).",
).strip()

col_sb1, col_sb2 = st.sidebar.columns(2)
with col_sb1:
    load_btn = st.button("📥 Load / Update", use_container_width=True)
with col_sb2:
    if st.button("🔄 Reset cache", use_container_width=True):
        st.session_state.pop("excel_bytes", None)
        st.cache_data.clear()
        st.rerun()

if not shared_url:
    st.info("👈 Paste a OneDrive/SharePoint Excel share link in the sidebar to load data.")
    st.stop()

if load_btn or ("excel_bytes" not in st.session_state):
    try:
        st.info("📥 Downloading Excel from OneDrive/SharePoint…")
        st.session_state.excel_bytes = download_excel_bytes_from_shared_link(access_token, shared_url)
    except Exception as e:
        st.error("Could not download the Excel from OneDrive/SharePoint.")
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
# - match by Building Address
# - budget lines:
#   A) "Labor Budget" column (by name, if exists)
#   B) columns F, G, H (by letter) — used as additional budgets to compare (same building)
# ============================================================
inv_addr = find_col(invoicing_df, "Building Address")
if not inv_addr:
    # fallback: sometimes called Address / Building / Location
    inv_addr = find_col(invoicing_df, "Address") or find_col(invoicing_df, "Building") or find_col(invoicing_df, "Location")

labor_budget_col = find_col(invoicing_df, "Labor Budget")

col_F = get_excel_col_by_letter(invoicing_df, "F")
col_G = get_excel_col_by_letter(invoicing_df, "G")
col_H = get_excel_col_by_letter(invoicing_df, "H")

# Build a budgets dataframe in "long" format: Building Address, Category, Budget
budgets_long = []

if inv_addr:
    inv_base = invoicing_df.copy()
    inv_base[inv_addr] = inv_base[inv_addr].astype(str).str.strip()

    # A) Labor Budget -> category "Labor"
    if labor_budget_col:
        b = inv_base[[inv_addr, labor_budget_col]].copy()
        b.columns = ["Building Address", "Budget"]
        b["Category"] = "Labor"
        b["Budget"] = safe_num(b["Budget"])
        budgets_long.append(b[["Building Address", "Category", "Budget"]])

    # B) F/G/H -> categories "Budget F", "Budget G", "Budget H"
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
# COMPARE: Actual vs Budget
# Rules:
# - Compare Labor actuals vs Labor Budget (Category match "Labor")
# - For other categories, you can still view actuals; budgets F/G/H shown separately
# ============================================================
compare = actuals.merge(budgets, on=["Building Address", "Category"], how="left")
compare["Budget"] = compare["Budget"].fillna(0)
compare["Variance"] = compare["Budget"] - compare["Actual"]  # positive = under spend, negative = over
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

# Top categories by actual
top = view.groupby("Category", as_index=False).agg(Actual=("Actual", "sum"), Budget=("Budget", "sum"))
top = top.sort_values("Actual", ascending=False).head(20)

fig1 = px.bar(
    top,
    x="Category",
    y="Actual",
    title="Top Categories by Actual (No Taxes)",
)
st.plotly_chart(fig1, use_container_width=True)

# Actual vs Budget by category (only where budget exists)
top2 = top[top["Budget"] > 0].copy()
if not top2.empty:
    top2_m = top2.melt(id_vars=["Category"], value_vars=["Actual", "Budget"], var_name="Type", value_name="Amount")
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
# DIAGNOSTICS (optional)
# ============================================================
with st.expander("🛠 Diagnostics (detected columns)"):
    st.write("Payments columns:", list(payments_df.columns))
    st.write("Invoicing columns:", list(invoicing_df.columns))
    st.write("Detected in Payments:", {"Building Address": c_addr, "Category": c_cat, "Amount without taxes": c_amt})
    st.write("Detected in Invoicing:", {"Building Address": inv_addr, "Labor Budget": labor_budget_col, "F": col_F, "G": col_G, "H": col_H})
