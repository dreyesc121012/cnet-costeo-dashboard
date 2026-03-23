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
REDIRECT_URI = str(st.secrets["REDIRECT_URI"]).strip().rstrip("/")
DEFAULT_SHARED_URL = str(st.secrets.get("ONEDRIVE_SHARED_URL", "")).strip()
DOMAIN_HINT = str(st.secrets.get("DOMAIN_HINT", "")).strip()
LOGIN_HINT = str(st.secrets.get("LOGIN_HINT", "")).strip()
ALLOWED_DOMAIN = str(st.secrets.get("ALLOWED_DOMAIN", "")).strip().lower()

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.Read.All"]

SHEET_PAYMENTS = "Summary PAYMENTS"
SHEET_INVOICING = "Invoicing"

st.set_page_config(page_title="Invoices Control", layout="wide")
st.title("📑 Invoice Category Control Dashboard")

# ============================================================
# HELPERS (URL params)
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
# FORMATTERS
# ============================================================
def fmt_currency(x) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"

def fmt_percent_ratio(x) -> str:
    try:
        return f"{float(x) * 100:,.2f}%"
    except Exception:
        return "0.00%"

def status_semaphore(pct_used: float, budget: float, real: float) -> str:
    if budget <= 0:
        if real > 0:
            return "⚪ No Budget"
        return "⚪ N/A"
    if real <= 0:
        return "🔴 No Real"
    if pct_used < 0.80:
        return "🟡 Below 80%"
    if pct_used <= 1.00:
        return "🟢 On Track"
    return "🔴 Over Budget"

STATUS_COLOR_MAP = {
    "🔴 No Real": "#E53935",
    "🟡 Below 80%": "#FBC02D",
    "🟢 On Track": "#43A047",
    "🔴 Over Budget": "#B71C1C",
    "⚪ No Budget": "#9E9E9E",
    "⚪ N/A": "#BDBDBD",
}

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

def clean_text_value(v) -> str:
    if pd.isna(v):
        return ""
    s = str(v).replace("\u00A0", " ").strip()
    if s.lower() in {"nan", "none", "null"}:
        return ""
    return s

def clean_series_text(s: pd.Series, blank_label: str = "Unspecified") -> pd.Series:
    out = s.astype(str).str.replace("\u00A0", " ", regex=False).str.strip()
    out = out.replace(
        {
            "nan": "",
            "None": "",
            "none": "",
            "NULL": "",
            "null": "",
            "<NA>": "",
        }
    )
    out = out.fillna("")
    out = out.mask(out.str.strip() == "", blank_label)
    return out

def find_col(df: pd.DataFrame, name: str):
    t = _norm(name)
    for c in df.columns:
        if _norm(c) == t:
            return c
    for c in df.columns:
        if t in _norm(c):
            return c
    return None

def find_first_existing_col(df: pd.DataFrame, candidates: List[str]):
    for cand in candidates:
        col = find_col(df, cand)
        if col:
            return col
    return None

def safe_num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0)

def _score_header_row(row_values: List[str], expected_keywords: List[str]) -> int:
    normalized = [_norm(v) for v in row_values]
    score = 0
    for kw in expected_keywords:
        nkw = _norm(kw)
        if any(nkw == cell or nkw in cell for cell in normalized):
            score += 1
    return score

def detect_header_row(raw_df: pd.DataFrame, expected_keywords: List[str], max_rows: int = 15) -> int:
    best_row = 0
    best_score = -1
    rows_to_check = min(max_rows, len(raw_df))
    for i in range(rows_to_check):
        row_values = raw_df.iloc[i].fillna("").astype(str).tolist()
        score = _score_header_row(row_values, expected_keywords)
        if score > best_score:
            best_score = score
            best_row = i
    return best_row

def read_sheet_with_detected_header(
    xls: pd.ExcelFile,
    sheet_name: str,
    expected_keywords: List[str],
    preview_rows: int = 15,
) -> Tuple[pd.DataFrame, int]:
    raw = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=preview_rows)
    header_row = detect_header_row(raw, expected_keywords, max_rows=preview_rows)
    df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    return df, header_row

def get_column_by_position(df: pd.DataFrame, zero_based_idx: int):
    if 0 <= zero_based_idx < len(df.columns):
        return df.columns[zero_based_idx]
    return None

def find_col_or_position(df: pd.DataFrame, candidates: List[str], zero_based_idx: int = None):
    col = find_first_existing_col(df, candidates)
    if col:
        return col
    if zero_based_idx is not None:
        return get_column_by_position(df, zero_based_idx)
    return None

def normalize_category_series(s: pd.Series) -> pd.Series:
    temp = s.astype(str).str.strip()
    mapping = {
        "labour": "Labor",
        "labor": "Labor",
        "supplies": "Supplies",
        "equipment": "Equipment",
        "pw": "PW",
        "power washing": "PW",
        "powerwashing": "PW",
    }
    return temp.apply(lambda x: mapping.get(x.strip().lower(), x.strip()))

# ============================================================
# READ EXCEL
# ============================================================
@st.cache_data(ttl=300, show_spinner=False)
def load_sheets_from_bytes(excel_bytes: bytes) -> Tuple[pd.DataFrame, pd.DataFrame, int, int]:
    xls = pd.ExcelFile(BytesIO(excel_bytes))

    payments_expected = [
        "Building Address",
        "Category",
        "Amount without taxes",
        "Amount",
        "Address",
        "Supplier",
        "Type of work",
        "Diverse expense",
        "Province",
    ]
    invoicing_expected = [
        "Building Address",
        "Company",
        "Company Name",
        "Client",
        "Customer",
        "Labor Budget",
        "Supplies Budget",
        "Equipment Budget",
        "PW Budget",
        "Address",
        "Building",
        "Location",
    ]

    payments, payments_header_row = read_sheet_with_detected_header(
        xls, SHEET_PAYMENTS, payments_expected
    )
    invoicing, invoicing_header_row = read_sheet_with_detected_header(
        xls, SHEET_INVOICING, invoicing_expected
    )

    return payments, invoicing, payments_header_row, invoicing_header_row

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

st.sidebar.header("📁 SharePoint Source")
st.sidebar.success(f"Logged in as {signed_in_email}")
st.success(f"✅ Signed in as {signed_in_email}")

if st.button("🚪 Sign out"):
    st.session_state.pop("token_result", None)
    st.session_state.pop("excel_file_map", None)
    st.session_state.pop("selected_item_ids", None)
    st.cache_data.clear()
    clear_query_params_compat()
    st.rerun()

# ============================================================
# SIDEBAR: Folder/File Source
# ============================================================
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
        st.session_state.pop("excel_file_map", None)
        st.session_state.pop("selected_item_ids", None)
        st.cache_data.clear()
        st.rerun()

if not shared_url:
    st.info("👈 Paste a SharePoint/OneDrive share link in the sidebar.")
    st.stop()

# ============================================================
# RESOLVE LINK
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

selected_item_ids: List[str] = []
selected_names: List[str] = []

if is_folder:
    st.sidebar.subheader("📄 Excel files in folder")

    marker = f"{len(access_token)}-{TENANT_ID[-6:]}"
    excels = cached_folder_excel_list(access_token, drive_id, root_item_id, shared_url, marker)

    if not excels:
        st.warning("No Excel files found in this folder.")
        st.stop()

    names = [f["name"] for f in excels]
    month_options = ["ALL MONTHS"] + names

    prev_selected_ids = st.session_state.get("selected_item_ids", [])
    default_selected = [f["name"] for f in excels if f["id"] in prev_selected_ids] if prev_selected_ids else []

    selected_names = st.sidebar.multiselect(
        "Choose one or more Excel files",
        month_options,
        default=default_selected,
    )

    if "ALL MONTHS" in selected_names:
        selected_names = names

    if not selected_names:
        st.info("Please select at least one Excel file.")
        st.stop()

    selected_item_ids = [f["id"] for f in excels if f["name"] in selected_names]

else:
    selected_item_ids = [root_item_id]
    selected_names = [meta.get("name", "Selected file")]
    st.sidebar.caption(f"Using file: {selected_names[0]}")

# ============================================================
# DOWNLOAD FILE(S)
# ============================================================
needs_download = (
    ("excel_file_map" not in st.session_state)
    or (st.session_state.get("selected_item_ids") != selected_item_ids)
    or refresh_btn
)

if needs_download:
    try:
        st.info("📥 Downloading selected Excel file(s) from SharePoint/OneDrive...")
        excel_file_map = {}

        for item_id, file_name in zip(selected_item_ids, selected_names):
            excel_file_map[file_name] = download_item_bytes(access_token, drive_id, item_id)

        st.session_state.excel_file_map = excel_file_map
        st.session_state.selected_item_ids = selected_item_ids

    except Exception as e:
        st.error("Could not download the selected Excel file(s).")
        st.code(str(e))
        st.stop()

excel_file_map = st.session_state.excel_file_map

# ============================================================
# LOAD DATA FROM MULTIPLE FILES
# ============================================================
all_payments = []
all_invoicing = []
payments_header_rows = {}
invoicing_header_rows = {}

try:
    for file_name, excel_bytes in excel_file_map.items():
        p_df, i_df, p_header, i_header = load_sheets_from_bytes(excel_bytes)

        p_df = p_df.copy()
        i_df = i_df.copy()

        p_df["Source File"] = file_name
        i_df["Source File"] = file_name

        all_payments.append(p_df)
        all_invoicing.append(i_df)

        payments_header_rows[file_name] = p_header
        invoicing_header_rows[file_name] = i_header

    payments_df = pd.concat(all_payments, ignore_index=True) if all_payments else pd.DataFrame()
    invoicing_df = pd.concat(all_invoicing, ignore_index=True) if all_invoicing else pd.DataFrame()

except Exception as e:
    st.error("Could not read the required sheets from the selected Excel files.")
    st.code(str(e))
    st.stop()

# ============================================================
# REQUIRED COLUMNS
# ============================================================
pay_addr = find_col(payments_df, "Building Address")
pay_cat = find_col(payments_df, "Category")
pay_amt = find_col(payments_df, "Amount without taxes")
if not pay_amt:
    pay_amt = find_col(payments_df, "Amount")

# Requested breakdown columns from Summary PAYMENTS
pay_supplier = find_col_or_position(payments_df, ["Supplier"], zero_based_idx=1)           # Column B
pay_type_of_work = find_col_or_position(payments_df, ["Type of work", "Type Of Work"], 5)  # Column F
pay_diverse_expense = find_col_or_position(
    payments_df,
    ["Diverse expense", "Diverse Expense"],
    6,  # Column G
)
pay_province = find_col_or_position(payments_df, ["Province"], 10)                          # Column K
pay_source_file = find_col(payments_df, "Source File")

inv_addr = find_col(invoicing_df, "Building Address")
if not inv_addr:
    inv_addr = (
        find_col(invoicing_df, "Address")
        or find_col(invoicing_df, "Building")
        or find_col(invoicing_df, "Location")
    )

inv_company = find_first_existing_col(
    invoicing_df,
    ["Company", "Company Name", "Client", "Customer", "Customer Name", "Account"]
)

labor_budget_col = find_col(invoicing_df, "Labor Budget")
supplies_budget_col = find_col(invoicing_df, "Supplies Budget")
equipment_budget_col = find_col(invoicing_df, "Equipment Budget")
pw_budget_col = find_col(invoicing_df, "PW Budget")
inv_source_file = find_col(invoicing_df, "Source File")

if not (pay_addr and pay_cat and pay_amt):
    st.error("Missing required columns in 'Summary PAYMENTS'.")
    st.write("Detected columns:", list(payments_df.columns))
    st.stop()

if not inv_addr:
    st.error("Missing 'Building Address' (or equivalent) in 'Invoicing'.")
    st.write("Detected columns:", list(invoicing_df.columns))
    st.stop()

# ============================================================
# BASE: INVOICING FIRST
# ============================================================
inv_base = invoicing_df.copy()
inv_base[inv_addr] = inv_base[inv_addr].astype(str).str.strip()

if inv_company:
    inv_base[inv_company] = clean_series_text(inv_base[inv_company], blank_label="Unassigned")
else:
    inv_base["__Company__"] = "Unassigned"
    inv_company = "__Company__"

if inv_source_file:
    inv_base[inv_source_file] = clean_series_text(inv_base[inv_source_file], blank_label="Unknown File")
else:
    inv_base["Source File"] = "Unknown File"
    inv_source_file = "Source File"

budget_rows = []

category_budget_map = [
    ("Labor", labor_budget_col),
    ("Supplies", supplies_budget_col),
    ("Equipment", equipment_budget_col),
    ("PW", pw_budget_col),
]

for category_name, budget_col in category_budget_map:
    if budget_col:
        b = inv_base[[inv_addr, inv_company, inv_source_file, budget_col]].copy()
        b.columns = ["Building Address", "Company", "Source File", "Budget"]
        b["Category"] = category_name
        b["Budget"] = safe_num(b["Budget"])
        budget_rows.append(b[["Source File", "Building Address", "Company", "Category", "Budget"]])

if budget_rows:
    budgets = pd.concat(budget_rows, ignore_index=True)
    budgets = budgets.groupby(
        ["Source File", "Building Address", "Company", "Category"],
        as_index=False
    )["Budget"].sum()
else:
    budgets = pd.DataFrame(columns=["Source File", "Building Address", "Company", "Category", "Budget"])

# ============================================================
# PAYMENTS BASE WITH EXTRA BREAKDOWN COLUMNS
# ============================================================
pay_cols_to_take = [pay_addr, pay_cat, pay_amt]
for extra_col in [pay_supplier, pay_type_of_work, pay_diverse_expense, pay_province, pay_source_file]:
    if extra_col and extra_col not in pay_cols_to_take:
        pay_cols_to_take.append(extra_col)

pay_base = payments_df[pay_cols_to_take].copy()

rename_map = {
    pay_addr: "Building Address",
    pay_cat: "Category",
    pay_amt: "Real",
}
if pay_supplier:
    rename_map[pay_supplier] = "Supplier"
if pay_type_of_work:
    rename_map[pay_type_of_work] = "Type of work"
if pay_diverse_expense:
    rename_map[pay_diverse_expense] = "Diverse expense"
if pay_province:
    rename_map[pay_province] = "Province"
if pay_source_file:
    rename_map[pay_source_file] = "Source File"

pay_base = pay_base.rename(columns=rename_map)

for col in ["Supplier", "Type of work", "Diverse expense", "Province", "Source File"]:
    if col not in pay_base.columns:
        pay_base[col] = "Unspecified"
    pay_base[col] = clean_series_text(pay_base[col], blank_label="Unspecified")

pay_base["Building Address"] = pay_base["Building Address"].astype(str).str.strip()
pay_base["Category"] = pay_base["Category"].astype(str).str.strip()
pay_base["Real"] = safe_num(pay_base["Real"])
pay_base["Category"] = normalize_category_series(pay_base["Category"])

pay_base = pay_base[
    pay_base["Category"].notna()
    & (pay_base["Category"].astype(str).str.strip() != "")
    & (~pay_base["Category"].astype(str).str.strip().str.lower().isin(["nan", "none", "null"]))
]

# ============================================================
# FILTERS
# ============================================================
st.sidebar.header("🔎 Filters")

all_source_files = sorted(pay_base["Source File"].dropna().unique().tolist())
sel_source_files = st.sidebar.multiselect("Excel File(s)", all_source_files, default=[])

all_suppliers = sorted(pay_base["Supplier"].dropna().unique().tolist())
sel_suppliers = st.sidebar.multiselect("Supplier", all_suppliers, default=[])

all_type_of_work = sorted(pay_base["Type of work"].dropna().unique().tolist())
sel_type_of_work = st.sidebar.multiselect("Type of work", all_type_of_work, default=[])

all_diverse_expense = sorted(pay_base["Diverse expense"].dropna().unique().tolist())
sel_diverse_expense = st.sidebar.multiselect("Diverse expense", all_diverse_expense, default=[])

all_provinces = sorted(pay_base["Province"].dropna().unique().tolist())
sel_provinces = st.sidebar.multiselect("Province", all_provinces, default=[])

# Apply payment-level filters first
filtered_pay_base = pay_base.copy()

if sel_source_files:
    filtered_pay_base = filtered_pay_base[filtered_pay_base["Source File"].isin(sel_source_files)]
if sel_suppliers:
    filtered_pay_base = filtered_pay_base[filtered_pay_base["Supplier"].isin(sel_suppliers)]
if sel_type_of_work:
    filtered_pay_base = filtered_pay_base[filtered_pay_base["Type of work"].isin(sel_type_of_work)]
if sel_diverse_expense:
    filtered_pay_base = filtered_pay_base[filtered_pay_base["Diverse expense"].isin(sel_diverse_expense)]
if sel_provinces:
    filtered_pay_base = filtered_pay_base[filtered_pay_base["Province"].isin(sel_provinces)]

# Actuals aggregated after payment filters
actuals = (
    filtered_pay_base.groupby(
        ["Source File", "Building Address", "Category"],
        as_index=False
    )["Real"].sum()
)

# ============================================================
# COMPARE
# ============================================================
compare = budgets.merge(
    actuals,
    on=["Source File", "Building Address", "Category"],
    how="left"
)

compare["Real"] = compare["Real"].fillna(0)
compare["Budget"] = compare["Budget"].fillna(0)
compare["Variance"] = compare["Budget"] - compare["Real"]
compare["% Used"] = compare.apply(
    lambda r: (r["Real"] / r["Budget"]) if r["Budget"] else 0.0,
    axis=1,
)
compare["Pending to Reach Budget"] = (compare["Budget"] - compare["Real"]).clip(lower=0)
compare["Status"] = compare.apply(
    lambda r: status_semaphore(r["% Used"], r["Budget"], r["Real"]),
    axis=1,
)

# Extra payment metadata aggregated for display
meta_by_building_cat = (
    filtered_pay_base.groupby(["Source File", "Building Address", "Category"], as_index=False)
    .agg(
        Supplier=("Supplier", lambda s: " | ".join(sorted(set([x for x in s if clean_text_value(x)]))[:10])),
        Type_of_work=("Type of work", lambda s: " | ".join(sorted(set([x for x in s if clean_text_value(x)]))[:10])),
        Diverse_expense=("Diverse expense", lambda s: " | ".join(sorted(set([x for x in s if clean_text_value(x)]))[:10])),
        Province=("Province", lambda s: " | ".join(sorted(set([x for x in s if clean_text_value(x)]))[:10])),
    )
)

compare = compare.merge(
    meta_by_building_cat,
    on=["Source File", "Building Address", "Category"],
    how="left"
)

# ============================================================
# ADDITIONAL COMPARE FILTERS
# ============================================================
all_companies = sorted(compare["Company"].dropna().unique().tolist())
sel_companies = st.sidebar.multiselect("Company", all_companies, default=[])

all_buildings = sorted(compare["Building Address"].dropna().unique().tolist())
sel_buildings = st.sidebar.multiselect("Building Address", all_buildings, default=[])

all_categories = sorted(compare["Category"].dropna().unique().tolist())
sel_categories = st.sidebar.multiselect("Category", all_categories, default=[])

all_statuses = sorted(compare["Status"].dropna().unique().tolist())
sel_statuses = st.sidebar.multiselect("Status", all_statuses, default=[])

view = compare.copy()
if sel_source_files:
    view = view[view["Source File"].isin(sel_source_files)]
if sel_companies:
    view = view[view["Company"].isin(sel_companies)]
if sel_buildings:
    view = view[view["Building Address"].isin(sel_buildings)]
if sel_categories:
    view = view[view["Category"].isin(sel_categories)]
if sel_statuses:
    view = view[view["Status"].isin(sel_statuses)]

# ============================================================
# BUILDING SUMMARY
# ============================================================
building_summary = (
    view.groupby(["Source File", "Building Address", "Company"], as_index=False)
    .agg(
        Real=("Real", "sum"),
        Budget=("Budget", "sum"),
        Pending_to_Reach_Budget=("Pending to Reach Budget", "sum"),
    )
)

building_summary["Variance"] = building_summary["Budget"] - building_summary["Real"]
building_summary["% Used"] = building_summary.apply(
    lambda r: (r["Real"] / r["Budget"]) if r["Budget"] else 0.0,
    axis=1,
)
building_summary["Status"] = building_summary.apply(
    lambda r: status_semaphore(r["% Used"], r["Budget"], r["Real"]),
    axis=1,
)

building_view = building_summary.copy()
if sel_statuses:
    building_view = building_view[building_view["Status"].isin(sel_statuses)]

# ============================================================
# EXECUTIVE SUMMARY
# ============================================================
st.subheader("📌 Executive Summary")

total_real = float(view["Real"].sum())
total_budget = float(view["Budget"].sum())
overall_used = (total_real / total_budget) if total_budget else 0.0
pending_total = max(total_budget - total_real, 0)

missing_real = building_view[
    (building_view["Budget"] > 0) & (building_view["Real"] <= 0)
].copy()

missing_real = missing_real.sort_values(
    ["Budget", "Pending_to_Reach_Budget"],
    ascending=[False, False]
)

k1, k2, k3, k4 = st.columns(4)
k1.metric("Total Real", fmt_currency(total_real))
k2.metric("Total Budget", fmt_currency(total_budget))
k3.metric("Budget Utilization", fmt_percent_ratio(overall_used))
k4.metric("Pending to Reach Budget", fmt_currency(pending_total))

left_col, right_col = st.columns([2.2, 1])

with left_col:
    st.markdown("### 🚨 Priority Addresses with No Real Yet")
    if not missing_real.empty:
        missing_real_display = missing_real[
            ["Status", "Source File", "Company", "Building Address", "Real", "Budget", "Pending_to_Reach_Budget", "% Used"]
        ].copy()

        st.dataframe(
            missing_real_display.style.format({
                "Real": fmt_currency,
                "Budget": fmt_currency,
                "Pending_to_Reach_Budget": fmt_currency,
                "% Used": fmt_percent_ratio,
            }),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.success("All budgeted addresses already have some Real recorded.")

with right_col:
    st.markdown("### 📊 Status Breakdown")
    status_breakdown = (
        view.groupby("Status", as_index=False)
        .agg(
            Buildings=("Building Address", "nunique"),
            Companies=("Company", "nunique"),
            Real=("Real", "sum"),
            Budget=("Budget", "sum"),
            Pending=("Pending to Reach Budget", "sum"),
        )
        .sort_values("Budget", ascending=False)
    )

    if not status_breakdown.empty:
        st.dataframe(
            status_breakdown.style.format({
                "Real": fmt_currency,
                "Budget": fmt_currency,
                "Pending": fmt_currency,
            }),
            use_container_width=True,
            hide_index=True,
        )

        fig_status = px.pie(
            status_breakdown,
            names="Status",
            values="Budget",
            title="Budget by Status",
            color="Status",
            color_discrete_map=STATUS_COLOR_MAP,
        )
        st.plotly_chart(fig_status, use_container_width=True)
    else:
        st.info("No data available for selected statuses.")

# ============================================================
# BREAKDOWNS REQUESTED
# ============================================================
st.subheader("🧩 Payments Breakdown")

b1, b2 = st.columns(2)

with b1:
    st.markdown("### Supplier Breakdown")
    supplier_breakdown = (
        filtered_pay_base.groupby("Supplier", as_index=False)
        .agg(
            Transactions=("Real", "size"),
            Real=("Real", "sum"),
            Buildings=("Building Address", "nunique"),
            Files=("Source File", "nunique"),
        )
        .sort_values("Real", ascending=False)
    )

    st.dataframe(
        supplier_breakdown.style.format({
            "Real": fmt_currency,
        }),
        use_container_width=True,
        hide_index=True,
    )

with b2:
    st.markdown("### Type of Work Breakdown")
    type_breakdown = (
        filtered_pay_base.groupby("Type of work", as_index=False)
        .agg(
            Transactions=("Real", "size"),
            Real=("Real", "sum"),
            Buildings=("Building Address", "nunique"),
            Files=("Source File", "nunique"),
        )
        .sort_values("Real", ascending=False)
    )

    st.dataframe(
        type_breakdown.style.format({
            "Real": fmt_currency,
        }),
        use_container_width=True,
        hide_index=True,
    )

b3, b4 = st.columns(2)

with b3:
    st.markdown("### Diverse Expense Breakdown")
    diverse_breakdown = (
        filtered_pay_base.groupby("Diverse expense", as_index=False)
        .agg(
            Transactions=("Real", "size"),
            Real=("Real", "sum"),
            Buildings=("Building Address", "nunique"),
            Files=("Source File", "nunique"),
        )
        .sort_values("Real", ascending=False)
    )

    st.dataframe(
        diverse_breakdown.style.format({
            "Real": fmt_currency,
        }),
        use_container_width=True,
        hide_index=True,
    )

with b4:
    st.markdown("### Province Breakdown")
    province_breakdown = (
        filtered_pay_base.groupby("Province", as_index=False)
        .agg(
            Transactions=("Real", "size"),
            Real=("Real", "sum"),
            Buildings=("Building Address", "nunique"),
            Files=("Source File", "nunique"),
        )
        .sort_values("Real", ascending=False)
    )

    st.dataframe(
        province_breakdown.style.format({
            "Real": fmt_currency,
        }),
        use_container_width=True,
        hide_index=True,
    )

# ============================================================
# COMPANY BREAKDOWN
# ============================================================
st.subheader("🏢 Company Breakdown")

company_breakdown = (
    view.groupby("Company", as_index=False)
    .agg(
        Real=("Real", "sum"),
        Budget=("Budget", "sum"),
        Pending=("Pending to Reach Budget", "sum"),
        Buildings=("Building Address", "nunique"),
        Files=("Source File", "nunique"),
    )
    .sort_values("Budget", ascending=False)
)

company_breakdown["% Used"] = company_breakdown.apply(
    lambda r: (r["Real"] / r["Budget"]) if r["Budget"] else 0.0,
    axis=1,
)

st.dataframe(
    company_breakdown.style.format({
        "Real": fmt_currency,
        "Budget": fmt_currency,
        "Pending": fmt_currency,
        "% Used": fmt_percent_ratio,
    }),
    use_container_width=True,
    hide_index=True,
)

# ============================================================
# DETAIL TABLE
# ============================================================
st.subheader("📋 Real vs Budget (by Building & Category)")

detail_columns = [
    "Status",
    "Source File",
    "Company",
    "Building Address",
    "Category",
    "Supplier",
    "Type_of_work",
    "Diverse_expense",
    "Province",
    "Real",
    "Budget",
    "Variance",
    "% Used",
    "Pending to Reach Budget",
]

detail_view_display = view[detail_columns].copy()
detail_view_display = detail_view_display.rename(
    columns={
        "Type_of_work": "Type of work",
        "Diverse_expense": "Diverse expense",
    }
)

st.dataframe(
    detail_view_display.style.format({
        "Real": fmt_currency,
        "Budget": fmt_currency,
        "Variance": fmt_currency,
        "Pending to Reach Budget": fmt_currency,
        "% Used": fmt_percent_ratio,
    }),
    use_container_width=True,
    hide_index=True,
)

# ============================================================
# BUILDING SUMMARY TABLE
# ============================================================
st.subheader("🏢 Building-Level Budget Tracking")

building_view = building_view.sort_values(
    ["Pending_to_Reach_Budget", "Budget"],
    ascending=[False, False]
)

st.dataframe(
    building_view[
        ["Status", "Source File", "Company", "Building Address", "Real", "Budget", "Variance", "Pending_to_Reach_Budget", "% Used"]
    ].style.format({
        "Real": fmt_currency,
        "Budget": fmt_currency,
        "Variance": fmt_currency,
        "Pending_to_Reach_Budget": fmt_currency,
        "% Used": fmt_percent_ratio,
    }),
    use_container_width=True,
    hide_index=True,
)

# ============================================================
# CHARTS
# ============================================================
st.subheader("📊 Executive Charts")

chart_building = building_view.head(15).copy()
if not chart_building.empty:
    chart_building_m = chart_building.melt(
        id_vars=["Building Address"],
        value_vars=["Real", "Budget"],
        var_name="Type",
        value_name="Amount",
    )
    fig_building = px.bar(
        chart_building_m,
        x="Building Address",
        y="Amount",
        color="Type",
        barmode="group",
        title="Top 15 Buildings: Real vs Budget",
    )
    fig_building.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig_building, use_container_width=True)

cat_summary = view.groupby("Category", as_index=False).agg(
    Real=("Real", "sum"),
    Budget=("Budget", "sum"),
)
if not cat_summary.empty:
    cat_summary_m = cat_summary.melt(
        id_vars=["Category"],
        value_vars=["Real", "Budget"],
        var_name="Type",
        value_name="Amount",
    )
    fig_cat = px.bar(
        cat_summary_m,
        x="Category",
        y="Amount",
        color="Type",
        barmode="group",
        title="Real vs Budget by Category",
    )
    st.plotly_chart(fig_cat, use_container_width=True)

supplier_chart = supplier_breakdown.head(15).copy()
if not supplier_chart.empty:
    fig_supplier = px.bar(
        supplier_chart,
        x="Supplier",
        y="Real",
        title="Top 15 Suppliers by Real",
    )
    fig_supplier.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig_supplier, use_container_width=True)

province_chart = province_breakdown.copy()
if not province_chart.empty:
    fig_province = px.pie(
        province_chart,
        names="Province",
        values="Real",
        title="Real by Province",
    )
    st.plotly_chart(fig_province, use_container_width=True)

# ============================================================
# DIAGNOSTICS
# ============================================================
with st.expander("🛠 Diagnostics"):
    st.write("Redirect URI used:", REDIRECT_URI)
    st.write("Allowed domain:", ALLOWED_DOMAIN)
    st.write("Signed in as:", signed_in_email)
    st.write("Selected Excel files:", selected_names)
    st.write("Link type:", "FOLDER" if is_folder else "FILE")
    st.write("Detected payments header rows (0-based):", payments_header_rows)
    st.write("Detected invoicing header rows (0-based):", invoicing_header_rows)
    st.write("Payments columns:", list(payments_df.columns))
    st.write("Invoicing columns:", list(invoicing_df.columns))
    st.write(
        "Detected in Payments:",
        {
            "Building Address": pay_addr,
            "Category": pay_cat,
            "Amount without taxes / Amount": pay_amt,
            "Supplier": pay_supplier,
            "Type of work": pay_type_of_work,
            "Diverse expense": pay_diverse_expense,
            "Province": pay_province,
            "Source File": pay_source_file,
        },
    )
    st.write(
        "Detected in Invoicing:",
        {
            "Building Address": inv_addr,
            "Company": inv_company,
            "Labor Budget": labor_budget_col,
            "Supplies Budget": supplies_budget_col,
            "Equipment Budget": equipment_budget_col,
            "PW Budget": pw_budget_col,
            "Source File": inv_source_file,
        },
    )
