import base64
from io import BytesIO
from datetime import timedelta, datetime

import pandas as pd
import requests
import streamlit as st
import msal

st.set_page_config(page_title="Comité Paritaire QC", layout="wide")
st.title("Comité Paritaire Québec - Weekly Report")

# ============================================================
# CONFIG (Secrets)
# ============================================================
CLIENT_ID = str(st.secrets["CLIENT_ID"]).strip()
CLIENT_SECRET = str(st.secrets["CLIENT_SECRET"]).strip()
TENANT_ID = str(st.secrets["TENANT_ID"]).strip()
REDIRECT_URI = str(st.secrets["REDIRECT_URI"]).strip().rstrip("/")
ONEDRIVE_FOLDER_URL = str(st.secrets["ONEDRIVE_FOLDER_URL"]).strip()
DOMAIN_HINT = str(st.secrets.get("DOMAIN_HINT", "")).strip()
LOGIN_HINT = str(st.secrets.get("LOGIN_HINT", "")).strip()
ALLOWED_DOMAIN = str(st.secrets.get("ALLOWED_DOMAIN", "")).strip().lower()

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.Read.All", "Sites.Read.All"]

# ============================================================
# URL PARAM HELPERS
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

st.sidebar.success(f"Logged in as {signed_in_email}")
st.success(f"✅ Signed in as {signed_in_email}")

if st.button("🚪 Sign out"):
    st.session_state.pop("token_result", None)
    st.session_state.pop("selected_item_ids", None)
    st.session_state.pop("cloud_files", None)
    clear_query_params_compat()
    st.rerun()

# ============================================================
# LOAD FILE LIST FROM ONEDRIVE FOLDER
# ============================================================
try:
    meta = resolve_shared_link(access_token, ONEDRIVE_FOLDER_URL)
except Exception as e:
    st.error("Could not resolve the SharePoint/OneDrive folder link.")
    st.code(str(e))
    st.stop()

drive_id = meta["parentReference"]["driveId"]
root_item_id = meta["id"]
is_folder = "folder" in meta

if not is_folder:
    st.error("ONEDRIVE_FOLDER_URL must be a folder link, not a file link.")
    st.stop()

try:
    children = list_children_all(access_token, drive_id, root_item_id)
    cloud_files = [c for c in children if c.get("id") and is_excel_name(c.get("name", ""))]
    cloud_files.sort(key=lambda x: (x.get("name") or "").lower())
except Exception as e:
    st.error("Could not list Excel files in the OneDrive folder.")
    st.code(str(e))
    st.stop()

if not cloud_files:
    st.warning("No Excel files found in the OneDrive folder.")
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

def load_selected_cloud_files(selected_files):
    dataframes = []

    for file_info in selected_files:
        try:
            file_bytes = BytesIO(download_item_bytes(access_token, drive_id, file_info["id"]))
            df = pd.read_excel(file_bytes, sheet_name="data")
            df.columns = [str(c).strip() for c in df.columns]
            df["source_file"] = file_info["name"]
            dataframes.append(df)
        except Exception as e:
            st.warning(f"Could not read {file_info['name']}: {e}")

    if dataframes:
        return pd.concat(dataframes, ignore_index=True)
    return pd.DataFrame()

df = load_selected_cloud_files(selected_files)

if df.empty:
    st.error("No valid data could be loaded from the selected files.")
    st.stop()
