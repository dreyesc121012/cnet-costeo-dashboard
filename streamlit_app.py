import base64
from io import BytesIO

import pandas as pd
import streamlit as st
import requests
import msal

# ============================================================
# CONFIG (Secrets)
# ============================================================
CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
ONEDRIVE_SHARED_URL = st.secrets["ONEDRIVE_SHARED_URL"]

# IMPORTANTE: sin slash final
REDIRECT_URI = st.secrets.get("REDIRECT_URI", "").strip().rstrip("/")

# ✅ SINGLE TENANT: usar tu TENANT_ID (NO /common)
TENANT_ID = st.secrets["TENANT_ID"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

SCOPES = ["User.Read", "Files.Read.All"] 

# Excel config
SHEET_REAL = "Real Master"
SHEET_FIXED = "Gasto Fijo"
HEADER_IDX = 6

st.set_page_config(page_title="CNET Costeo Dashboard", layout="wide")


# ============================================================
# HELPERS
# ============================================================
def _get_query_params() -> dict:
    """Devuelve query params como dict[str,str], compatible con varias versiones de Streamlit."""
    try:
        qp = st.query_params  # Streamlit nuevo
        out = {}
        for k in qp.keys():
            v = qp.get(k)
            if isinstance(v, list):
                out[k] = v[0] if v else ""
            else:
                out[k] = str(v) if v is not None else ""
        return out
    except Exception:
        pass

    try:
        qp = st.experimental_get_query_params()  # Streamlit viejo
        return {k: (v[0] if isinstance(v, list) and v else str(v)) for k, v in qp.items()}
    except Exception:
        return {}


def _clear_query_params():
    """Limpia la URL (quita ?code=...)."""
    try:
        st.query_params.clear()  # Streamlit nuevo
    except Exception:
        try:
            st.experimental_set_query_params()  # Streamlit viejo
        except Exception:
            pass


def make_share_id(shared_url: str) -> str:
    """Convierte shared URL a shareId para Microsoft Graph: shares/{shareId}/driveItem"""
    b = base64.b64encode(shared_url.encode("utf-8")).decode("utf-8")
    b = b.rstrip("=").replace("/", "_").replace("+", "-")
    return "u!" + b


def graph_get(url: str, access_token: str) -> requests.Response:
    return requests.get(
        url,
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=60
    )


def download_excel_bytes_from_shared_link(access_token: str, shared_url: str) -> bytes:
    """Descarga el archivo Excel desde un Shared Link (SharePoint/OneDrive)."""
    share_id = make_share_id(shared_url)

    meta_url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem"
    meta = graph_get(meta_url, access_token)

    if meta.status_code != 200:
        raise RuntimeError(
            f"Error resolviendo shared link: {meta.status_code}\n{meta.text}\n\n"
            f"TIP: Genera un link NUEVO (Share -> Copy link) y reemplaza ONEDRIVE_SHARED_URL."
        )

    meta_json = meta.json()
    item_id = meta_json["id"]
    drive_id = meta_json["parentReference"]["driveId"]

    content_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    file_r = graph_get(content_url, access_token)

    if file_r.status_code != 200:
        raise RuntimeError(f"Error descargando archivo: {file_r.status_code}\n{file_r.text}")

    return file_r.content


def read_excel_from_bytes(excel_bytes: bytes) -> pd.DataFrame:
    raw = pd.read_excel(BytesIO(excel_bytes), sheet_name=SHEET_REAL, header=None)

    headers = raw.iloc[HEADER_IDX].tolist()
    headers = [("Unnamed" if pd.isna(c) else str(c).strip()) for c in headers]

    # Uniquificar headers repetidos
    seen = {}
    out = []
    for c in headers:
        c = c if c else "Unnamed"
        if c in seen:
            seen[c] += 1
            out.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            out.append(c)

    df = raw.iloc[HEADER_IDX + 1:].copy()
    df.columns = out
    df = df.reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_fixed_total(excel_bytes: bytes) -> float:
    fx = pd.read_excel(BytesIO(excel_bytes), sheet_name=SHEET_FIXED, header=None)
    amounts = pd.to_numeric(fx.iloc[:, 2], errors="coerce")  # columna 3
    return float(amounts.fillna(0).sum())


def find_col(df: pd.DataFrame, name: str):
    if name in df.columns:
        return name
    n = name.strip().lower()
    for c in df.columns:
        if str(c).strip().lower() == n:
            return c
    for c in df.columns:
        if n in str(c).strip().lower():
            return c
    return None


def safe_pct(x: float, base: float) -> float:
    return (x / base) if base not in (0, None) else 0.0


# ============================================================
# MSAL APP
# ============================================================

def get_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        token_cache=None,   # 👈 importante
    )



# ============================================================
# UI
# ============================================================
st.title("📊 CNET Costeo & Neto Dashboard")

if not REDIRECT_URI:
    st.error("Falta REDIRECT_URI en Secrets. Ej: https://cnet-dashboard.streamlit.app (sin slash final).")
    st.stop()

app = get_msal_app()
qp = _get_query_params()

# ============================================================
# LOGIN (SIN FLOW, SIN STATE)
# ============================================================
if "token_result" not in st.session_state:

    # Si venimos del redirect con ?code=
    if qp.get("code"):
        try:
            result = app.acquire_token_by_authorization_code(
                code=qp["code"],
                scopes=SCOPES,
                redirect_uri=REDIRECT_URI,
            )
        except Exception as e:
            st.error(f"Error completando login: {e}")
            st.stop()

        if "access_token" in result:
            st.session_state.token_result = result
            _clear_query_params()
            st.rerun()
        else:
            st.error("No se pudo obtener access_token.")
            st.code(result)
            st.stop()

    # Si NO hay code -> mostrar link de login
    st.warning("No has iniciado sesión en OneDrive/SharePoint")

    auth_url = app.get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )

    st.markdown("### 🔐 Inicia sesión")
    st.link_button("Iniciar sesión OneDrive", auth_url)
    st.caption(f"Auth URL (debe decir /{TENANT_ID}/): {auth_url}")

    st.stop()

token_result = st.session_state.token_result

if "access_token" not in token_result:
    st.error("No se pudo obtener token válido.")
    st.code(token_result)
    st.stop()

st.success("✅ Conectado a OneDrive/SharePoint (token activo)")

# Refresh
if st.button("🔄 Refresh datos"):
    st.session_state.pop("excel_bytes", None)
    st.rerun()

# Logout
if st.button("🔒 Cerrar sesión"):
    for k in ["token_result", "excel_bytes"]:
        st.session_state.pop(k, None)
    _clear_query_params()
    st.rerun()

# ============================================================
# Descargar Excel y cargar datos
# ============================================================
try:
    if "excel_bytes" not in st.session_state:
        st.info("📥 Descargando Excel desde SharePoint/OneDrive…")
        st.session_state.excel_bytes = download_excel_bytes_from_shared_link(
            token_result["access_token"],
            ONEDRIVE_SHARED_URL
        )
    excel_bytes = st.session_state.excel_bytes
except Exception as e:
    st.error("No pude descargar el archivo desde OneDrive/SharePoint.")
    st.code(str(e))
    st.stop()

df = read_excel_from_bytes(excel_bytes)
fixed_total = load_fixed_total(excel_bytes)

# ============================================================
# TU LÓGICA (KPIs)
# ============================================================
COL_INCOME = "Total to Bill"
COL_COST   = "Total Cost Month"
COL_MGMT   = "Total Management Fee"
COL_ROY    = "Royalty CNET Group Inc 5%"

c_income = find_col(df, COL_INCOME)
c_cost   = find_col(df, COL_COST)
c_mgmt   = find_col(df, COL_MGMT)
c_roy    = find_col(df, COL_ROY)

missing = [k for k, v in {
    COL_INCOME: c_income,
    COL_COST: c_cost,
    COL_MGMT: c_mgmt,
    COL_ROY: c_roy,
}.items() if v is None]

if missing:
    st.error(f"No encontré estas columnas en 'Real Master': {missing}")
    with st.expander("Ver columnas detectadas"):
        st.write(df.columns.tolist())
    st.stop()

df[c_income] = pd.to_numeric(df[c_income], errors="coerce")
df[c_cost]   = pd.to_numeric(df[c_cost], errors="coerce")
df[c_mgmt]   = pd.to_numeric(df[c_mgmt], errors="coerce")
df[c_roy]    = pd.to_numeric(df[c_roy], errors="coerce")

income = float(df[c_income].fillna(0).sum())
cost = float(df[c_cost].fillna(0).sum())
gross = income - cost
mgmt_fee_total = float(df[c_mgmt].fillna(0).sum())
royalty_total = float(df[c_roy].fillna(0).sum())
net = gross - fixed_total
new_total = net + mgmt_fee_total + royalty_total

p_cost  = safe_pct(cost, income)
p_gross = safe_pct(gross, income)
p_fixed = safe_pct(fixed_total, income)
p_net   = safe_pct(net, income)
p_mgmt  = safe_pct(mgmt_fee_total, income)
p_roy   = safe_pct(royalty_total, income)
p_new   = safe_pct(new_total, income)

st.subheader("📌 KPIs (Ejecutivo)")
k1, k2, k3, k4 = st.columns(4)
k1.metric("Ingresos (Total to Bill)", f"${income:,.2f}")
k2.metric("Costos (Total Cost Month)", f"${cost:,.2f}", f"{p_cost*100:,.2f}%")
k3.metric("Gross (Ingreso - Costo)", f"${gross:,.2f}", f"{p_gross*100:,.2f}%")
k4.metric("Gastos fijos (Gasto Fijo)", f"${fixed_total:,.2f}", f"{p_fixed*100:,.2f}%")

st.subheader("Resumen")
summary = pd.DataFrame([
    ["Ingresos", income, 1.0],
    ["Costos", cost, p_cost],
    ["Gross", gross, p_gross],
    ["Gastos fijos", fixed_total, p_fixed],
    ["Neto", net, p_net],
    ["Total Management Fee", mgmt_fee_total, p_mgmt],
    ["Royalty CNET Group Inc 5%", royalty_total, p_roy],
    ["Nuevo Total", new_total, p_new],
], columns=["Concepto", "Monto", "% sobre Ingresos"])

summary["Monto"] = summary["Monto"].map(lambda x: f"${x:,.2f}")
summary["% sobre Ingresos"] = summary["% sobre Ingresos"].map(lambda x: f"{x*100:,.2f}%")

st.dataframe(summary, use_container_width=True)
