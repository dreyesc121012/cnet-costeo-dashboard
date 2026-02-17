import os
import json
import base64
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st
import requests
import msal
import plotly.graph_objects as go

# -------------------------
# CONFIG (Secrets)
# -------------------------
CLIENT_ID = st.secrets["CLIENT_ID"]
TENANT_ID = st.secrets["TENANT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]

ONEDRIVE_SHARED_URL = st.secrets["ONEDRIVE_SHARED_URL"]
REDIRECT_URI = st.secrets.get("REDIRECT_URI")  # ej: https://cnet-dashboard.streamlit.app

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.Read.All"]  # Delegated

# Tu Excel
SHEET_REAL = "Real Master"
SHEET_FIXED = "Gasto Fijo"
HEADER_IDX = 6

st.set_page_config(page_title="CNET Costeo Dashboard", layout="wide")

# -------------------------
# TOKEN CACHE (persistente)
# -------------------------
CACHE_FILE = ".msal_cache.bin"

def load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        cache.deserialize(open(CACHE_FILE, "r").read())
    return cache

def save_cache(cache):
    if cache.has_state_changed:
        open(CACHE_FILE, "w").write(cache.serialize())

@st.cache_resource
def get_msal_app():
    cache = load_cache()
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        token_cache=cache,
    )
    return app, cache

def get_token_silent():
    app, cache = get_msal_app()
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        save_cache(cache)
        return result
    return None

def make_share_id(shared_url: str) -> str:
    """
    Convierte shared URL a shareId para Graph:
    shares/{shareId}/driveItem
    """
    b = base64.b64encode(shared_url.encode("utf-8")).decode("utf-8")
    b = b.rstrip("=").replace("/", "_").replace("+", "-")
    return "u!" + b

def graph_get(url, access_token):
    r = requests.get(url, headers={"Authorization": f"Bearer {access_token}"})
    return r

def download_excel_bytes_from_shared_link(access_token: str, shared_url: str) -> bytes:
    share_id = make_share_id(shared_url)

    # 1) Resolver el shared link a driveItem
    meta_url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem"
    meta = graph_get(meta_url, access_token)
    if meta.status_code != 200:
        raise RuntimeError(f"Error resolviendo shared link: {meta.status_code} {meta.text}")

    meta_json = meta.json()
    item_id = meta_json["id"]
    drive_id = meta_json["parentReference"]["driveId"]

    # 2) Descargar contenido
    content_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    file_r = graph_get(content_url, access_token)
    if file_r.status_code != 200:
        raise RuntimeError(f"Error descargando archivo: {file_r.status_code} {file_r.text}")

    return file_r.content

def read_excel_from_bytes(excel_bytes: bytes) -> pd.DataFrame:
    raw = pd.read_excel(BytesIO(excel_bytes), sheet_name=SHEET_REAL, header=None)
    headers = raw.iloc[HEADER_IDX].tolist()
    headers = [("Unnamed" if pd.isna(c) else str(c).strip()) for c in headers]

    # unique
    seen = {}
    out = []
    for c in headers:
        if c == "":
            c = "Unnamed"
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
    amounts = pd.to_numeric(fx.iloc[:, 2], errors="coerce")
    return float(amounts.fillna(0).sum())

# -------------------------
# UI
# -------------------------
st.title("📊 CNET Costeo & Neto Dashboard")

# 1) Intentar token silencioso
token_result = get_token_silent()

# 2) Si no hay token, iniciar login (auth code flow)
if not token_result:
    st.warning("No has iniciado sesión en OneDrive")

    app, cache = get_msal_app()

    # iniciar flow una sola vez
    if "flow" not in st.session_state:
        st.session_state.flow = app.initiate_auth_code_flow(
            scopes=SCOPES,
            redirect_uri=REDIRECT_URI
        )

    auth_url = st.session_state.flow["auth_uri"]

    st.markdown(f"### 🔐 Inicia sesión")
    st.link_button("Iniciar sesión OneDrive", auth_url)

    # Capturar code cuando Microsoft redirige a tu app
    qp = st.query_params
    if "code" in qp:
        try:
            result = app.acquire_token_by_auth_code_flow(
                st.session_state.flow,
                dict(qp),
                scopes=SCOPES,
            )
            save_cache(cache)
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"Error completando login: {e}")

    st.stop()

# Si hay token
if "access_token" not in token_result:
    st.error(f"No se pudo obtener token: {token_result}")
    st.stop()

st.success("Conectado a OneDrive (token activo)")

if st.button("🔄 Refresh datos"):
    st.cache_data.clear()
    st.rerun()

# -------------------------
# Descargar Excel y cargar datos
# -------------------------
try:
    excel_bytes = download_excel_bytes_from_shared_link(token_result["access_token"], ONEDRIVE_SHARED_URL)
except Exception as e:
    st.error("No pude descargar el archivo desde OneDrive/SharePoint.")
    st.code(str(e))
    st.stop()

df = read_excel_from_bytes(excel_bytes)
fixed_total = load_fixed_total(excel_bytes)

# -------------------------
# TU LOGICA (tu dashboard actual)
# -------------------------
def find_col(df, name):
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

def safe_pct(x, base):
    return (x / base) if base not in (0, None) else 0.0

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

st.subheader("📊 KPIs (Ejecutivo)")
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
