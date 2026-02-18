import base64
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st
import requests
import msal
import plotly.graph_objects as go

# PDF
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# ============================================================
# CONFIG (Secrets)
# ============================================================
CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
ONEDRIVE_SHARED_URL = st.secrets["ONEDRIVE_SHARED_URL"]

REDIRECT_URI = st.secrets.get("REDIRECT_URI", "").strip().rstrip("/")

TENANT_ID = st.secrets["TENANT_ID"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

SCOPES = ["User.Read", "Files.Read.All"]

SHEET_REAL = "Real Master"
SHEET_FIXED = "Gasto Fijo"
HEADER_IDX = 6

MONTH_COL = "Month"
YEAR_COL = "Year"

st.set_page_config(page_title="CNET Costeo Dashboard", layout="wide")

# ============================================================
# HELPERS
# ============================================================

def make_share_id(shared_url: str) -> str:
    b = base64.b64encode(shared_url.encode("utf-8")).decode("utf-8")
    b = b.rstrip("=").replace("/", "_").replace("+", "-")
    return "u!" + b

def graph_get(url: str, access_token: str):
    return requests.get(
        url,
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=60
    )

def download_excel_bytes_from_shared_link(access_token: str, shared_url: str) -> bytes:
    share_id = make_share_id(shared_url)

    meta_url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem"
    meta = graph_get(meta_url, access_token)

    if meta.status_code != 200:
        raise RuntimeError(f"Error resolviendo shared link:\n{meta.text}")

    meta_json = meta.json()
    item_id = meta_json["id"]
    drive_id = meta_json["parentReference"]["driveId"]

    content_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    file_r = graph_get(content_url, access_token)

    if file_r.status_code != 200:
        raise RuntimeError(f"Error descargando archivo:\n{file_r.text}")

    return file_r.content

def make_unique_columns(cols):
    seen = {}
    out = []
    for c in cols:
        c = "Unnamed" if pd.isna(c) else str(c).strip()
        if c in seen:
            seen[c] += 1
            out.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            out.append(c)
    return out

@st.cache_data(ttl=300)
def read_real_master_from_bytes(excel_bytes: bytes):
    raw = pd.read_excel(BytesIO(excel_bytes), sheet_name=SHEET_REAL, header=None)
    headers = make_unique_columns(raw.iloc[HEADER_IDX].tolist())
    df = raw.iloc[HEADER_IDX + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)
    return df

@st.cache_data(ttl=300)
def load_fixed_total_from_bytes(excel_bytes: bytes):
    fx = pd.read_excel(BytesIO(excel_bytes), sheet_name=SHEET_FIXED, header=None)
    amounts = pd.to_numeric(fx.iloc[:, 2], errors="coerce")
    return float(amounts.fillna(0).sum())

def find_col(df, name):
    if name in df.columns:
        return name
    for c in df.columns:
        if str(c).strip().lower() == name.lower():
            return c
    return None

def safe_pct(x, base):
    return (x / base) if base not in (0, None) else 0

def sanitize_for_arrow(df):
    df2 = df.copy()
    for col in df2.columns:
        if df2[col].dtype == "object":
            df2[col] = df2[col].astype(str)
    return df2

# ============================================================
# MONTH + YEAR TEXT HANDLING
# ============================================================

_MONTH_MAP = {
    "jan":1,"january":1,"feb":2,"february":2,"mar":3,"march":3,
    "apr":4,"apil":4,"may":5,"jun":6,"june":6,"jul":7,"july":7,
    "aug":8,"august":8,"sep":9,"sept":9,"september":9,
    "oct":10,"october":10,"nov":11,"november":11,
    "dec":12,"december":12
}

def build_month_fields(df):
    out = df.copy()

    out["_YearInt"] = pd.to_numeric(out[YEAR_COL], errors="coerce")
    month_raw = out[MONTH_COL].astype(str).str.strip().str.lower()
    month_clean = month_raw.str.replace(r"[^a-z]", "", regex=True)
    out["_MonthNum"] = month_clean.map(_MONTH_MAP)

    out["_MonthKey"] = (
        out["_YearInt"].astype("Int64").astype(str) + "-" +
        out["_MonthNum"].astype("Int64").astype(str).str.zfill(2)
    )

    out["_MonthText"] = (
        out[MONTH_COL].astype(str).str.strip() + " " +
        out["_YearInt"].astype("Int64").astype(str)
    )

    bad = out["_MonthNum"].isna() | out["_YearInt"].isna()
    out.loc[bad, ["_MonthKey","_MonthText"]] = pd.NA

    return out

# ============================================================
# FILTERS
# ============================================================

def add_filters(df):
    st.sidebar.header("Filtros Ejecutivos")

    if MONTH_COL in df.columns and YEAR_COL in df.columns:
        df = build_month_fields(df)

        years = sorted([int(y) for y in df["_YearInt"].dropna().unique()])
        sel_y = st.sidebar.multiselect("Year", years, default=years)
        if sel_y:
            df = df[df["_YearInt"].isin(sel_y)]

        month_table = (
            df[["_MonthKey","_MonthText"]]
            .dropna()
            .drop_duplicates()
            .sort_values("_MonthKey")
        )

        sel_m = st.sidebar.multiselect(
            "Month",
            month_table["_MonthText"].tolist()
        )

        if sel_m:
            df = df[df["_MonthText"].isin(sel_m)]

    return df

# ============================================================
# MSAL LOGIN
# ============================================================

def get_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        token_cache=None
    )

st.title("📊 CNET Costeo & Neto Dashboard")

app = get_msal_app()

if "token_result" not in st.session_state:
    auth_url = app.get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )
    st.link_button("Iniciar sesión OneDrive", auth_url)
    st.stop()

token_result = st.session_state.token_result

# ============================================================
# LOAD DATA
# ============================================================

excel_bytes = download_excel_bytes_from_shared_link(
    token_result["access_token"],
    ONEDRIVE_SHARED_URL
)

df_all = read_real_master_from_bytes(excel_bytes)
fixed_total = load_fixed_total_from_bytes(excel_bytes)

df = add_filters(df_all.copy())

# ============================================================
# KPIs
# ============================================================

COL_INCOME = "Total to Bill"
COL_COST   = "Total Cost Month"
COL_MGMT   = "Total Management Fee"
COL_ROY    = "Royalty CNET Group Inc 5%"

df[COL_INCOME] = pd.to_numeric(df[COL_INCOME], errors="coerce")
df[COL_COST]   = pd.to_numeric(df[COL_COST], errors="coerce")
df[COL_MGMT]   = pd.to_numeric(df[COL_MGMT], errors="coerce")
df[COL_ROY]    = pd.to_numeric(df[COL_ROY], errors="coerce")

income = df[COL_INCOME].sum()
cost = df[COL_COST].sum()
gross = income - cost
mgmt = df[COL_MGMT].sum()
roy = df[COL_ROY].sum()
net = gross - fixed_total
new_total = net + mgmt + roy

# ============================================================
# BREAKDOWN POR MES (SIN HORAS)
# ============================================================

st.subheader("🗓️ Breakdown por Mes (filtrado)")

if "_MonthKey" not in df.columns:
    df = build_month_fields(df)

group = (
    df.dropna(subset=["_MonthKey"])
      .groupby(["_MonthKey","_MonthText"])
      .agg(
          Income=(COL_INCOME,"sum"),
          Cost=(COL_COST,"sum"),
          Mgmt=(COL_MGMT,"sum"),
          Royalty=(COL_ROY,"sum"),
      )
      .reset_index()
      .sort_values("_MonthKey")
)

group["Gross"] = group["Income"] - group["Cost"]
group["New Total"] = group["Gross"] - fixed_total + group["Mgmt"] + group["Royalty"]

x_text = group["_MonthText"].tolist()

fig = go.Figure()
fig.add_trace(go.Bar(name="Income", x=x_text, y=group["Income"]))
fig.add_trace(go.Bar(name="Cost", x=x_text, y=group["Cost"]))
fig.add_trace(go.Scatter(name="New Total", x=x_text, y=group["New Total"], mode="lines+markers"))

fig.update_layout(
    title="Mes a Mes (filtrado): Income vs Cost + New Total",
    barmode="group",
    xaxis_title="Month",
    yaxis_title="Amount"
)

fig.update_xaxes(type="category", categoryorder="array", categoryarray=x_text)

st.plotly_chart(fig, use_container_width=True)

# ============================================================
# SUMMARY TABLE
# ============================================================

st.subheader("Resumen Ejecutivo")

summary = pd.DataFrame([
    ["Ingresos", income],
    ["Costos", cost],
    ["Gross", gross],
    ["Gasto Fijo", fixed_total],
    ["Nuevo Total", new_total],
], columns=["Concepto","Monto"])

summary["Monto"] = summary["Monto"].map(lambda x: f"${x:,.2f}")

st.dataframe(summary, use_container_width=True)
