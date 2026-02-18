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

# Excel config
SHEET_REAL = "Real Master"
SHEET_FIXED = "Gasto Fijo"
HEADER_IDX = 6

# Columna exacta de mes
MONTH_COL = "Month"   # texto: "January", "February", etc.

st.set_page_config(page_title="CNET Costeo Dashboard", layout="wide")

# ============================================================
# HELPERS (Streamlit URL params)
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
        pass

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
# HELPERS (Graph download)
# ============================================================
def make_share_id(shared_url: str) -> str:
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

# ============================================================
# HELPERS (Excel parsing)
# ============================================================
def make_unique_columns(cols):
    seen = {}
    out = []
    for c in cols:
        c = "Unnamed" if pd.isna(c) else str(c).strip()
        if c == "":
            c = "Unnamed"
        if c in seen:
            seen[c] += 1
            out.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            out.append(c)
    return out

@st.cache_data(ttl=300, show_spinner=False)
def read_real_master_from_bytes(excel_bytes: bytes) -> pd.DataFrame:
    raw = pd.read_excel(BytesIO(excel_bytes), sheet_name=SHEET_REAL, header=None)
    headers = make_unique_columns(raw.iloc[HEADER_IDX].tolist())

    df = raw.iloc[HEADER_IDX + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    return df

@st.cache_data(ttl=300, show_spinner=False)
def load_fixed_total_from_bytes(excel_bytes: bytes) -> float:
    fx = pd.read_excel(BytesIO(excel_bytes), sheet_name=SHEET_FIXED, header=None)
    amounts = pd.to_numeric(fx.iloc[:, 2], errors="coerce")
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

def sanitize_for_arrow(df: pd.DataFrame) -> pd.DataFrame:
    df2 = df.copy()
    for col in df2.columns:
        if df2[col].dtype == "object":
            df2[col] = df2[col].astype(str)
    return df2

def pick_building_col(df: pd.DataFrame):
    candidates = ["Building", "Building ID", "BuildinG ID", "Building Name", "Site", "Location", "Branch"]
    for name in candidates:
        c = find_col(df, name)
        if c:
            return c
    for c in df.columns:
        if "build" in str(c).strip().lower():
            return c
    return None

# ============================================================
# MONTH TEXT -> YYYY-MM (needs Year)
# ============================================================
_MONTH_MAP = {
    "jan": "January", "january": "January",
    "feb": "February", "february": "February",
    "mar": "March", "march": "March",
    "apr": "April", "april": "April",
    "may": "May",
    "jun": "June", "june": "June",
    "jul": "July", "july": "July",
    "aug": "August", "august": "August",
    "sep": "September", "sept": "September", "september": "September",
    "oct": "October", "october": "October",
    "nov": "November", "november": "November",
    "dec": "December", "december": "December",
}

def month_text_to_month_label(s: pd.Series, year: int) -> pd.Series:
    """
    Convierte Month texto ('January') + year -> 'YYYY-MM'.
    Soporta 'Jan', 'January', 'January 2026', etc.
    """
    ss = s.astype(str).str.strip()

    # Si trae algo como "January 2026" o "Jan-2026", intentamos parse directo primero
    dt_direct = pd.to_datetime(ss, errors="coerce", infer_datetime_format=True)
    out = dt_direct.dt.to_period("M").astype(str)

    # Donde no pudo parsear, normalizamos por nombre de mes + year
    mask_bad = dt_direct.isna() & ss.notna()
    if mask_bad.any():
        raw = ss[mask_bad].str.lower()

        # queda solo palabra principal (January, Jan)
        raw = raw.str.replace(r"[^a-z]", " ", regex=True).str.split().str[0].fillna("")

        norm = raw.map(_MONTH_MAP)
        # armamos fecha "January 1, 2026"
        assembled = norm.fillna("") + " 1 " + str(year)
        dt2 = pd.to_datetime(assembled, errors="coerce")
        out.loc[mask_bad] = dt2.dt.to_period("M").astype(str)

    out = out.replace("NaT", pd.NA)
    return out

# ============================================================
# FILTERS (Month fijo + Year selector)
# ============================================================
def add_filters(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    st.sidebar.header("Filtros Ejecutivos")

    # Year selector (porque Month viene sin año)
    default_year = datetime.now().year
    year = st.sidebar.number_input("Año (Year) para Month", min_value=2000, max_value=2100, value=default_year, step=1)

    # Month filter (Month texto + Year)
    if MONTH_COL in df.columns:
        month_labels = month_text_to_month_label(df[MONTH_COL], year=int(year))
        valid_months = sorted([m for m in month_labels.dropna().unique().tolist() if m])
        if valid_months:
            sel_m = st.sidebar.multiselect("Mes (YYYY-MM)", valid_months, default=[])
            if sel_m:
                df = df[month_labels.isin(sel_m)]
    else:
        st.sidebar.warning("No encontré la columna Month para filtrar por mes.")

    # Company
    c_company = find_col(df, "Company")
    if c_company:
        sel = st.sidebar.multiselect("Company", sorted(df[c_company].dropna().unique()))
        if sel:
            df = df[df[c_company].isin(sel)]

    # Province
    c_prov = find_col(df, "Province")
    if c_prov:
        sel = st.sidebar.multiselect("Province", sorted(df[c_prov].dropna().unique()))
        if sel:
            df = df[df[c_prov].isin(sel)]

    # Client
    c_client = find_col(df, "Client")
    if c_client:
        sel = st.sidebar.multiselect("Client", sorted(df[c_client].dropna().unique()))
        if sel:
            df = df[df[c_client].isin(sel)]

    # Project Name
    c_proj = find_col(df, "Project Name")
    if c_proj:
        sel = st.sidebar.multiselect("Project (Project Name)", sorted(df[c_proj].dropna().unique()))
        if sel:
            df = df[df[c_proj].isin(sel)]

    # Building
    c_bld = pick_building_col(df)
    if c_bld:
        uniq = sorted(df[c_bld].dropna().astype(str).unique().tolist())
        sel = st.sidebar.multiselect("Building", uniq)
        if sel:
            df = df[df[c_bld].astype(str).isin(sel)]

    return df, int(year)

# ============================================================
# PDF REPORT
# ============================================================
def build_pdf_report(
    *,
    income, cost, gross, fixed_total, net, mgmt_fee_total, royalty_total, new_total,
    p_cost, p_gross, p_fixed, p_net, p_mgmt, p_roy, p_new,
    target, gross_margin, net_margin, final_margin,
    fig_waterfall=None, fig_gauge=None,
):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    width, height = letter

    y = height - 50
    c.setFont("Helvetica-Bold", 16)
    c.drawString(40, y, "CNET Costeo & Neto - Executive Summary")
    y -= 16
    c.setFont("Helvetica", 9)
    c.drawString(40, y, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    y -= 10
    c.drawString(40, y, f"Target Margin: {target:.0%}")
    y -= 20

    rows = [
        ("Ingresos (Total to Bill)", income, 1.0),
        ("Costos (Total Cost Month)", cost, p_cost),
        ("Gross (Ingreso - Costo)", gross, p_gross),
        ("Gastos fijos (Gasto Fijo)", fixed_total, p_fixed),
        ("Neto (Gross - Fijos)", net, p_net),
        ("Management Fee", mgmt_fee_total, p_mgmt),
        ("Royalty 5%", royalty_total, p_roy),
        ("Nuevo Total", new_total, p_new),
    ]

    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "KPIs")
    y -= 14

    c.setFont("Helvetica", 10)
    c.drawString(40, y, "Concepto")
    c.drawRightString(360, y, "Monto")
    c.drawRightString(520, y, "% Ingresos")
    y -= 12

    for label, val, pct in rows:
        c.drawString(40, y, label[:55])
        c.drawRightString(360, y, f"${val:,.2f}")
        c.drawRightString(520, y, f"{pct:.1%}")
        y -= 14

    y -= 8
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "Executive Margins")
    y -= 14
    c.setFont("Helvetica", 10)
    c.drawString(40, y, f"Gross Margin: {gross_margin:.1%}")
    y -= 14
    c.drawString(40, y, f"Net Margin: {net_margin:.1%}")
    y -= 14
    c.drawString(40, y, f"Final Margin (after fees): {final_margin:.1%}")
    y -= 18

    def add_plotly_image(fig, title, y_top):
        if fig is None:
            return y_top
        try:
            img_bytes = fig.to_image(format="png")  # requiere kaleido
            img = ImageReader(BytesIO(img_bytes))
            c.setFont("Helvetica-Bold", 11)
            c.drawString(40, y_top, title)
            y_top -= 10
            c.drawImage(img, 40, y_top - 220, width=520, height=220, preserveAspectRatio=True, mask='auto')
            return y_top - 235
        except Exception:
            c.setFont("Helvetica", 9)
            c.drawString(40, y_top, f"{title} (no se pudo exportar imagen - instala kaleido)")
            return y_top - 15

    y = add_plotly_image(fig_gauge, "Gauge - Final Margin", y)
    y = add_plotly_image(fig_waterfall, "Cascada - Ingresos → Costos → Fijos → Fees → Total", y)

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.getvalue()

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
# UI + LOGIN
# ============================================================
st.title("📊 CNET Costeo & Neto Dashboard")

if not REDIRECT_URI:
    st.error("Falta REDIRECT_URI en Secrets. Ej: https://cnet-dashboard.streamlit.app (sin slash final).")
    st.stop()

app = get_msal_app()
qp = _get_query_params()

if "token_result" not in st.session_state:

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

# Header actions
colA, colB = st.columns([1, 3])
with colA:
    if st.button("🔄 Refresh datos"):
        st.session_state.pop("excel_bytes", None)
        read_real_master_from_bytes.clear()
        load_fixed_total_from_bytes.clear()
        st.rerun()

with colB:
    if st.button("🔒 Cerrar sesión"):
        for k in ["token_result", "excel_bytes"]:
            st.session_state.pop(k, None)
        _clear_query_params()
        st.rerun()

# ============================================================
# Download + Load
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

df_all = read_real_master_from_bytes(excel_bytes)
fixed_total = load_fixed_total_from_bytes(excel_bytes)

# aplicar filtros (incluye Month + Year)
df, selected_year = add_filters(df_all.copy())

# ============================================================
# KPIs base
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

# ============================================================
# SEMÁFORO + GAUGE
# ============================================================
st.subheader("📌 Margen Ejecutivo (KPIs + Semáforo)")

target = st.slider("Margen objetivo (%)", 0, 60, 25) / 100
zona_amarilla = 0.05

gross_margin = gross / income if income else 0
net_margin   = net / income if income else 0
final_margin = new_total / income if income else 0

def semaforo(m, tgt):
    if m >= tgt + zona_amarilla:
        return "🟢"
    elif m >= tgt - zona_amarilla:
        return "🟡"
    return "🔴"

c1, c2, c3 = st.columns(3)
c1.metric("Gross Margin", f"{gross_margin:.1%}", f"{semaforo(gross_margin, target)} vs {target:.0%}")
c2.metric("Net Margin", f"{net_margin:.1%}", f"{semaforo(net_margin, target)} vs {target:.0%}")
c3.metric("Final Margin (after fees)", f"{final_margin:.1%}", f"{semaforo(final_margin, target)} vs {target:.0%}")

st.caption("Gauge: Margen final (después de fees)")
gauge_max = 60
fig_gauge = go.Figure(go.Indicator(
    mode="gauge+number",
    value=float(final_margin * 100),
    number={"suffix": "%"},
    gauge={
        "axis": {"range": [0, gauge_max]},
        "threshold": {"line": {"width": 4}, "value": float(target * 100)},
        "steps": [
            {"range": [0, max(0, (target - zona_amarilla) * 100)]},
            {"range": [max(0, (target - zona_amarilla) * 100), (target + zona_amarilla) * 100]},
            {"range": [(target + zona_amarilla) * 100, gauge_max]},
        ],
    }
))
st.plotly_chart(fig_gauge, use_container_width=True)

# ============================================================
# KPI CARDS
# ============================================================
st.subheader("📊 KPIs (Ejecutivo)")
k1, k2, k3, k4 = st.columns(4)
k1.metric("Ingresos (Total to Bill)", f"${income:,.2f}")
k2.metric("Costos (Total Cost Month)", f"${cost:,.2f}", f"{p_cost*100:,.2f}%")
k3.metric("Gross (Ingreso - Costo)", f"${gross:,.2f}", f"{p_gross*100:,.2f}%")
k4.metric("Gastos fijos (Gasto Fijo)", f"${fixed_total:,.2f}", f"{p_fixed*100:,.2f}%")

k5, k6, k7, k8 = st.columns(4)
k5.metric("Neto (Gross - Fijos)", f"${net:,.2f}", f"{p_net*100:,.2f}%")
k6.metric("Total Management Fee", f"${mgmt_fee_total:,.2f}", f"{p_mgmt*100:,.2f}%")
k7.metric("Royalty CNET Group Inc 5%", f"${royalty_total:,.2f}", f"{p_roy*100:,.2f}%")
k8.metric("Nuevo Total", f"${new_total:,.2f}", f"{p_new*100:,.2f}%")

# ============================================================
# WATERFALL
# ============================================================
st.subheader("📉 Cascada Ejecutiva")
fig_waterfall = go.Figure(go.Waterfall(
    orientation="v",
    measure=["absolute", "relative", "relative", "relative", "relative", "total"],
    x=["Ingresos", "Costos", "Gross", "Gastos fijos", "Mgmt+Royalty", "Nuevo Total"],
    y=[income, -cost, gross, -fixed_total, (mgmt_fee_total + royalty_total), new_total],
))
fig_waterfall.update_layout(title="Cascada: Ingresos → Costos → Fijos → +Fees → Nuevo Total", showlegend=False)
st.plotly_chart(fig_waterfall, use_container_width=True)

# ============================================================
# ✅ BREAKDOWN POR MES (FILTRADO) usando Year + Month texto
# ============================================================
st.subheader("🗓️ Breakdown por Mes (filtrado)")

if MONTH_COL in df.columns:
    df_m = df.copy()  # filtrado
    df_m["_MonthLabel"] = month_text_to_month_label(df_m[MONTH_COL], year=selected_year)

    # asegurar numéricos
    for col in [COL_INCOME, COL_COST, COL_MGMT, COL_ROY]:
        c = find_col(df_m, col)
        if c:
            df_m[c] = pd.to_numeric(df_m[c], errors="coerce")

    mi = find_col(df_m, COL_INCOME)
    mc = find_col(df_m, COL_COST)
    mm = find_col(df_m, COL_MGMT)
    mr = find_col(df_m, COL_ROY)

    if df_m["_MonthLabel"].notna().any() and all([mi, mc, mm, mr]):
        g = df_m.groupby("_MonthLabel", dropna=True).agg(
            Income=(mi, "sum"),
            Cost=(mc, "sum"),
            MgmtFee=(mm, "sum"),
            Royalty=(mr, "sum"),
        ).reset_index().rename(columns={"_MonthLabel": "Month"})

        g["Gross"] = g["Income"] - g["Cost"]
        g["Net (Gross - Fixed)"] = g["Gross"] - fixed_total
        g["New Total"] = g["Net (Gross - Fixed)"] + g["MgmtFee"] + g["Royalty"]

        g["Month_dt"] = pd.to_datetime(g["Month"] + "-01", errors="coerce")
        g = g.sort_values("Month_dt").drop(columns=["Month_dt"])

        g_show = g.copy()
        for c in ["Income", "Cost", "Gross", "MgmtFee", "Royalty", "Net (Gross - Fixed)", "New Total"]:
            g_show[c] = g_show[c].map(lambda x: f"${float(x):,.2f}")
        st.dataframe(g_show, use_container_width=True)

        fig_month = go.Figure()
        fig_month.add_trace(go.Bar(name="Income", x=g["Month"], y=g["Income"]))
        fig_month.add_trace(go.Bar(name="Cost", x=g["Month"], y=g["Cost"]))
        fig_month.add_trace(go.Scatter(name="New Total", x=g["Month"], y=g["New Total"], mode="lines+markers"))
        fig_month.update_layout(
            title=f"Mes a Mes (filtrado) - Year {selected_year}: Income vs Cost + New Total",
            barmode="group",
            xaxis_title="Month",
            yaxis_title="Amount",
        )
        st.plotly_chart(fig_month, use_container_width=True)
    else:
        st.info("No pude armar el breakdown por mes: revisa que Month tenga valores válidos (January, Feb, etc.).")
else:
    st.info("No existe la columna Month en el dataframe filtrado.")

# ============================================================
# EXPORT PDF
# ============================================================
st.divider()
st.subheader("📄 Export Executive Report (PDF)")

pdf_bytes = build_pdf_report(
    income=income, cost=cost, gross=gross, fixed_total=fixed_total, net=net,
    mgmt_fee_total=mgmt_fee_total, royalty_total=royalty_total, new_total=new_total,
    p_cost=p_cost, p_gross=p_gross, p_fixed=p_fixed, p_net=p_net, p_mgmt=p_mgmt, p_roy=p_roy, p_new=p_new,
    target=target, gross_margin=gross_margin, net_margin=net_margin, final_margin=final_margin,
    fig_waterfall=fig_waterfall,
    fig_gauge=fig_gauge,
)

st.download_button(
    "⬇️ Descargar PDF Ejecutivo",
    data=pdf_bytes,
    file_name="CNET_Executive_Report.pdf",
    mime="application/pdf",
)

# ============================================================
# TABLES
# ============================================================
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

with st.expander("Detalle Real Master (filtrado)"):
    st.dataframe(sanitize_for_arrow(df), use_container_width=True)

with st.expander("Detalle Real Master (sin filtrar)"):
    st.dataframe(sanitize_for_arrow(df_all), use_container_width=True)
