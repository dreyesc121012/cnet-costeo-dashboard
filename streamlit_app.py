from io import BytesIO
from datetime import datetime, timezone
import time

import pandas as pd
import streamlit as st
import plotly.graph_objects as go

import requests
import msal

# PDF
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader


# =========================
# CONFIG
# =========================
SHEET_REAL = "Real Master"
SHEET_FIXED = "Gasto Fijo"
HEADER_IDX = 6  # headers reales están en fila 7 (index 6)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Permisos delegados (ya los tienes en Entra)
SCOPES = ["Files.Read.All", "User.Read"]

st.set_page_config(page_title="CNET Costeo Dashboard", layout="wide")


# =========================
# AUTH (Delegated - Device Code)
# =========================
def _authority() -> str:
    tenant_id = st.secrets["TENANT_ID"]
    return f"https://login.microsoftonline.com/{tenant_id}"


def _public_app() -> msal.PublicClientApplication:
    return msal.PublicClientApplication(
        client_id=st.secrets["CLIENT_ID"],
        authority=_authority(),
    )


def is_token_valid() -> bool:
    tok = st.session_state.get("graph_access_token")
    exp = st.session_state.get("graph_token_expires_at", 0)
    # margen de 60s para evitar token “casi expirado”
    return bool(tok) and time.time() < (exp - 60)


def login_device_flow():
    app = _public_app()
    flow = app.initiate_device_flow(scopes=SCOPES)

    if "user_code" not in flow:
        st.error("No se pudo iniciar el Device Flow. Revisa que 'Permitir flujos de clientes públicos' esté habilitado.")
        st.stop()

    st.info(flow["message"])
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        st.error("Login falló. Detalle:")
        st.json(result)
        st.stop()

    st.session_state["graph_access_token"] = result["access_token"]
    # expires_in viene en segundos
    st.session_state["graph_token_expires_at"] = time.time() + int(result.get("expires_in", 3599))


def logout():
    st.session_state.pop("graph_access_token", None)
    st.session_state.pop("graph_token_expires_at", None)
    # también limpia cache para forzar relectura
    st.cache_data.clear()
    st.rerun()


def get_access_token() -> str:
    if not is_token_valid():
        login_device_flow()
    return st.session_state["graph_access_token"]


# =========================
# HELPERS
# =========================
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


def safe_pct(x, base):
    return (x / base) if base not in (0, None) else 0.0


def find_col(df, name):
    """Encuentra columna por match exacto / ignorando espacios / contains."""
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


def sanitize_for_arrow(df: pd.DataFrame) -> pd.DataFrame:
    """
    Evita errores de pyarrow/streamlit al mostrar columnas object con mezcla int/str.
    """
    df2 = df.copy()
    for col in df2.columns:
        if df2[col].dtype == "object":
            df2[col] = df2[col].astype(str)
    return df2


# =========================
# ONEDRIVE DOWNLOAD
# =========================
@st.cache_data(ttl=60)
def download_excel_bytes_from_onedrive() -> bytes:
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    file_path = st.secrets["ONEDRIVE_FILE_PATH"]  # ej: CNET/Master January 2026.xlsx
    url = f"{GRAPH_BASE}/me/drive/root:/{file_path}:/content"

    r = requests.get(url, headers=headers, timeout=60)
    if r.status_code != 200:
        st.error("No pude descargar el Excel desde OneDrive.")
        st.write("Revisa que la ruta sea correcta y que tu usuario tenga permisos sobre el archivo.")
        st.write(f"Ruta usada: {file_path}")
        st.code(r.text)
        st.stop()
    return r.content


# =========================
# LOADERS (cache)
# =========================
@st.cache_data(ttl=60)
def load_real_master_from_bytes(xlsx_bytes: bytes) -> pd.DataFrame:
    raw = pd.read_excel(BytesIO(xlsx_bytes), sheet_name=SHEET_REAL, header=None)
    headers = make_unique_columns(raw.iloc[HEADER_IDX].tolist())

    df = raw.iloc[HEADER_IDX + 1 :].copy()
    df.columns = headers
    df = df.reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    return df


@st.cache_data(ttl=60)
def load_fixed_expenses_total_from_bytes(xlsx_bytes: bytes) -> float:
    fx = pd.read_excel(BytesIO(xlsx_bytes), sheet_name=SHEET_FIXED, header=None)
    # Montos en columna 3 (index 2)
    amounts = pd.to_numeric(fx.iloc[:, 2], errors="coerce")
    return float(amounts.fillna(0).sum())


# =========================
# FILTERS
# =========================
def add_filters(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("Filtros Ejecutivos")

    if "Company" in df.columns:
        sel = st.sidebar.multiselect("Company", sorted(df["Company"].dropna().unique()))
        if sel:
            df = df[df["Company"].isin(sel)]

    if "Province" in df.columns:
        sel = st.sidebar.multiselect("Province", sorted(df["Province"].dropna().unique()))
        if sel:
            df = df[df["Province"].isin(sel)]

    if "Client" in df.columns:
        sel = st.sidebar.multiselect("Client", sorted(df["Client"].dropna().unique()))
        if sel:
            df = df[df["Client"].isin(sel)]

    if "Project Name" in df.columns:
        sel = st.sidebar.multiselect("Project (Project Name)", sorted(df["Project Name"].dropna().unique()))
        if sel:
            df = df[df["Project Name"].isin(sel)]

    return df


# =========================
# PDF REPORT
# =========================
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

    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "KPIs")
    y -= 14

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

    c.setFont("Helvetica", 10)
    c.drawString(40, y, "Concepto")
    c.drawRightString(360, y, "Monto")
    c.drawRightString(520, y, "% Ingresos")
    y -= 12

    for label, val, pct in rows:
        c.drawString(40, y, label[:45])
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
            c.drawImage(img, 40, y_top - 220, width=520, height=220, preserveAspectRatio=True, mask="auto")
            return y_top - 235
        except Exception:
            c.setFont("Helvetica", 9)
            c.drawString(40, y_top, f"{title} (no se pudo exportar imagen)")
            return y_top - 15

    y = add_plotly_image(fig_gauge, "Gauge - Final Margin", y)
    y = add_plotly_image(fig_waterfall, "Cascada - Ingresos → Costos → Fijos → Fees → Total", y)

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.getvalue()


# =========================
# UI HEADER
# =========================
st.title("📊 CNET Costeo & Neto Dashboard")

top_left, top_right = st.columns([2, 3])
with top_left:
    if is_token_valid():
        st.success("✅ Conectado a OneDrive (token activo)")
        if st.button("🔒 Cerrar sesión OneDrive"):
            logout()
    else:
        st.warning("⚠️ No has iniciado sesión en OneDrive")
        if st.button("🔑 Iniciar sesión OneDrive"):
            login_device_flow()
            st.rerun()

with top_right:
    if st.button("🔄 Refresh datos"):
        st.cache_data.clear()
        st.rerun()

st.divider()

# Si no hay token válido, no seguimos (evita errores)
if not is_token_valid():
    st.info("Haz clic en **Iniciar sesión OneDrive** para cargar el Excel.")
    st.stop()

# =========================
# MAIN LOAD (from OneDrive)
# =========================
xlsx_bytes = download_excel_bytes_from_onedrive()

df = load_real_master_from_bytes(xlsx_bytes)
df = add_filters(df)
fixed_total = load_fixed_expenses_total_from_bytes(xlsx_bytes)

# Columnas clave
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

# Convertir a numérico
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

# % sobre ingresos
p_cost  = safe_pct(cost, income)
p_gross = safe_pct(gross, income)
p_fixed = safe_pct(fixed_total, income)
p_net   = safe_pct(net, income)
p_mgmt  = safe_pct(mgmt_fee_total, income)
p_roy   = safe_pct(royalty_total, income)
p_new   = safe_pct(new_total, income)

# =========================
# EXEC MARGIN (PRO)
# =========================
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

# =========================
# KPI CARDS
# =========================
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

# =========================
# WATERFALL
# =========================
st.subheader("📉 Cascada Ejecutiva")
fig = go.Figure(go.Waterfall(
    orientation="v",
    measure=["absolute", "relative", "relative", "relative", "relative", "total"],
    x=["Ingresos", "Costos", "Gross", "Gastos fijos", "Mgmt+Royalty", "Nuevo Total"],
    y=[income, -cost, gross, -fixed_total, (mgmt_fee_total + royalty_total), new_total],
))
fig.update_layout(title="Cascada: Ingresos → Costos → Fijos → +Fees → Nuevo Total", showlegend=False)
st.plotly_chart(fig, use_container_width=True)

# =========================
# OTTAWA vs QUEBEC
# =========================
if "Province" in df.columns:
    st.subheader("🏙️ Comparación Ottawa vs Quebec")
    prov = df["Province"].astype(str)

    df_on = df[prov.str.contains("ON", case=False, na=False) | prov.str.contains("Ottawa", case=False, na=False)]
    df_qc = df[prov.str.contains("QC", case=False, na=False) | prov.str.contains("Quebec", case=False, na=False)]

    def kpi_from(d):
        inc = float(pd.to_numeric(d[c_income], errors="coerce").fillna(0).sum())
        cst = float(pd.to_numeric(d[c_cost], errors="coerce").fillna(0).sum())
        grs = inc - cst
        mgm = float(pd.to_numeric(d[c_mgmt], errors="coerce").fillna(0).sum())
        roy = float(pd.to_numeric(d[c_roy], errors="coerce").fillna(0).sum())
        nt = grs - fixed_total
        tot = nt + mgm + roy
        return inc, grs, tot

    a1, a2, a3 = kpi_from(df_on) if len(df_on) else (0.0, 0.0, 0.0)
    b1, b2, b3 = kpi_from(df_qc) if len(df_qc) else (0.0, 0.0, 0.0)

    cA, cB = st.columns(2)
    with cA:
        st.markdown("**Ottawa / ON (detectado)**")
        st.metric("Ingresos", f"${a1:,.2f}")
        st.metric("Gross", f"${a2:,.2f}")
        st.metric("Nuevo Total", f"${a3:,.2f}")
    with cB:
        st.markdown("**Quebec / QC (detectado)**")
        st.metric("Ingresos", f"${b1:,.2f}")
        st.metric("Gross", f"${b2:,.2f}")
        st.metric("Nuevo Total", f"${b3:,.2f}")

# =========================
# EXPORT PDF
# =========================
st.divider()
st.subheader("📄 Export Executive Report (PDF)")

pdf_bytes = build_pdf_report(
    income=income, cost=cost, gross=gross, fixed_total=fixed_total, net=net,
    mgmt_fee_total=mgmt_fee_total, royalty_total=royalty_total, new_total=new_total,
    p_cost=p_cost, p_gross=p_gross, p_fixed=p_fixed, p_net=p_net,
    p_mgmt=p_mgmt, p_roy=p_roy, p_new=p_new,
    target=target, gross_margin=gross_margin, net_margin=net_margin, final_margin=final_margin,
    fig_waterfall=fig,
    fig_gauge=fig_gauge,
)

st.download_button(
    "⬇️ Descargar PDF Ejecutivo",
    data=pdf_bytes,
    file_name="CNET_Executive_Report.pdf",
    mime="application/pdf",
)

# =========================
# TABLES
# =========================
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

with st.expander("Detalle Real Master"):
    st.dataframe(sanitize_for_arrow(df), use_container_width=True)
