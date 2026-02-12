import os
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from datetime import datetime

EXCEL_PATH = "Master January 2026.xlsx"
SHEET_REAL = "Real Master"
SHEET_FIXED = "Gasto Fijo"

st.set_page_config(page_title="CNET Costeo Dashboard", layout="wide")
st.title("📊 CNET Costeo & Neto Dashboard")
colA, colB = st.columns([1, 3])

with colA:
    if st.button("🔄 Refresh datos"):
        st.cache_data.clear()
        st.rerun()

with colB:
    auto = st.toggle("Auto refresh (cada 60s)", value=False)
    if auto:
        import time
        time.sleep(60)
        st.cache_data.clear()
        st.rerun()
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

@st.cache_data(ttl=60)
def load_real_master(path: str):
    raw = pd.read_excel(path, sheet_name=SHEET_REAL, header=None)

    # En este archivo, los headers reales están en la fila 7 (index 6)
    header_idx = 6
    headers = make_unique_columns(raw.iloc[header_idx].tolist())

    df = raw.iloc[header_idx + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    # Limpieza de nombres
    df.columns = [str(c).strip() for c in df.columns]

    return df
def build_pdf_report(
    income, cost, gross, fixed_total, net, mgmt_fee_total, royalty_total, new_total,
    p_cost, p_gross, p_fixed, p_net, p_mgmt, p_roy, p_new,
    fig_waterfall=None, fig_gauge=None
):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    width, height = letter

    y = height - 50
    c.setFont("Helvetica-Bold", 16)
    c.drawString(40, y, "CNET Costeo & Neto - Executive Summary")
    y -= 18
    c.setFont("Helvetica", 9)
    c.drawString(40, y, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    y -= 25

    # Tabla KPI
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "KPIs")
    y -= 14

    rows = [
        ("Ingresos", income, 1.0),
        ("Costos", cost, p_cost),
        ("Gross", gross, p_gross),
        ("Gastos fijos", fixed_total, p_fixed),
        ("Neto", net, p_net),
        ("Management Fee", mgmt_fee_total, p_mgmt),
        ("Royalty 5%", royalty_total, p_roy),
        ("Nuevo Total", new_total, p_new),
    ]

    c.setFont("Helvetica", 10)
    for label, val, pct in rows:
        c.drawString(40, y, label)
        c.drawRightString(360, y, f"${val:,.2f}")
        c.drawRightString(520, y, f"{pct:.1%}")
        y -= 14

    y -= 10

    # Imágenes (si kaleido está instalado)
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
            # Si falla, simplemente no mete la imagen
            c.setFont("Helvetica", 9)
            c.drawString(40, y_top, f"{title} (no se pudo exportar la imagen)")
            return y_top - 15

    y = add_plotly_image(fig_gauge, "Gauge - Margen Final", y)
    y = add_plotly_image(fig_waterfall, "Cascada - Ingresos → Costos → Fijos → Fees → Total", y)

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.getvalue()
@st.cache_data(ttl=60)
def load_fixed_expenses_total(path: str) -> float:
    fx = pd.read_excel(path, sheet_name=SHEET_FIXED, header=None)

    # Montos están en la columna 2 (tercera columna) según tu hoja
    amounts = pd.to_numeric(fx.iloc[:, 2], errors="coerce")
    total_fixed = float(amounts.fillna(0).sum())
    return total_fixed

def safe_pct(x, base):
    return (x / base) if base not in (0, None) else 0.0

if not os.path.exists(EXCEL_PATH):
    st.error(f"No encuentro el archivo: {EXCEL_PATH} en el repositorio.")
    st.stop()

df = load_real_master(EXCEL_PATH)
df = add_filters(df)
def add_filters(df):
    st.sidebar.header("Filtros Ejecutivos")

    if "Company" in df.columns:
        sel_company = st.sidebar.multiselect(
            "Company", sorted(df["Company"].dropna().unique())
        )
        if sel_company:
            df = df[df["Company"].isin(sel_company)]

    if "Province" in df.columns:
        sel_province = st.sidebar.multiselect(
            "Province", sorted(df["Province"].dropna().unique())
        )
        if sel_province:
            df = df[df["Province"].isin(sel_province)]

    if "Client" in df.columns:
        sel_client = st.sidebar.multiselect(
            "Client", sorted(df["Client"].dropna().unique())
        )
        if sel_client:
            df = df[df["Client"].isin(sel_client)]

    if "Project" in df.columns:
        sel_project = st.sidebar.multiselect(
            "Project", sorted(df["Project"].dropna().unique())
        )
        if sel_project:
            df = df[df["Project"].isin(sel_project)]

    return df
fixed_total = load_fixed_expenses_total(EXCEL_PATH)

# Columnas (en tu archivo existen así; ojo que una tiene espacio al final)
COL_INCOME = "Total to Bill"
COL_COST = "Total Cost Month"
COL_MGMT = "Total Management Fee"  # si no existe, buscamos variante con espacio
COL_ROYALTY = "Royalty CNET Group Inc 5%"

# Resolver variantes con espacios
def find_col(name):
    if name in df.columns:
        return name
    # buscar por coincidencia ignorando espacios
    n = name.strip().lower()
    for c in df.columns:
        if str(c).strip().lower() == n:
            return c
    # buscar por contains
    for c in df.columns:
        if n in str(c).strip().lower():
            return c
    return None

c_income = find_col(COL_INCOME)
c_cost = find_col(COL_COST)
c_mgmt = find_col(COL_MGMT)
c_roy = find_col(COL_ROYALTY)

missing = [k for k,v in {
    "Total to Bill": c_income,
    "Total Cost Month": c_cost,
    "Total Management Fee": c_mgmt,
    "Royalty CNET Group Inc 5%": c_roy
}.items() if v is None]

if missing:
    st.error(f"No encontré estas columnas en Real Master: {missing}")
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
st.subheader("📌 Margen Ejecutivo (KPIs + Semáforo)")

# Targets
target = st.slider("Margen objetivo (%)", 0, 60, 25) / 100
zona_amarilla = 0.05  # +/- 5% alrededor del target

gross_margin = gross / income if income else 0
net_margin   = net / income if income else 0
final_margin = new_total / income if income else 0  # después de fees

def semaforo(m, tgt):
    if m >= tgt + zona_amarilla:
        return "🟢"
    elif m >= tgt - zona_amarilla:
        return "🟡"
    else:
        return "🔴"

c1, c2, c3 = st.columns(3)
c1.metric("Gross Margin", f"{gross_margin:.1%}", f"{semaforo(gross_margin, target)} vs {target:.0%}")
c2.metric("Net Margin",   f"{net_margin:.1%}",   f"{semaforo(net_margin, target)} vs {target:.0%}")
c3.metric("Final Margin (after fees)", f"{final_margin:.1%}", f"{semaforo(final_margin, target)} vs {target:.0%}")

# ---- Gauge simple (tipo velocímetro) ----
st.caption("Gauge: Margen final (después de fees)")

gauge_max = 0.60  # 60% techo visual
g = max(0, min(final_margin, gauge_max))

import plotly.graph_objects as go
fig_gauge = go.Figure(go.Indicator(
    mode="gauge+number",
    value=float(g*100),
    number={"suffix": "%"},
    gauge={
        "axis": {"range": [0, gauge_max*100]},
        "threshold": {"line": {"width": 4}, "value": float(target*100)},
        "steps": [
            {"range": [0, max(0,(target - zona_amarilla)*100)]},
            {"range": [max(0,(target - zona_amarilla)*100), (target + zona_amarilla)*100]},
            {"range": [(target + zona_amarilla)*100, gauge_max*100]},
        ],
    }
))
st.plotly_chart(fig_gauge, use_container_width=True)
st.subheader("Indicador de Margen Ejecutivo")

target = st.slider("Margen objetivo (%)", 0, 60, 25) / 100
margin = gross / income if income != 0 else 0

if margin >= target:
    st.success(f"✅ Margen OK: {margin:.1%} (Objetivo {target:.0%})")
else:
    st.error(f"⚠️ Margen bajo: {margin:.1%} (Objetivo {target:.0%})")

# Percentajes sobre ingresos
p_cost = safe_pct(cost, income)
p_gross = safe_pct(gross, income)
p_fixed = safe_pct(fixed_total, income)
p_net = safe_pct(net, income)
p_mgmt = safe_pct(mgmt_fee_total, income)
p_roy = safe_pct(royalty_total, income)
p_new = safe_pct(new_total, income)

# KPIs (valores)
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
st.divider()
st.subheader("📄 Export Executive Report")

pdf_bytes = build_pdf_report(
    income=income, 
    cost=cost, 
    gross=gross, 
    fixed_total=fixed_total, 
    net=net,
    mgmt_fee_total=mgmt_fee_total, 
    royalty_total=royalty_total, 
    new_total=new_total,
    p_cost=p_cost, 
    p_gross=p_gross, 
    p_fixed=p_fixed, 
    p_net=p_net, 
    p_mgmt=p_mgmt, 
    p_roy=p_roy, 
    p_new=p_new,
    fig_waterfall=fig,        # ⚠️ usa aquí el nombre real de tu figura
    fig_gauge=fig_gauge       # ⚠️ usa aquí tu gauge
)

st.download_button(
    "⬇️ Descargar PDF Ejecutivo",
    data=pdf_bytes,
    file_name="CNET_Executive_Report.pdf",
    mime="application/pdf"
)
# Waterfall
fig = go.Figure(go.Waterfall(
    orientation="v",
    measure=["absolute", "relative", "relative", "relative", "relative", "total"],
    x=["Ingresos", "Costos", "Gross", "Gastos fijos", "Mgmt+Royalty", "Nuevo Total"],
    y=[
        income,
        -cost,
        gross,            # mostramos el nivel gross como paso informativo
        -fixed_total,
        (mgmt_fee_total + royalty_total),
        new_total
    ],
))
fig.update_layout(title="Cascada: Ingresos → Costos → Fijos → +Fees → Nuevo Total", showlegend=False)
st.plotly_chart(fig, use_container_width=True)

# Tabla resumen (valores + %)
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

st.subheader("Resumen")
st.dataframe(summary, use_container_width=True)


with st.expander("Detalle Real Master"):
    st.dataframe(df, use_container_width=True)
