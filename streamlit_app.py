import os
import time
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st
import plotly.graph_objects as go

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader


# =========================
# CONFIG
# =========================
EXCEL_PATH = "Master January 2026.xlsx"
SHEET_REAL = "Real Master"
SHEET_FIXED = "Gasto Fijo"

st.set_page_config(page_title="CNET Costeo Dashboard", layout="wide")
st.title("📊 CNET Costeo & Neto Dashboard")


# =========================
# TOP BAR (Refresh)
# =========================
colA, colB = st.columns([1, 3])

with colA:
    if st.button("🔄 Refresh datos"):
        st.cache_data.clear()
        st.rerun()

with colB:
    auto = st.toggle("Auto refresh (cada 60s)", value=False)
    if auto:
        time.sleep(60)
        st.cache_data.clear()
        st.rerun()


# =========================
# HELPERS / LOADERS
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
    """Busca columna por match exacto, por strip/lower, o por contains."""
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


   @st.cache_data(ttl=60)
def load_real_master(path: str) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=SHEET_REAL, header=None)

    header_idx = 6
    headers = make_unique_columns(raw.iloc[header_idx].tolist())

    df = raw.iloc[header_idx + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    # Limpieza nombres
    df.columns = [str(c).strip() for c in df.columns]

    # 🔥 FIX DEFINITIVO PARA ERROR ARROW
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str)

    return df


@st.cache_data(ttl=60)
def load_fixed_expenses_total(path: str) -> float:
    fx = pd.read_excel(path, sheet_name=SHEET_FIXED, header=None)
    # Montos en 3ra columna (index 2)
    amounts = pd.to_numeric(fx.iloc[:, 2], errors="coerce")
    return float(amounts.fillna(0).sum())


def add_filters(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("Filtros Ejecutivos")

    # Estos nombres dependen de tu archivo; si existen, aparecen
    if "Company" in df.columns:
        opts = sorted(df["Company"].dropna().unique())
        sel = st.sidebar.multiselect("Company", opts)
        if sel:
            df = df[df["Company"].isin(sel)]

    if "Province" in df.columns:
        opts = sorted(df["Province"].dropna().unique())
        sel = st.sidebar.multiselect("Province", opts)
        if sel:
            df = df[df["Province"].isin(sel)]

    if "Client" in df.columns:
        opts = sorted(df["Client"].dropna().unique())
        sel = st.sidebar.multiselect("Client", opts)
        if sel:
            df = df[df["Client"].isin(sel)]

    # A veces se llama Project o Proyecto, soportamos ambos
    if "Project" in df.columns:
        opts = sorted(df["Project"].dropna().unique())
        sel = st.sidebar.multiselect("Project", opts)
        if sel:
            df = df[df["Project"].isin(sel)]
    elif "Proyecto" in df.columns:
        opts = sorted(df["Proyecto"].dropna().unique())
        sel = st.sidebar.multiselect("Proyecto", opts)
        if sel:
            df = df[df["Proyecto"].isin(sel)]

    return df


def build_pdf_report(
    income, cost, gross, fixed_total, net, mgmt_fee_total, royalty_total, new_total,
    p_cost, p_gross, p_fixed, p_net, p_mgmt, p_roy, p_new,
    fig_waterfall=None, fig_gauge=None,
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

    # Inserta imágenes de Plotly (requiere kaleido)
    def add_plotly_image(fig, title, y_top):
        if fig is None:
            return y_top
        try:
            img_bytes = fig.to_image(format="png")
            img = ImageReader(BytesIO(img_bytes))
            c.setFont("Helvetica-Bold", 11)
            c.drawString(40, y_top, title)
            y_top -= 10
            c.drawImage(img, 40, y_top - 220, width=520, height=220, preserveAspectRatio=True, mask="auto")
            return y_top - 235
        except Exception:
            c.setFont("Helvetica", 9)
            c.drawString(40, y_top, f"{title} (no se pudo exportar imagen; instala 'kaleido')")
            return y_top - 15

    y = add_plotly_image(fig_gauge, "Gauge - Margen Final", y)
    y = add_plotly_image(fig_waterfall, "Cascada - Ingresos → Costos → Fijos → Fees → Total", y)

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.getvalue()


# =========================
# LOAD DATA
# =========================
if not os.path.exists(EXCEL_PATH):
    st.error(f"No encuentro el archivo: {EXCEL_PATH} (debe estar en la misma carpeta del proyecto).")
    st.stop()

df = load_real_master(EXCEL_PATH)
df = add_filters(df)

fixed_total = load_fixed_expenses_total(EXCEL_PATH)


# =========================
# REQUIRED COLUMNS
# =========================
COL_INCOME = "Total to Bill"
COL_COST = "Total Cost Month"
COL_MGMT = "Total Management Fee"
COL_ROYALTY = "Royalty CNET Group Inc 5%"

c_income = find_col(df, COL_INCOME)
c_cost = find_col(df, COL_COST)
c_mgmt = find_col(df, COL_MGMT)
c_roy = find_col(df, COL_ROYALTY)

missing = [k for k, v in {
    COL_INCOME: c_income,
    COL_COST: c_cost,
    COL_MGMT: c_mgmt,
    COL_ROYALTY: c_roy
}.items() if v is None]

if missing:
    st.error(f"No encontré estas columnas en '{SHEET_REAL}': {missing}")
    with st.expander("Ver columnas detectadas"):
        st.write(df.columns.tolist())
    st.stop()

# Numeric
df[c_income] = pd.to_numeric(df[c_income], errors="coerce")
df[c_cost] = pd.to_numeric(df[c_cost], errors="coerce")
df[c_mgmt] = pd.to_numeric(df[c_mgmt], errors="coerce")
df[c_roy] = pd.to_numeric(df[c_roy], errors="coerce")


# =========================
# CALCULATIONS
# =========================
income = float(df[c_income].fillna(0).sum())
cost = float(df[c_cost].fillna(0).sum())
gross = income - cost

mgmt_fee_total = float(df[c_mgmt].fillna(0).sum())
royalty_total = float(df[c_roy].fillna(0).sum())

net = gross - fixed_total
new_total = net + mgmt_fee_total + royalty_total

# Percentajes sobre ingresos
p_cost = safe_pct(cost, income)
p_gross = safe_pct(gross, income)
p_fixed = safe_pct(fixed_total, income)
p_net = safe_pct(net, income)
p_mgmt = safe_pct(mgmt_fee_total, income)
p_roy = safe_pct(royalty_total, income)
p_new = safe_pct(new_total, income)


# =========================
# MARGEN EJECUTIVO (PRO)
# =========================
st.subheader("📌 Margen Ejecutivo (KPIs + Semáforo)")

target = st.slider("Margen objetivo (%)", 0, 60, 25) / 100
zona_amarilla = 0.05  # ±5%

gross_margin = gross / income if income else 0
net_margin = net / income if income else 0
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
c2.metric("Net Margin", f"{net_margin:.1%}", f"{semaforo(net_margin, target)} vs {target:.0%}")
c3.metric("Final Margin (after fees)", f"{final_margin:.1%}", f"{semaforo(final_margin, target)} vs {target:.0%}")

st.caption("Gauge: Margen final (después de fees)")

gauge_max = 0.60
g = max(0, min(final_margin, gauge_max))

fig_gauge = go.Figure(go.Indicator(
    mode="gauge+number",
    value=float(g * 100),
    number={"suffix": "%"},
    gauge={
        "axis": {"range": [0, gauge_max * 100]},
        "threshold": {"line": {"width": 4}, "value": float(target * 100)},
        "steps": [
            {"range": [0, max(0, (target - zona_amarilla) * 100)]},
            {"range": [max(0, (target - zona_amarilla) * 100), (target + zona_amarilla) * 100]},
            {"range": [(target + zona_amarilla) * 100, gauge_max * 100]},
        ],
    }
))
st.plotly_chart(fig_gauge, use_container_width=True)


# =========================
# KPIs (VALORES + %)
# =========================
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
st.subheader("Cascada: Ingresos → Costos → Fijos → +Fees → Nuevo Total")

fig = go.Figure(go.Waterfall(
    orientation="v",
    measure=["absolute", "relative", "relative", "relative", "relative", "total"],
    x=["Ingresos", "Costos", "Gross", "Gastos fijos", "Mgmt+Royalty", "Nuevo Total"],
    y=[
        income,
        -cost,
        gross,
        -fixed_total,
        (mgmt_fee_total + royalty_total),
        new_total,
    ],
))
fig.update_layout(showlegend=False)
st.plotly_chart(fig, use_container_width=True)


# =========================
# EXPORT PDF (PRO)
# =========================
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
# SUMMARY TABLE
# =========================
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
    st.dataframe(df.astype(str), use_container_width=True)
