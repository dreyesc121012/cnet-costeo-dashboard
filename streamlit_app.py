import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import os

# =========================
# CONFIGURACIÓN
# =========================
EXCEL_PATH = "REAL MASTER.xlsx"
SHEET_NAME = "Real Master"

ADMIN_FIXED = 4500
FEE_RATE = 0.05

st.set_page_config(page_title="CNET Costeo Dashboard", layout="wide")

st.title("📊 CNET Costeo & Neto Dashboard")

# =========================
# CARGAR EXCEL
# =========================
@st.cache_data(ttl=60)
def load_data():
    raw = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, header=None)

    header_row = raw.iloc[5].tolist()
    data = raw.iloc[6:].copy()
    data.columns = header_row
    data = data.reset_index(drop=True)

    data["Total to Bill"] = pd.to_numeric(data["Total to Bill"], errors="coerce")
    data["Total Cost Month"] = pd.to_numeric(data["Total Cost Month"], errors="coerce")

    return data

if not os.path.exists(EXCEL_PATH):
    st.error("No encuentro el archivo REAL MASTER.xlsx")
    st.stop()

df = load_data()

# =========================
# FILTROS
# =========================
col1, col2, col3 = st.columns(3)

company = col1.selectbox("Company", ["Todos"] + list(df["Company"].dropna().unique()))
client = col2.selectbox("Client", ["Todos"] + list(df["Client"].dropna().unique()))
project = col3.selectbox("Project Name", ["Todos"] + list(df["Project Name"].dropna().unique()))

filtered = df.copy()

if company != "Todos":
    filtered = filtered[filtered["Company"] == company]

if client != "Todos":
    filtered = filtered[filtered["Client"] == client]

if project != "Todos":
    filtered = filtered[filtered["Project Name"] == project]

# =========================
# CÁLCULOS
# =========================
income = filtered["Total to Bill"].sum()
cost = filtered["Total Cost Month"].sum()

gross = income - cost
after_admin = gross - ADMIN_FIXED
fee = after_admin * FEE_RATE if after_admin > 0 else 0
net_final = after_admin - fee

# =========================
# KPIs
# =========================
k1, k2, k3, k4, k5 = st.columns(5)

k1.metric("Ingresos", f"${income:,.2f}")
k2.metric("Costos", f"${cost:,.2f}")
k3.metric("Utilidad Bruta", f"${gross:,.2f}")
k4.metric("Después Admin ($4,500)", f"${after_admin:,.2f}")
k5.metric("Neto Final (menos 5%)", f"${net_final:,.2f}")

# =========================
# GRÁFICO CASCADA
# =========================
fig = go.Figure(go.Waterfall(
    name="",
    orientation="v",
    measure=["relative", "relative", "relative", "relative", "total"],
    x=["Ingresos", "Costos", "Admin", "Fee 5%", "Neto Final"],
    y=[income, -cost, -ADMIN_FIXED, -fee, net_final],
))

fig.update_layout(title="Cascada Financiera", showlegend=False)

st.plotly_chart(fig, use_container_width=True)

# =========================
# TABLA DETALLE
# =========================
with st.expander("Ver detalle"):
    st.dataframe(filtered)
