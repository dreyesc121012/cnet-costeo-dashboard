import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import os

EXCEL_PATH = "REAL MASTER.xlsx"
SHEET_NAME = "Real Master"

ADMIN_FIXED = 4500
FEE_RATE = 0.05

st.set_page_config(page_title="CNET Costeo Dashboard", layout="wide")
st.title("📊 CNET Costeo & Neto Dashboard")

@st.cache_data(ttl=60)
def load_data():
    raw = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)

    # Limpiar nombres de columnas
    raw.columns = raw.columns.str.strip()

    return raw

if not os.path.exists(EXCEL_PATH):
    st.error("No encuentro el archivo REAL MASTER.xlsx")
    st.stop()

df = load_data()

st.write("Columnas detectadas:", df.columns.tolist())

# Buscar columnas automáticamente
income_col = [c for c in df.columns if "bill" in c.lower()]
cost_col = [c for c in df.columns if "cost" in c.lower()]

if not income_col or not cost_col:
    st.error("No se encontraron columnas de ingreso o costo.")
    st.stop()

income_col = income_col[0]
cost_col = cost_col[0]

df[income_col] = pd.to_numeric(df[income_col], errors="coerce")
df[cost_col] = pd.to_numeric(df[cost_col], errors="coerce")

income = df[income_col].sum()
cost = df[cost_col].sum()

gross = income - cost
after_admin = gross - ADMIN_FIXED
fee = after_admin * FEE_RATE if after_admin > 0 else 0
net_final = after_admin - fee

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Ingresos", f"${income:,.2f}")
k2.metric("Costos", f"${cost:,.2f}")
k3.metric("Utilidad Bruta", f"${gross:,.2f}")
k4.metric("Después Admin ($4,500)", f"${after_admin:,.2f}")
k5.metric("Neto Final (menos 5%)", f"${net_final:,.2f}")

fig = go.Figure(go.Waterfall(
    measure=["relative", "relative", "relative", "relative", "total"],
    x=["Ingresos", "Costos", "Admin", "Fee 5%", "Neto Final"],
    y=[income, -cost, -ADMIN_FIXED, -fee, net_final],
))

fig.update_layout(title="Cascada Financiera", showlegend=False)

st.plotly_chart(fig, use_container_width=True)
