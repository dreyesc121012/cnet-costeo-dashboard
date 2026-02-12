import os
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

EXCEL_PATH = "REAL MASTER.xlsx"
SHEET_NAME = "Real Master"

ADMIN_FIXED = 4500
FEE_RATE = 0.05

st.set_page_config(page_title="CNET Costeo Dashboard", layout="wide")
st.title("📊 CNET Costeo & Neto Dashboard")

def _normalize(s):
    return str(s).strip().lower()

def find_header_row(raw: pd.DataFrame, max_scan_rows: int = 60) -> int:
    """
    Busca la fila que contiene headers reales. Detecta por palabras clave:
    bill + cost + total (o variantes).
    """
    scan_rows = min(max_scan_rows, len(raw))
    best_row = None
    best_score = -1

    for i in range(scan_rows):
        row_vals = raw.iloc[i].tolist()
        joined = " | ".join(_normalize(x) for x in row_vals if pd.notna(x))

        score = 0
        if "bill" in joined: score += 3
        if "cost" in joined: score += 3
        if "total" in joined: score += 2
        if "project" in joined: score += 1
        if "client" in joined: score += 1
        if "company" in joined: score += 1

        if score > best_score:
            best_score = score
            best_row = i

    return best_row if best_row is not None else 0

@st.cache_data(ttl=60)
def load_data():
    # Leer sin headers (para evitar Unnamed)
    raw = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, header=None)

    header_idx = find_header_row(raw)
    headers = raw.iloc[header_idx].tolist()

    # Limpiar headers: convertir NaN a "Unnamed", strip espacios
    clean_headers = []
    seen = {}
    for h in headers:
        name = "Unnamed" if pd.isna(h) else str(h).strip()
        if name == "":
            name = "Unnamed"
        # Evitar duplicados
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        clean_headers.append(name)

    df = raw.iloc[header_idx + 1:].copy()
    df.columns = clean_headers
    df = df.reset_index(drop=True)

    # Limpieza de columnas (por si vienen con espacios)
    df.columns = [c.strip() for c in df.columns]

    return df, header_idx

if not os.path.exists(EXCEL_PATH):
    st.error("No encuentro el archivo REAL MASTER.xlsx en el repositorio.")
    st.stop()

df, header_idx = load_data()

st.caption(f"✅ Header detectado en fila (Excel): {header_idx + 1}")

# Mostrar columnas detectadas (solo para debug)
with st.expander("Ver columnas detectadas"):
    st.write(df.columns.tolist())

# Encontrar columnas de ingreso/costo (variantes)
income_candidates = [c for c in df.columns if ("bill" in c.lower() and "total" in c.lower()) or c.lower() == "total to bill"]
cost_candidates   = [c for c in df.columns if ("cost" in c.lower() and "total" in c.lower()) or c.lower() == "total cost month"]

if not income_candidates or not cost_candidates:
    st.error("No encontré las columnas de Ingresos/Costos. Revisa el expander de columnas detectadas.")
    st.stop()

income_col = income_candidates[0]
cost_col = cost_candidates[0]

df[income_col] = pd.to_numeric(df[income_col], errors="coerce")
df[cost_col] = pd.to_numeric(df[cost_col], errors="coerce")

income = float(df[income_col].fillna(0).sum())
cost = float(df[cost_col].fillna(0).sum())

gross = income - cost
after_admin = gross - ADMIN_FIXED
fee = (after_admin * FEE_RATE) if after_admin > 0 else 0
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

with st.expander("Ver detalle"):
    st.dataframe(df, use_container_width=True)
