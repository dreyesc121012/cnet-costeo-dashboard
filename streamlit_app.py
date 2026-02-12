import os
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

EXCEL_PATH = "Master January 2026.xlsx"
SHEET_REAL = "Real Master"
SHEET_FIXED = "Gasto Fijo"

st.set_page_config(page_title="CNET Costeo Dashboard", layout="wide")
st.title("📊 CNET Costeo & Neto Dashboard")

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
