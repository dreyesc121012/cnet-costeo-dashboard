import io
import re
import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

st.set_page_config(page_title="CNET - Invoice Reports", layout="wide")

# =========================================================
# Helpers
# =========================================================
def safe_num(x) -> float:
    try:
        if pd.isna(x):
            return 0.0
        s = str(x).replace("$", "").replace(",", "").strip()
        return float(s) if s else 0.0
    except Exception:
        return 0.0

def contains_ci(text, needle: str) -> bool:
    return needle.lower() in str(text or "").lower()

def read_any_table(file_bytes: bytes) -> pd.DataFrame:
    # Auto-detect excel vs csv
    try:
        return pd.read_excel(io.BytesIO(file_bytes))
    except Exception:
        try:
            return pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8", engine="python")
        except Exception:
            return pd.read_csv(io.BytesIO(file_bytes), encoding="latin1", engine="python")

# =========================================================
# Login + Download (CSRF)
# =========================================================
def download_export_file() -> bytes:
    session = requests.Session()

    # 1) GET login page to get CSRF
    r1 = session.get("https://app.master.cnetfranchise.com/login", timeout=60)
    r1.raise_for_status()

    soup = BeautifulSoup(r1.text, "lxml")
    token_el = soup.select_one('input[name="_csrf_token"]')
    if not token_el or not token_el.get("value"):
        raise Exception("No pude encontrar _csrf_token en /login")
    csrf = token_el["value"]

    # 2) POST login_check
    payload = {
        "_csrf_token": csrf,
        "_username": st.secrets["SYSTEM_USERNAME"],
        "_password": st.secrets["SYSTEM_PASSWORD"],
        "_submit": "Login",
    }
    r2 = session.post(
        "https://app.master.cnetfranchise.com/login_check",
        data=payload,
        allow_redirects=True,
        timeout=60
    )
    r2.raise_for_status()
    if "login" in r2.url.lower():
        raise Exception("Login falló. Revisa usuario/clave.")

    # 3) GET export (tu export actual devuelve CSV)
    export_url = st.secrets["EXPORT_EXCEL_URL"]
    r3 = session.get(export_url, timeout=60)
    r3.raise_for_status()
    return r3.content

# =========================================================
# Province inference (from tax columns like "QST QC")
# =========================================================
def add_province_from_taxes(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    tax_cols = [c for c in df.columns if any(k in c.upper() for k in ["GST","HST","QST","PST","RST","TAX"])]

    def province_from_taxes(row):
        hits = []
        for c in tax_cols:
            if safe_num(row.get(c, 0)) > 0:
                hits.append(c.upper())

        if not hits:
            return "Unknown"

        # Specific rule: QST QC => Quebec
        for c in hits:
            if "QST" in c and "QC" in c:
                return "Quebec"

        mapping = {
            "QC": "Quebec", "ON": "Ontario", "BC": "British Columbia", "AB": "Alberta",
            "MB": "Manitoba", "SK": "Saskatchewan", "NB": "New Brunswick", "NS": "Nova Scotia",
            "NL": "Newfoundland and Labrador", "PE": "Prince Edward Island",
            "NT": "Northwest Territories", "NU": "Nunavut", "YT": "Yukon"
        }

        for code, prov in mapping.items():
            for c in hits:
                if re.search(rf"\b{code}\b", c):
                    return prov

        # fallback
        for c in hits:
            if "QST" in c:
                return "Quebec"
            if "HST" in c:
                return "HST Province (Unknown)"
            if "GST" in c:
                return "GST Only (Unknown)"

        return "Unknown"

    df["Province"] = df.apply(province_from_taxes, axis=1)
    return df

# =========================================================
# Build columns (your rules)
# =========================================================
def build_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Required columns (adjust if your export uses different names)
    col_work = "Work Description"
    col_vendor = "Vendor Company Name"
    col_buyer = "Buyer Company Name"
    col_total = "Total Amount Without Taxes"

    missing = [c for c in [col_work, col_vendor, col_buyer, col_total] if c not in df.columns]
    if missing:
        raise Exception(f"Faltan columnas requeridas en el archivo: {missing}")

    df[col_total] = df[col_total].apply(safe_num)

    # Service
    df["Service"] = df[col_work].apply(lambda x: "Regular" if contains_ci(x, "janitorial") else "One Shot")

    # Service and Name
    df["Service and Name"] = df["Service"].astype(str) + " " + df[col_buyer].astype(str)

    # Brokerage
    df["Brokerage"] = df[col_buyer].apply(lambda x: "Brokerage5" if contains_ci(x, "5BF") else "Without Brokerage")

    # 3% Royalty Fee Group
    df["3% Royalty Fee Group"] = df.apply(
        lambda r: r[col_total] * 0.03 if contains_ci(r["Service and Name"], "BGIS SCS regular") else 0.0,
        axis=1
    )

    # 3% Royalty Fee Master
    df["3% Royalty Fee Master"] = df.apply(
        lambda r: r[col_total] * 0.03 if contains_ci(r["Service and Name"], "BGIS SCS regular") else 0.0,
        axis=1
    )

    # 5% Royalty Fee Group
    df["5% Royalty Fee Group"] = df.apply(
        lambda r: 0.0 if contains_ci(r["Service and Name"], "BGIS SCS regular") else r[col_total] * 0.05,
        axis=1
    )

    # 5% Royalty Fee Master2
    df["5% Royalty Fee Master2"] = df.apply(
        lambda r: 0.0 if contains_ci(r["Service and Name"], "BGIS SCS regular") else r[col_total] * 0.05,
        axis=1
    )

    # 1% Marketing Fee
    df["1% Marketing Fee"] = df[col_total] * 0.01

    # Brokerages rate (numeric)
    df["Brokerages"] = df["Brokerage"].apply(lambda x: 0.05 if x == "Brokerage5" else 0.0)

    # 5% Brokerage Fee
    df["5% Brokerage Fee"] = df.apply(
        lambda r: r[col_total] * r["Brokerages"] if r["Brokerages"] == 0.05 else 0.0,
        axis=1
    )

    # 2.5% Brokerage Fee (placeholder)
    df["2.5% Brokerage Fee"] = 0.0

    # Province
    df = add_province_from_taxes(df)

    return df

# =========================================================
# Reports
# =========================================================
def report_resume_without_fees(df: pd.DataFrame) -> pd.DataFrame:
    # Pivot like your Excel:
    # Rows: Vendor -> Service -> Buyer
    # Columns: Province
    # Values: Sum Total Amount Without Taxes
    pivot = pd.pivot_table(
        df,
        index=["Vendor Company Name", "Service", "Buyer Company Name"],
        columns=["Province"],
        values="Total Amount Without Taxes",
        aggfunc="sum",
        fill_value=0
    ).reset_index()
    return pivot

def report_validation(df: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "Total Amount Without Taxes",
        "3% Royalty Fee Group",
        "5% Royalty Fee Group",
        "1% Marketing Fee",
        "5% Brokerage Fee",
        "2.5% Brokerage Fee",
    ]
    cols = [c for c in cols if c in df.columns]

    rep = (
        df.groupby(["Province", "Vendor Company Name", "Brokerage", "Buyer Company Name"], dropna=False)[cols]
          .sum()
          .reset_index()
    )
    return rep

def report_breakdown(df: pd.DataFrame) -> pd.DataFrame:
    # "Company" column: use Company Name if exists, else fallback
    company_col = "Company Name" if "Company Name" in df.columns else ("Company" if "Company" in df.columns else "Vendor Company Name")

    group_cols = [
        company_col,
        "Service",
        "Province",
        "Buyer Company Name",
        "Brokerage",
        "Vendor Company Name",
    ]

    value_cols = [
        "Total Amount Without Taxes",
        "3% Royalty Fee Group",
        "5% Royalty Fee Group",
        "1% Marketing Fee",
        "5% Brokerage Fee",
        "2.5% Brokerage Fee",
    ]
    value_cols = [c for c in value_cols if c in df.columns]

    rep = (
        df.groupby(group_cols, dropna=False)[value_cols]
          .sum()
          .reset_index()
    ).rename(columns={company_col: "Company"})

    if "Total Amount Without Taxes" in rep.columns:
        rep = rep.sort_values("Total Amount Without Taxes", ascending=False)

    return rep

def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")

# =========================================================
# UI
# =========================================================
st.title("CNET - Invoice Reports")

with st.sidebar:
    st.subheader("Select report")

    report_choice = st.radio(
        "Report type",
        [
            "Resume Without Fees",
            "Validation",
            "Jeff-validation",
            "Breakdown"
        ],
        index=0
    )

    st.divider()
    st.subheader("Month filter")

    months = [
        ("January", 1), ("February", 2), ("March", 3), ("April", 4),
        ("May", 5), ("June", 6), ("July", 7), ("August", 8),
        ("September", 9), ("October", 10), ("November", 11), ("December", 12),
    ]
    month_name = st.selectbox("Month", [m[0] for m in months], index=0)
    month_num = dict(months)[month_name]
    year = st.number_input("Year", min_value=2020, max_value=2035, value=2026, step=1)

    run = st.button("Download + Process")

if not run:
    st.stop()

with st.spinner("Downloading export from CNET..."):
    content = download_export_file()

df_raw = read_any_table(content)

# Parse date and filter by month/year
if "Creation Date" not in df_raw.columns:
    st.error("No encuentro la columna 'Creation Date' para filtrar por mes.")
    st.stop()

df_raw["Creation Date"] = pd.to_datetime(df_raw["Creation Date"], errors="coerce")
df_month = df_raw[(df_raw["Creation Date"].dt.month == month_num) & (df_raw["Creation Date"].dt.year == int(year))].copy()

st.success(f"Archivo descargado. Filas totales: {len(df_raw):,} | Filas del mes: {len(df_month):,}")

# Apply business rules
df_final = build_columns(df_month)

# Report output
if report_choice == "Resume Without Fees":
    st.subheader("Resume Without Fees")
    rep = report_resume_without_fees(df_final)

elif report_choice == "Validation":
    st.subheader("Validation")
    rep = report_validation(df_final)

elif report_choice == "Jeff-validation":
    st.subheader("Jeff-validation")
    # Same structure as Validation (if Jeff has special rule, we add it)
    rep = report_validation(df_final)

else:
    st.subheader("Breakdown (Company / Service / Province / Buyer / Brokerage / Vendor)")
    rep = report_breakdown(df_final)

st.dataframe(rep, use_container_width=True)

# Download button for selected report
st.download_button(
    label="Download selected report (CSV)",
    data=df_to_csv_bytes(rep),
    file_name=f"{report_choice.replace(' ', '_').lower()}_{month_name.lower()}_{int(year)}.csv",
    mime="text/csv"
)

# Preview calculated columns so you can verify Province + rules
st.divider()
st.subheader("Preview - Calculated Columns (incluye Province)")

cols_show = [
    "Creation Date",
    "Work Description",
    "Vendor Company Name",
    "Buyer Company Name",
    "Service",
    "Service and Name",
    "Brokerage",
    "Province",
    "Total Amount Without Taxes",
    "3% Royalty Fee Group",
    "3% Royalty Fee Master",
    "5% Royalty Fee Group",
    "5% Royalty Fee Master2",
    "1% Marketing Fee",
    "Brokerages",
    "5% Brokerage Fee",
    "2.5% Brokerage Fee",
]
cols_show = [c for c in cols_show if c in df_final.columns]
st.dataframe(df_final[cols_show].head(200), use_container_width=True)
