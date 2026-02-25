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
        if s.startswith("(") and s.endswith(")"):
            s = "-" + s[1:-1]
        return float(s) if s else 0.0
    except Exception:
        return 0.0

def contains_ci(text, needle: str) -> bool:
    return needle.lower() in str(text or "").lower()

def read_any_table(file_bytes: bytes) -> pd.DataFrame:
    try:
        return pd.read_excel(io.BytesIO(file_bytes))
    except Exception:
        try:
            return pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8", engine="python")
        except Exception:
            return pd.read_csv(io.BytesIO(file_bytes), encoding="latin1", engine="python")

def fmt_money(x: float) -> str:
    try:
        return f"${x:,.2f}"
    except Exception:
        return "$0.00"

def format_report_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """
    Formatea SOLO para visualización:
    - columnas numéricas a 2 decimales y separador de miles: 1,234.56
    """
    out = df.copy()
    num_cols = out.select_dtypes(include="number").columns.tolist()
    for c in num_cols:
        out[c] = out[c].map(lambda v: f"{v:,.2f}")
    return out

# =========================================================
# Login + Download (CSRF)
# =========================================================
def download_export_file() -> bytes:
    session = requests.Session()

    r1 = session.get("https://app.master.cnetfranchise.com/login", timeout=60)
    r1.raise_for_status()

    soup = BeautifulSoup(r1.text, "lxml")
    token_el = soup.select_one('input[name="_csrf_token"]')
    if not token_el or not token_el.get("value"):
        raise Exception("No pude encontrar _csrf_token en /login")
    csrf = token_el["value"]

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

    export_url = st.secrets["EXPORT_EXCEL_URL"]
    r3 = session.get(export_url, timeout=60)
    r3.raise_for_status()
    return r3.content

# =========================================================
# Province inference
# =========================================================
def add_province_from_taxes(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    tax_cols = [c for c in df.columns if any(k in c.upper() for k in ["GST", "HST", "QST", "PST", "RST", "TAX"])]

    def province_from_taxes(row):
        hits = []
        for c in tax_cols:
            if abs(safe_num(row.get(c, 0))) > 0:
                hits.append(c.upper())

        if not hits:
            txt = f"{row.get('Buyer Company Name','')} {row.get('Vendor Company Name','')} {row.get('Work Description','')}".upper()

            if "QUEBEC" in txt or " QC " in txt or txt.endswith(" QC"):
                return "Quebec"
            if "ONTARIO" in txt or " ON " in txt or txt.endswith(" ON"):
                return "Ontario"
            if "NOVA SCOTIA" in txt or " NS " in txt or txt.endswith(" NS"):
                return "Nova Scotia"
            if "NEW BRUNSWICK" in txt or " NB " in txt or txt.endswith(" NB"):
                return "New Brunswick"
            if "PRINCE EDWARD" in txt or " PEI " in txt or txt.endswith(" PEI"):
                return "Prince Edward Island"
            if "BRITISH COLUMBIA" in txt or " BC " in txt or txt.endswith(" BC"):
                return "British Columbia"
            if "ALBERTA" in txt or " AB " in txt or txt.endswith(" AB"):
                return "Alberta"
            if "MANITOBA" in txt or " MB " in txt or txt.endswith(" MB"):
                return "Manitoba"
            if "SASKATCHEWAN" in txt or " SK " in txt or txt.endswith(" SK"):
                return "Saskatchewan"
            if "NEWFOUNDLAND" in txt or " NL " in txt or txt.endswith(" NL"):
                return "Newfoundland and Labrador"
            return "Unknown"

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
# Build columns
# =========================================================
def build_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    col_work = "Work Description"
    col_vendor = "Vendor Company Name"
    col_buyer = "Buyer Company Name"
    col_total = "Total Amount Without Taxes"

    missing = [c for c in [col_work, col_vendor, col_buyer, col_total] if c not in df.columns]
    if missing:
        raise Exception(f"Faltan columnas requeridas en el archivo: {missing}")

    df[col_total] = df[col_total].apply(safe_num)

    df["Service"] = df[col_work].apply(lambda x: "Regular" if contains_ci(x, "janitorial") else "One Shot")
    df["Service and Name"] = df["Service"].astype(str) + " " + df[col_buyer].astype(str)
    df["Brokerage"] = df[col_buyer].apply(lambda x: "Brokerage5" if contains_ci(x, "5BF") else "Without Brokerage")

    df["3% Royalty Fee Group"] = df.apply(
        lambda r: r[col_total] * 0.03 if contains_ci(r["Service and Name"], "BGIS SCS regular") else 0.0,
        axis=1
    )
    df["3% Royalty Fee Master"] = df["3% Royalty Fee Group"]

    df["5% Royalty Fee Group"] = df.apply(
        lambda r: 0.0 if contains_ci(r["Service and Name"], "BGIS SCS regular") else r[col_total] * 0.05,
        axis=1
    )
    df["5% Royalty Fee Master2"] = df["5% Royalty Fee Group"]

    df["1% Marketing Fee"] = df[col_total] * 0.01

    df["Brokerages"] = df["Brokerage"].apply(lambda x: 0.05 if x == "Brokerage5" else 0.0)
    df["5% Brokerage Fee"] = df.apply(
        lambda r: r[col_total] * r["Brokerages"] if r["Brokerages"] == 0.05 else 0.0,
        axis=1
    )
    df["2.5% Brokerage Fee"] = 0.0

    df = add_province_from_taxes(df)
    return df

# =========================================================
# Totals helpers
# =========================================================
def add_grand_total_row(df: pd.DataFrame, label_cols: list[str], total_label: str = "Grand Total") -> pd.DataFrame:
    df2 = df.copy()
    num_cols = df2.select_dtypes(include="number").columns.tolist()

    total_row = {c: "" for c in df2.columns}
    for c in label_cols:
        if c in df2.columns:
            total_row[c] = total_label

    for c in num_cols:
        total_row[c] = float(df2[c].sum())

    return pd.concat([df2, pd.DataFrame([total_row])], ignore_index=True)

def show_kpis_from_df(df: pd.DataFrame, hide_marketing: bool = False):
    total = float(df.get("Total Amount Without Taxes", pd.Series([0])).sum())
    r3 = float(df.get("3% Royalty Fee Group", pd.Series([0])).sum())
    r5 = float(df.get("5% Royalty Fee Group", pd.Series([0])).sum())
    m1 = float(df.get("1% Marketing Fee", pd.Series([0])).sum()) if (not hide_marketing) else None
    b5 = float(df.get("5% Brokerage Fee", pd.Series([0])).sum())
    b25 = float(df.get("2.5% Brokerage Fee", pd.Series([0])).sum())

    if hide_marketing:
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total Amount (No Taxes)", fmt_money(total))
        c2.metric("3% Royalty Fee", fmt_money(r3))
        c3.metric("5% Royalty Fee", fmt_money(r5))
        c4.metric("5% Brokerage Fee", fmt_money(b5))
        c5.metric("2.5% Brokerage Fee", fmt_money(b25))
    else:
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Total Amount (No Taxes)", fmt_money(total))
        c2.metric("3% Royalty Fee", fmt_money(r3))
        c3.metric("5% Royalty Fee", fmt_money(r5))
        c4.metric("1% Marketing Fee", fmt_money(m1 or 0.0))
        c5.metric("5% Brokerage Fee", fmt_money(b5))
        c6.metric("2.5% Brokerage Fee", fmt_money(b25))

# =========================================================
# Jeff exclusions: ONLY Buyer Company Name
# =========================================================
JEFF_EXCLUDE_BUYERS = [
    "12433087 Canada Inc",
    "12433087 Canada Inc - Master",
    "Rojo construction Management Inc",
    "Academie St Laurent Academy Inc",
    "10342548 Canada Inc",
    "Osgoode Properties Limited",
    "2501308 Ontario Inc",
    "Hallmark Housekeeping",
    "Hotel plaza de la Chaudiere Inc",
    "Allen Maintenance Ltd (do not use)",
    "CCI Ottawa",
    "13037622 Canada Inc",
    "Aylmer Street Developments",
    "Syndicat de Copropietaires Jardins Maisonneuve",
    "Ladies Space",
    "INDIGO PARK CANADA Inc",
    "TAYANTI OTTAWA INC",
    "TERLIN CONSTRUCTION LTD",
    "ICS CLEAN INC",
    "ALPINE BUILDING MAINTENANCE INC.",
    "Stationnements Parkeo Inc",
    "Allen Maintenance Ltd",
    "Evripos",
]

def apply_jeff_exclusions_only_buyer(df: pd.DataFrame) -> pd.DataFrame:
    if "Buyer Company Name" not in df.columns:
        return df
    buyer = df["Buyer Company Name"].astype(str).fillna("")
    mask_exclude = pd.Series(False, index=df.index)
    for name in JEFF_EXCLUDE_BUYERS:
        mask_exclude = mask_exclude | buyer.str.contains(re.escape(name), case=False, na=False)
    return df[~mask_exclude].copy()

# =========================================================
# Reports
# =========================================================
def report_resume_without_fees(df: pd.DataFrame) -> pd.DataFrame:
    pivot = pd.pivot_table(
        df,
        index=["Vendor Company Name", "Service", "Buyer Company Name"],
        columns=["Province"],
        values="Total Amount Without Taxes",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    province_cols = [c for c in pivot.columns if c not in ["Vendor Company Name", "Service", "Buyer Company Name"]]
    pivot["Row Total"] = pivot[province_cols].sum(axis=1) if province_cols else 0.0

    pivot = add_grand_total_row(pivot, label_cols=["Vendor Company Name"], total_label="Grand Total")
    return pivot

def report_validation(df: pd.DataFrame, hide_marketing: bool = False) -> pd.DataFrame:
    cols = [
        "Total Amount Without Taxes",
        "3% Royalty Fee Group",
        "5% Royalty Fee Group",
        "1% Marketing Fee",
        "5% Brokerage Fee",
        "2.5% Brokerage Fee",
    ]
    if hide_marketing:
        cols = [c for c in cols if c != "1% Marketing Fee"]
    cols = [c for c in cols if c in df.columns]

    rep = (
        df.groupby(["Province", "Vendor Company Name", "Brokerage", "Buyer Company Name"], dropna=False)[cols]
          .sum()
          .reset_index()
    )
    rep = add_grand_total_row(rep, label_cols=["Province", "Vendor Company Name"], total_label="Grand Total")
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
        ["Resume Without Fees", "Validation", "Jeff-validation"],
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

    st.divider()
    run = st.button("Download + Process")

if run:
    with st.spinner("Downloading export from CNET..."):
        content = download_export_file()

    df_raw = read_any_table(content)
    if "Creation Date" not in df_raw.columns:
        st.error("No encuentro la columna 'Creation Date' para filtrar por mes.")
        st.stop()

    df_raw["Creation Date"] = pd.to_datetime(df_raw["Creation Date"], errors="coerce")
    df_month = df_raw[
        (df_raw["Creation Date"].dt.month == month_num) &
        (df_raw["Creation Date"].dt.year == int(year))
    ].copy()

    df_final = build_columns(df_month)
    st.session_state["df_final"] = df_final
    st.session_state["loaded_month"] = month_name
    st.session_state["loaded_year"] = int(year)

if "df_final" not in st.session_state:
    st.info("Selecciona Month/Year y presiona **Download + Process** para cargar la data.")
    st.stop()

df_final = st.session_state["df_final"]
loaded_month = st.session_state.get("loaded_month", month_name)
loaded_year = st.session_state.get("loaded_year", int(year))
st.success(f"Data cargada en memoria: {loaded_month} {loaded_year} | Rows: {len(df_final):,}")

# Filters
company_col = "Company Name" if "Company Name" in df_final.columns else ("Company" if "Company" in df_final.columns else None)

with st.sidebar:
    st.divider()
    st.subheader("Filters")

    if company_col:
        companies = sorted(df_final[company_col].dropna().astype(str).unique().tolist())
        sel_company = st.multiselect("Company", companies, default=[])
    else:
        sel_company = []

    services = sorted(df_final["Service"].dropna().astype(str).unique().tolist())
    sel_service = st.multiselect("Service", services, default=[])

    provinces = sorted(df_final["Province"].dropna().astype(str).unique().tolist())
    sel_province = st.multiselect("Province", provinces, default=[])

    buyers = sorted(df_final["Buyer Company Name"].dropna().astype(str).unique().tolist())
    sel_buyer = st.multiselect("Buyer Company Name", buyers, default=[])

    brokerages = sorted(df_final["Brokerage"].dropna().astype(str).unique().tolist())
    sel_brokerage = st.multiselect("Brokerage", brokerages, default=[])

    vendors = sorted(df_final["Vendor Company Name"].dropna().astype(str).unique().tolist())
    sel_vendor = st.multiselect("Vendor Company Name", vendors, default=[])

df_filtered = df_final.copy()
if company_col and sel_company:
    df_filtered = df_filtered[df_filtered[company_col].astype(str).isin(sel_company)]
if sel_service:
    df_filtered = df_filtered[df_filtered["Service"].astype(str).isin(sel_service)]
if sel_province:
    df_filtered = df_filtered[df_filtered["Province"].astype(str).isin(sel_province)]
if sel_buyer:
    df_filtered = df_filtered[df_filtered["Buyer Company Name"].astype(str).isin(sel_buyer)]
if sel_brokerage:
    df_filtered = df_filtered[df_filtered["Brokerage"].astype(str).isin(sel_brokerage)]
if sel_vendor:
    df_filtered = df_filtered[df_filtered["Vendor Company Name"].astype(str).isin(sel_vendor)]

st.caption(f"Filtered rows: {len(df_filtered):,}")

# Jeff: apply exclusions only on Buyer Company Name + hide marketing
hide_marketing = (report_choice == "Jeff-validation")
df_for_report = df_filtered
if report_choice == "Jeff-validation":
    df_for_report = apply_jeff_exclusions_only_buyer(df_filtered)

show_kpis_from_df(df_for_report, hide_marketing=hide_marketing)

# Report
if report_choice == "Resume Without Fees":
    st.subheader("Resume Without Fees (with totals)")
    rep = report_resume_without_fees(df_for_report)
elif report_choice == "Validation":
    st.subheader("Validation (with totals)")
    rep = report_validation(df_for_report, hide_marketing=False)
else:
    st.subheader("Jeff-validation (NO Marketing Fee) + exclusions (Buyer only) + totals")
    rep = report_validation(df_for_report, hide_marketing=True)

# ✅ Show formatted table (2 decimals + thousands)
st.dataframe(format_report_for_display(rep), use_container_width=True)

st.download_button(
    label="Download selected report (CSV)",
    data=df_to_csv_bytes(rep),  # download keeps numeric values
    file_name=f"{report_choice.replace(' ', '_').lower()}_{loaded_month.lower()}_{loaded_year}.csv",
    mime="text/csv"
)

# Preview (raw, for audit)
st.divider()
st.subheader("Preview - Calculated Columns (include Province)")

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
if company_col:
    cols_show.insert(0, company_col)

cols_show = [c for c in cols_show if c in df_for_report.columns]
st.dataframe(df_for_report[cols_show].head(200), use_container_width=True)
