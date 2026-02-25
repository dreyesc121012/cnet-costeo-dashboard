import io
import re
import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

# =========================================================
# CONFIG
# =========================================================
st.set_page_config(page_title="CNET - Invoice Reports", layout="wide")
TITLE = "CNET - Invoice Reports"

# Jeff exclusions (Buyer ONLY, exact match)
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

# =========================================================
# Helpers
# =========================================================
def norm_name(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip().lower())

EXCLUDE_SET_NORM = set(norm_name(x) for x in JEFF_EXCLUDE_BUYERS)

def safe_num(x) -> float:
    try:
        if pd.isna(x):
            return 0.0
        s = str(x).strip()
        if not s:
            return 0.0
        s = s.replace("$", "").replace(",", "").strip()
        if s.startswith("(") and s.endswith(")"):
            s = "-" + s[1:-1]
        return float(s)
    except Exception:
        return 0.0

def fmt_money(x: float) -> str:
    try:
        return f"${x:,.2f}"
    except Exception:
        return "$0.00"

def format_report_for_display(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    num_cols = out.select_dtypes(include="number").columns.tolist()
    for c in num_cols:
        out[c] = out[c].map(lambda v: f"{float(v):,.2f}")
    return out

def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")

def month_name_to_num(name: str) -> int:
    months = {
        "January": 1, "February": 2, "March": 3, "April": 4,
        "May": 5, "June": 6, "July": 7, "August": 8,
        "September": 9, "October": 10, "November": 11, "December": 12,
    }
    return months[name]

def get_company_col(df: pd.DataFrame):
    if "Company Name" in df.columns:
        return "Company Name"
    if "Company" in df.columns:
        return "Company"
    return None

def is_bgis_scs_regular(service_and_name: str) -> bool:
    s = re.sub(r"\s+", " ", str(service_and_name or "").strip().lower())
    return ("bgis scs" in s) and ("regular" in s)

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
    r3 = session.get(export_url, timeout=120)
    r3.raise_for_status()
    return r3.content

def read_any_table(file_bytes: bytes) -> pd.DataFrame:
    try:
        return pd.read_excel(io.BytesIO(file_bytes))
    except Exception:
        try:
            return pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8", engine="python")
        except Exception:
            return pd.read_csv(io.BytesIO(file_bytes), encoding="latin1", engine="python")

# =========================================================
# Province inference (from taxes columns)
# =========================================================
def add_province_from_taxes(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    tax_cols = [c for c in df.columns if any(k in c.upper() for k in ["GST", "HST", "QST", "PST", "RST", "TAX"])]

    mapping = {
        "QC": "Quebec", "ON": "Ontario", "BC": "British Columbia", "AB": "Alberta",
        "MB": "Manitoba", "SK": "Saskatchewan", "NB": "New Brunswick", "NS": "Nova Scotia",
        "NL": "Newfoundland and Labrador", "PE": "Prince Edward Island",
        "NT": "Northwest Territories", "NU": "Nunavut", "YT": "Yukon",
        "PEI": "Prince Edward Island",
    }

    def province_from_row(row) -> str:
        hits = []
        for c in tax_cols:
            v = safe_num(row.get(c, 0))
            if abs(v) > 0:  # includes negatives
                hits.append(c.upper())

        if not hits:
            txt = f"{row.get('Buyer Company Name','')} {row.get('Vendor Company Name','')} {row.get('Work Description','')}".upper()
            if "QUEBEC" in txt or re.search(r"\bQC\b", txt):
                return "Quebec"
            if "ONTARIO" in txt or re.search(r"\bON\b", txt):
                return "Ontario"
            if "NOVA SCOTIA" in txt or re.search(r"\bNS\b", txt):
                return "Nova Scotia"
            if "NEW BRUNSWICK" in txt or re.search(r"\bNB\b", txt):
                return "New Brunswick"
            if "PRINCE EDWARD" in txt or "PEI" in txt or re.search(r"\bPE\b", txt):
                return "Prince Edward Island"
            return "Unknown"

        for c in hits:
            if ("QST" in c) and ("QC" in c):
                return "Quebec"

        for c in hits:
            for code, prov in mapping.items():
                if re.search(rf"\b{re.escape(code)}\b", c):
                    return prov

        for c in hits:
            if "QST" in c:
                return "Quebec"
            if "HST" in c:
                return "HST Province (Unknown)"
            if "GST" in c:
                return "GST Only (Unknown)"

        return "Unknown"

    df["Province"] = df.apply(province_from_row, axis=1)
    return df

# =========================================================
# Build calculated columns (YOUR RULES)
# =========================================================
def build_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    required = ["Work Description", "Buyer Company Name", "Vendor Company Name", "Total Amount Without Taxes"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise Exception(f"Faltan columnas requeridas en el archivo: {missing}")

    col_total = "Total Amount Without Taxes"
    df[col_total] = df[col_total].apply(safe_num)

    df["Service"] = df["Work Description"].astype(str).apply(
        lambda x: "Regular" if "janitorial" in x.lower() else "One Shot"
    )

    df["Service and Name"] = df["Service"].astype(str) + " " + df["Buyer Company Name"].astype(str)

    df["Brokerage"] = df["Buyer Company Name"].astype(str).apply(
        lambda x: "Brokerage5" if "5bf" in x.lower() else "Without Brokerage"
    )

    df["3% Royalty Fee Group"] = df.apply(
        lambda r: r[col_total] * 0.03 if is_bgis_scs_regular(r["Service and Name"]) else 0.0,
        axis=1
    )
    df["3% Royalty Fee Master"] = df["3% Royalty Fee Group"]

    # ✅ FIX: 5% royalty = 0 if BGIS SCS + Regular (any order), else Total*0.05
    df["5% Royalty Fee Group"] = df.apply(
        lambda r: 0.0 if is_bgis_scs_regular(r["Service and Name"]) else r[col_total] * 0.05,
        axis=1
    )
    df["5% Royalty Fee Master2"] = df["5% Royalty Fee Group"]

    df["1% Marketing Fee"] = df[col_total] * 0.01

    df["Brokerages"] = df["Brokerage"].astype(str).apply(lambda x: 0.05 if x == "Brokerage5" else 0.0)

    df["5% Brokerage Fee"] = df.apply(
        lambda r: r[col_total] * r["Brokerages"] if float(r["Brokerages"]) > 0 else 0.0,
        axis=1
    )

    df["2.5% Brokerage Fee"] = 0.0

    df = add_province_from_taxes(df)

    numeric_cols = [
        "Total Amount Without Taxes",
        "3% Royalty Fee Group", "3% Royalty Fee Master",
        "5% Royalty Fee Group", "5% Royalty Fee Master2",
        "1% Marketing Fee",
        "Brokerages", "5% Brokerage Fee", "2.5% Brokerage Fee",
    ]
    for c in numeric_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    return df

# =========================================================
# Jeff exclusions (Buyer ONLY, exact match)
# =========================================================
def apply_jeff_exclusions_only_buyer_exact(df: pd.DataFrame) -> pd.DataFrame:
    if "Buyer Company Name" not in df.columns:
        return df.copy()
    buyer_norm = df["Buyer Company Name"].astype(str).map(norm_name)
    mask_exclude = buyer_norm.isin(EXCLUDE_SET_NORM)  # exact match only
    return df[~mask_exclude].copy()

# =========================================================
# Totals + KPIs
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

def kpi_values(df: pd.DataFrame, hide_marketing: bool):
    total = float(df.get("Total Amount Without Taxes", pd.Series([0])).sum())
    r3 = float(df.get("3% Royalty Fee Group", pd.Series([0])).sum())
    r5 = float(df.get("5% Royalty Fee Group", pd.Series([0])).sum())
    b5 = float(df.get("5% Brokerage Fee", pd.Series([0])).sum())
    b25 = float(df.get("2.5% Brokerage Fee", pd.Series([0])).sum())
    m1 = float(df.get("1% Marketing Fee", pd.Series([0])).sum()) if not hide_marketing else None
    return {"total": total, "r3": r3, "r5": r5, "m1": m1, "b5": b5, "b25": b25}

def show_kpis_from_df(df: pd.DataFrame, hide_marketing: bool = False):
    vals = kpi_values(df, hide_marketing)
    if hide_marketing:
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total Amount (No Taxes)", fmt_money(vals["total"]))
        c2.metric("3% Royalty Fee", fmt_money(vals["r3"]))
        c3.metric("5% Royalty Fee", fmt_money(vals["r5"]))
        c4.metric("5% Brokerage Fee", fmt_money(vals["b5"]))
        c5.metric("2.5% Brokerage Fee", fmt_money(vals["b25"]))
    else:
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Total Amount (No Taxes)", fmt_money(vals["total"]))
        c2.metric("3% Royalty Fee", fmt_money(vals["r3"]))
        c3.metric("5% Royalty Fee", fmt_money(vals["r5"]))
        c4.metric("1% Marketing Fee", fmt_money(vals["m1"] or 0.0))
        c5.metric("5% Brokerage Fee", fmt_money(vals["b5"]))
        c6.metric("2.5% Brokerage Fee", fmt_money(vals["b25"]))

def delta_pct(curr: float, prev: float) -> float:
    if prev == 0:
        return 0.0
    return (curr - prev) / prev * 100.0

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
        fill_value=0.0
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

# =========================================================
# Executive Summary + Comparison
# =========================================================
def executive_summary(df: pd.DataFrame):
    st.subheader("Executive Summary")

    # By Province
    prov_sum = (
        df.groupby("Province", dropna=False)[
            ["Total Amount Without Taxes", "3% Royalty Fee Group", "5% Royalty Fee Group", "5% Brokerage Fee", "2.5% Brokerage Fee"]
        ]
        .sum()
        .reset_index()
    )
    prov_sum["Province"] = prov_sum["Province"].fillna("Unknown")

    st.markdown("### By Province (Total Amount Without Taxes)")
    c1, c2 = st.columns([2, 1])
    with c1:
        st.bar_chart(prov_sum.set_index("Province")["Total Amount Without Taxes"].sort_values(ascending=False))
    with c2:
        st.dataframe(format_report_for_display(prov_sum.sort_values("Total Amount Without Taxes", ascending=False)), use_container_width=True)

    # Service Mix
    st.markdown("### Service Mix (Regular vs One Shot)")
    mix = (
        df.groupby("Service")["Total Amount Without Taxes"]
        .sum()
        .reset_index()
        .sort_values("Total Amount Without Taxes", ascending=False)
    )
    m1, m2 = st.columns([1, 2])
    with m1:
        st.dataframe(format_report_for_display(mix), use_container_width=True)
    with m2:
        st.bar_chart(mix.set_index("Service")["Total Amount Without Taxes"])

    # Top 10 Buyers / Vendors
    st.markdown("### Concentration (Top 10)")
    t1, t2 = st.columns(2)

    top_buyers = (
        df.groupby("Buyer Company Name")["Total Amount Without Taxes"]
        .sum()
        .reset_index()
        .sort_values("Total Amount Without Taxes", ascending=False)
        .head(10)
    )
    top_vendors = (
        df.groupby("Vendor Company Name")["Total Amount Without Taxes"]
        .sum()
        .reset_index()
        .sort_values("Total Amount Without Taxes", ascending=False)
        .head(10)
    )

    with t1:
        st.markdown("**Top 10 Buyers**")
        st.bar_chart(top_buyers.set_index("Buyer Company Name")["Total Amount Without Taxes"])
        st.dataframe(format_report_for_display(top_buyers), use_container_width=True)

    with t2:
        st.markdown("**Top 10 Vendors**")
        st.bar_chart(top_vendors.set_index("Vendor Company Name")["Total Amount Without Taxes"])
        st.dataframe(format_report_for_display(top_vendors), use_container_width=True)

def comparison_section(df_curr: pd.DataFrame, df_prev: pd.DataFrame, hide_marketing: bool, label_curr: str, label_prev: str):
    st.subheader(f"Comparison: {label_curr} vs {label_prev}")

    curr = kpi_values(df_curr, hide_marketing)
    prev = kpi_values(df_prev, hide_marketing)

    # KPI comparison cards
    if hide_marketing:
        cols = st.columns(5)
        items = [
            ("Total (No Taxes)", "total"),
            ("3% Royalty", "r3"),
            ("5% Royalty", "r5"),
            ("5% Brokerage", "b5"),
            ("2.5% Brokerage", "b25"),
        ]
    else:
        cols = st.columns(6)
        items = [
            ("Total (No Taxes)", "total"),
            ("3% Royalty", "r3"),
            ("5% Royalty", "r5"),
            ("1% Marketing", "m1"),
            ("5% Brokerage", "b5"),
            ("2.5% Brokerage", "b25"),
        ]

    for i, (title, key) in enumerate(items):
        c = cols[i]
        curr_val = curr[key] if curr[key] is not None else 0.0
        prev_val = prev[key] if prev[key] is not None else 0.0
        delta_val = curr_val - prev_val
        delta_str = f"{fmt_money(delta_val)} ({delta_pct(curr_val, prev_val):.1f}%)" if prev_val != 0 else fmt_money(delta_val)
        c.metric(title, fmt_money(curr_val), delta_str)

    st.markdown("### By Province (Total Amount Without Taxes) - Comparison")

    curr_prov = df_curr.groupby("Province")["Total Amount Without Taxes"].sum().rename("Current")
    prev_prov = df_prev.groupby("Province")["Total Amount Without Taxes"].sum().rename("Compare")

    prov_compare = pd.concat([curr_prov, prev_prov], axis=1).fillna(0.0).reset_index().rename(columns={"index": "Province"})
    prov_compare["Delta"] = prov_compare["Current"] - prov_compare["Compare"]

    st.dataframe(format_report_for_display(prov_compare.sort_values("Current", ascending=False)), use_container_width=True)

    # bar chart with two columns in a single df (Streamlit bar_chart supports multi-series)
    chart_df = prov_compare.set_index("Province")[["Current", "Compare"]].sort_values("Current", ascending=False)
    st.bar_chart(chart_df)

# =========================================================
# UI
# =========================================================
st.title(TITLE)

month_labels = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

with st.sidebar:
    st.subheader("Select report")
    report_choice = st.radio("Report type", ["Resume Without Fees", "Validation", "Jeff-validation"], index=0)

    st.divider()
    st.subheader("Month filter (Main)")
    month_name = st.selectbox("Month", month_labels, index=0, key="main_month")
    year = st.number_input("Year", min_value=2020, max_value=2035, value=2026, step=1, key="main_year")

    st.divider()
    st.subheader("Comparison")
    enable_compare = st.checkbox("Enable comparison", value=True)
    compare_month = st.selectbox("Compare month", month_labels, index=1, key="cmp_month")  # default February
    compare_year = st.number_input("Compare year", min_value=2020, max_value=2035, value=2026, step=1, key="cmp_year")

    st.divider()
    run = st.button("Download + Process", type="primary")

# Download/process on click (store raw data for flexible month filters)
if run:
    with st.spinner("Downloading export from CNET..."):
        content = download_export_file()
    df_raw = read_any_table(content)

    if "Creation Date" not in df_raw.columns:
        st.error("No encuentro la columna 'Creation Date' en el export. Revisa el archivo exportado.")
        st.stop()

    df_raw["Creation Date"] = pd.to_datetime(df_raw["Creation Date"], errors="coerce")

    st.session_state["df_raw"] = df_raw  # ✅ keep raw
    st.session_state["last_download_ok"] = True

# Need raw data first
if "df_raw" not in st.session_state:
    st.info("Selecciona Month/Year y presiona **Download + Process** para cargar la data.")
    st.stop()

df_raw = st.session_state["df_raw"].copy()

# Build month df (main)
m = month_name_to_num(month_name)
y = int(year)
df_month = df_raw[(df_raw["Creation Date"].dt.month == m) & (df_raw["Creation Date"].dt.year == y)].copy()
df_main = build_columns(df_month)

st.success(f"Data cargada en memoria: {month_name} {y} | Rows: {len(df_main):,}")

# Build compare df if enabled
df_cmp = None
label_main = f"{month_name} {y}"
label_cmp = f"{compare_month} {int(compare_year)}"

if enable_compare:
    cm = month_name_to_num(compare_month)
    cy = int(compare_year)
    df_month_cmp = df_raw[(df_raw["Creation Date"].dt.month == cm) & (df_raw["Creation Date"].dt.year == cy)].copy()
    df_cmp = build_columns(df_month_cmp)

# ============== Sidebar Filters (filters only) ==============
company_col = get_company_col(df_main)

with st.sidebar:
    st.divider()
    st.subheader("Filters")

    if company_col:
        companies = sorted(df_main[company_col].dropna().astype(str).unique().tolist())
        sel_company = st.multiselect("Company", companies, default=[])
    else:
        sel_company = []

    services = sorted(df_main["Service"].dropna().astype(str).unique().tolist())
    sel_service = st.multiselect("Service", services, default=[])

    provinces = sorted(df_main["Province"].dropna().astype(str).unique().tolist())
    sel_province = st.multiselect("Province", provinces, default=[])

    buyers = sorted(df_main["Buyer Company Name"].dropna().astype(str).unique().tolist())
    sel_buyer = st.multiselect("Buyer Company Name", buyers, default=[])

    brokerages = sorted(df_main["Brokerage"].dropna().astype(str).unique().tolist())
    sel_brokerage = st.multiselect("Brokerage", brokerages, default=[])

    vendors = sorted(df_main["Vendor Company Name"].dropna().astype(str).unique().tolist())
    sel_vendor = st.multiselect("Vendor Company Name", vendors, default=[])

def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if company_col and sel_company:
        out = out[out[company_col].astype(str).isin(sel_company)]
    if sel_service:
        out = out[out["Service"].astype(str).isin(sel_service)]
    if sel_province:
        out = out[out["Province"].astype(str).isin(sel_province)]
    if sel_buyer:
        out = out[out["Buyer Company Name"].astype(str).isin(sel_buyer)]
    if sel_brokerage:
        out = out[out["Brokerage"].astype(str).isin(sel_brokerage)]
    if sel_vendor:
        out = out[out["Vendor Company Name"].astype(str).isin(sel_vendor)]
    return out

df_filtered = apply_filters(df_main)
st.caption(f"Filtered rows: {len(df_filtered):,}")

# Jeff rules apply only to Jeff-validation
hide_marketing = (report_choice == "Jeff-validation")

df_for_report = df_filtered
if report_choice == "Jeff-validation":
    df_for_report = apply_jeff_exclusions_only_buyer_exact(df_filtered)

# KPIs
show_kpis_from_df(df_for_report, hide_marketing=hide_marketing)

# =========================================================
# Executive + Comparison
# =========================================================
st.divider()
executive_summary(df_for_report)

if enable_compare and df_cmp is not None:
    # apply same filters to compare month (except options list comes from main; still works)
    df_cmp_filtered = apply_filters(df_cmp)

    if report_choice == "Jeff-validation":
        df_cmp_filtered = apply_jeff_exclusions_only_buyer_exact(df_cmp_filtered)

    st.divider()
    comparison_section(
        df_curr=df_for_report,
        df_prev=df_cmp_filtered,
        hide_marketing=hide_marketing,
        label_curr=label_main,
        label_prev=label_cmp
    )

# =========================================================
# Build selected report
# =========================================================
st.divider()

if report_choice == "Resume Without Fees":
    st.subheader("Resume Without Fees (with totals)")
    rep = report_resume_without_fees(df_for_report)

elif report_choice == "Validation":
    st.subheader("Validation (with totals)")
    rep = report_validation(df_for_report, hide_marketing=False)

else:
    st.subheader("Jeff-validation (NO Marketing Fee) + EXACT Buyer exclusions + totals")
    rep = report_validation(df_for_report, hide_marketing=True)

st.dataframe(format_report_for_display(rep), use_container_width=True)

st.download_button(
    label="Download selected report (CSV)",
    data=df_to_csv_bytes(rep),
    file_name=f"{report_choice.replace(' ', '_').lower()}_{month_name.lower()}_{y}.csv",
    mime="text/csv"
)

# =========================================================
# Preview
# =========================================================
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
st.dataframe(df_for_report[cols_show].head(300), use_container_width=True)
