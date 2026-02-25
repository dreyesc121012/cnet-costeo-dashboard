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


def fmt_money(x: float) -> str:
    return f"${x:,.2f}"


def format_report_for_display(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    num_cols = out.select_dtypes(include="number").columns.tolist()
    for c in num_cols:
        out[c] = out[c].map(lambda v: f"{v:,.2f}")
    return out


def norm_name(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip().lower())


def is_bgis_scs_regular(service_and_name: str) -> bool:
    s = re.sub(r"\s+", " ", str(service_and_name or "").strip().lower())
    return ("bgis scs" in s) and ("regular" in s)


# =========================================================
# Download with login
# =========================================================

def download_export_file() -> bytes:
    session = requests.Session()

    r1 = session.get("https://app.master.cnetfranchise.com/login", timeout=60)
    soup = BeautifulSoup(r1.text, "lxml")
    csrf = soup.select_one('input[name="_csrf_token"]')["value"]

    payload = {
        "_csrf_token": csrf,
        "_username": st.secrets["SYSTEM_USERNAME"],
        "_password": st.secrets["SYSTEM_PASSWORD"],
        "_submit": "Login",
    }

    session.post("https://app.master.cnetfranchise.com/login_check", data=payload)
    export_url = st.secrets["EXPORT_EXCEL_URL"]
    r3 = session.get(export_url)
    return r3.content


def read_any_table(file_bytes: bytes) -> pd.DataFrame:
    try:
        return pd.read_excel(io.BytesIO(file_bytes))
    except:
        return pd.read_csv(io.BytesIO(file_bytes))


# =========================================================
# Province detection
# =========================================================

def add_province_from_taxes(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    tax_cols = [c for c in df.columns if "TAX" in c.upper() or "GST" in c.upper() or "QST" in c.upper()]

    def detect(row):
        for c in tax_cols:
            if abs(safe_num(row.get(c, 0))) > 0:
                name = c.upper()
                if "QST" in name or "QC" in name:
                    return "Quebec"
                if "HST" in name or "ON" in name:
                    return "Ontario"
                if "NS" in name:
                    return "Nova Scotia"
        return "Unknown"

    df["Province"] = df.apply(detect, axis=1)
    return df


# =========================================================
# Build Calculated Columns
# =========================================================

def build_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    col_total = "Total Amount Without Taxes"

    df[col_total] = df[col_total].apply(safe_num)

    df["Service"] = df["Work Description"].apply(
        lambda x: "Regular" if "janitorial" in str(x).lower() else "One Shot"
    )

    df["Service and Name"] = df["Service"] + " " + df["Buyer Company Name"]

    df["Brokerage"] = df["Buyer Company Name"].apply(
        lambda x: "Brokerage5" if "5bf" in str(x).lower() else "Without Brokerage"
    )

    # 3%
    df["3% Royalty Fee Group"] = df.apply(
        lambda r: r[col_total] * 0.03 if is_bgis_scs_regular(r["Service and Name"]) else 0.0,
        axis=1
    )
    df["3% Royalty Fee Master"] = df["3% Royalty Fee Group"]

    # 5%
    df["5% Royalty Fee Group"] = df.apply(
        lambda r: 0.0 if is_bgis_scs_regular(r["Service and Name"]) else r[col_total] * 0.05,
        axis=1
    )
    df["5% Royalty Fee Master2"] = df["5% Royalty Fee Group"]

    df["1% Marketing Fee"] = df[col_total] * 0.01

    df["Brokerages"] = df["Brokerage"].apply(lambda x: 0.05 if x == "Brokerage5" else 0.0)
    df["5% Brokerage Fee"] = df[col_total] * df["Brokerages"]
    df["2.5% Brokerage Fee"] = 0.0

    df = add_province_from_taxes(df)
    return df


# =========================================================
# Jeff exclusions EXACT
# =========================================================

JEFF_EXCLUDE_BUYERS = [
    "Allen Maintenance Ltd",
    "12433087 Canada Inc",
    "CCI Ottawa",
    "ICS CLEAN INC",
]

EXCLUDE_SET = set(norm_name(x) for x in JEFF_EXCLUDE_BUYERS)


def apply_jeff_exclusions_only_buyer_exact(df: pd.DataFrame):
    buyer_norm = df["Buyer Company Name"].map(norm_name)
    mask = buyer_norm.isin(EXCLUDE_SET)
    return df[~mask].copy()


# =========================================================
# Totals
# =========================================================

def add_grand_total_row(df):
    total_row = {}
    for c in df.columns:
        if df[c].dtype != object:
            total_row[c] = df[c].sum()
        else:
            total_row[c] = "Grand Total"
    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)


def show_kpis(df, hide_marketing=False):
    total = df["Total Amount Without Taxes"].sum()
    r3 = df["3% Royalty Fee Group"].sum()
    r5 = df["5% Royalty Fee Group"].sum()
    b5 = df["5% Brokerage Fee"].sum()
    b25 = df["2.5% Brokerage Fee"].sum()

    if hide_marketing:
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total Amount", fmt_money(total))
        c2.metric("3% Royalty", fmt_money(r3))
        c3.metric("5% Royalty", fmt_money(r5))
        c4.metric("5% Brokerage", fmt_money(b5))
        c5.metric("2.5% Brokerage", fmt_money(b25))
    else:
        m1 = df["1% Marketing Fee"].sum()
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Total Amount", fmt_money(total))
        c2.metric("3% Royalty", fmt_money(r3))
        c3.metric("5% Royalty", fmt_money(r5))
        c4.metric("1% Marketing", fmt_money(m1))
        c5.metric("5% Brokerage", fmt_money(b5))
        c6.metric("2.5% Brokerage", fmt_money(b25))


# =========================================================
# UI
# =========================================================

st.title("CNET - Invoice Reports")

with st.sidebar:
    report = st.radio("Report", ["Resume Without Fees", "Validation", "Jeff-validation"])
    run = st.button("Download + Process")

if run:
    content = download_export_file()
    df_raw = read_any_table(content)
    df = build_columns(df_raw)
    st.session_state["df"] = df

if "df" not in st.session_state:
    st.stop()

df = st.session_state["df"]

if report == "Jeff-validation":
    df = apply_jeff_exclusions_only_buyer_exact(df)
    hide_marketing = True
else:
    hide_marketing = False

show_kpis(df, hide_marketing)

group_cols = ["Province", "Vendor Company Name", "Brokerage", "Buyer Company Name"]

sum_cols = [
    "Total Amount Without Taxes",
    "3% Royalty Fee Group",
    "5% Royalty Fee Group",
    "5% Brokerage Fee",
    "2.5% Brokerage Fee"
]

if not hide_marketing:
    sum_cols.insert(3, "1% Marketing Fee")

report_df = df.groupby(group_cols)[sum_cols].sum().reset_index()
report_df = add_grand_total_row(report_df)

st.dataframe(format_report_for_display(report_df), use_container_width=True)

st.download_button(
    "Download selected report (CSV)",
    report_df.to_csv(index=False).encode("utf-8"),
    file_name="report.csv"
)
