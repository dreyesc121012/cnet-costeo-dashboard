import base64
from io import BytesIO
from datetime import timedelta, datetime
import unicodedata

import pandas as pd
import requests
import streamlit as st
import msal


# ============================================================
# CONFIG
# ============================================================
COMITE_CLASS_A_RATE = 21.57
REER_PER_HOUR = 0.45

st.set_page_config(page_title="Comité Paritaire QC", layout="wide")
st.title("Comité Paritaire Québec - Weekly Report")


# ============================================================
# HELPERS
# ============================================================
def normalize_text(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s


def safe_text_series(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip().fillna("")


# ============================================================
# COMMITTEE LOGIC (🔥 CORRECTA)
# ============================================================
def calculate_weekly_committee_hours(total_pay, raw_hours_sum, hourly_rate_min):
    try:
        total_pay = float(total_pay)
    except:
        total_pay = 0

    try:
        raw_hours_sum = float(raw_hours_sum)
    except:
        raw_hours_sum = 0

    try:
        hourly_rate_min = float(hourly_rate_min)
    except:
        hourly_rate_min = 0

    # 🔥 CASO 1: salario menor a comité
    if hourly_rate_min > 0 and hourly_rate_min < COMITE_CLASS_A_RATE:
        return round(total_pay / COMITE_CLASS_A_RATE, 2)

    # 🔥 CASO 2: trabajo flat
    if raw_hours_sum <= 1 and total_pay > COMITE_CLASS_A_RATE:
        return round(total_pay / COMITE_CLASS_A_RATE, 2)

    # 🔥 CASO 3: salario mayor o igual
    return round(raw_hours_sum, 2)


# ============================================================
# BUILD SUMMARY
# ============================================================
def build_weekly_summary(df):
    grouped = (
        df.groupby(["vendor_company", "employee", "week_label"])
        .agg(
            raw_hours_sum=("hours", "sum"),
            total_pay=("total_pay", "sum"),
            hourly_rate_min=("hourly_rate", "min"),
        )
        .reset_index()
    )

    grouped["committee_hours"] = grouped.apply(
        lambda r: calculate_weekly_committee_hours(
            r["total_pay"],
            r["raw_hours_sum"],
            r["hourly_rate_min"],
        ),
        axis=1,
    )

    grouped["reer"] = grouped["committee_hours"] * REER_PER_HOUR
    grouped["total_with_reer"] = grouped["total_pay"] + grouped["reer"]

    return grouped


# ============================================================
# SAMPLE DATA (usa tus datos reales aquí)
# ============================================================
df = pd.DataFrame({
    "vendor_company": ["12433087 Canada Inc"] * 5,
    "employee": ["Aaron Guzman"] * 5,
    "date": pd.date_range("2026-01-04", periods=5),
    "hours": [5.5, 5.5, 5.5, 5.5, 5.5],
    "hourly_rate": [17.75]*5,
    "total_pay": [97.63]*5
})

# ============================================================
# WEEK LOGIC
# ============================================================
start_date = pd.to_datetime("2026-01-04")

def assign_week(d):
    if start_date <= d <= start_date + timedelta(days=6):
        return "2026-01-10"
    return None

df["week_label"] = df["date"].apply(assign_week)
df = df[df["week_label"].notna()]


# ============================================================
# SUMMARY
# ============================================================
weekly_summary = build_weekly_summary(df)


# ============================================================
# 🔥 FORMAT 2 DECIMALES (STREAMLIT FIX REAL)
# ============================================================
def format_dataframe(df):
    numeric_cols = df.select_dtypes(include=["number"]).columns
    column_config = {}

    for col in numeric_cols:
        column_config[col] = st.column_config.NumberColumn(
            col,
            format="%.2f"
        )

    return column_config


# ============================================================
# UI
# ============================================================
st.subheader("Employee summary")

st.dataframe(
    weekly_summary,
    use_container_width=True,
    column_config=format_dataframe(weekly_summary)
)


# ============================================================
# METRICS
# ============================================================
col1, col2, col3 = st.columns(3)

col1.metric("Total Pay", f"${weekly_summary['total_pay'].sum():,.2f}")
col2.metric("Committee Hours", f"{weekly_summary['committee_hours'].sum():,.2f}")
col3.metric("REER", f"${weekly_summary['reer'].sum():,.2f}")


# ============================================================
# DOWNLOAD EXCEL
# ============================================================
def export_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)

        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]

        format_2_dec = workbook.add_format({'num_format': '0.00'})

        for col_num, col_name in enumerate(df.columns):
            worksheet.set_column(col_num, col_num, 18, format_2_dec)

    output.seek(0)
    return output


st.download_button(
    "Download Excel",
    export_excel(weekly_summary),
    "report.xlsx"
)
