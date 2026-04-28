import base64
import os
from io import BytesIO
from datetime import datetime, timedelta
import re

import pandas as pd
import streamlit as st
import requests
import msal
import plotly.graph_objects as go

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# ============================================================
# CONFIG - STREAMLIT SECRETS
# ============================================================
CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
TENANT_ID = st.secrets["TENANT_ID"]
REDIRECT_URI = st.secrets.get("REDIRECT_URI", "").strip().rstrip("/")

# NEW SHAREPOINT FOLDER LINK
# You can keep this directly here, or move it to Streamlit secrets as ONEDRIVE_SHARED_URL.
ONEDRIVE_SHARED_URL = st.secrets.get(
    "ONEDRIVE_SHARED_URL",
    "https://groupcastillo.sharepoint.com/:f:/s/GroupCastilloTeamSite/IgDJ46w1V3YWT7e0yB8CKkD9AenZh0xzbn8pNRRGuDcIpPw?e=s4L0Z9"
)

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.Read.All"]

st.set_page_config(page_title="CNET Regular Hours Report", layout="wide")

# ============================================================
# EXCEL STRUCTURE
# ============================================================
SHEET_DATA = "DATA"

# Excel columns by position, zero-based index
COL_VENDOR_COMPANY = 0      # A
COL_EMPLOYEE_NAME = 1       # B
COL_EMPLOYEE_CLASS = 8      # I
COL_RATE = 19               # T

# Day columns. Based on your description: K4 = SA/Saturday, L4 = Sunday, etc.
# K:R = 8 columns if used. If your file only uses K:Q, change end to 17.
DAY_COL_START = 10          # K
DAY_COL_END_EXCLUSIVE = 18  # R included because Python excludes the end

DAY_HEADER_ROW = 3          # Excel row 4
DATE_ROW = 4                # Excel row 5
DATA_START_ROW = 5          # Excel row 6

TYPE_OF_WORK_DEFAULT = "REGULAR"
DEFAULT_CLASS_WHEN_NO_CLASS = "Class A"
FIRST_COMMITTEE_WEEK_START = pd.Timestamp("2026-01-04")

DAY_MAP = {
    "SA": "Saturday", "SAT": "Saturday", "SATURDAY": "Saturday",
    "SU": "Sunday", "SUN": "Sunday", "SUNDAY": "Sunday",
    "MO": "Monday", "MON": "Monday", "MONDAY": "Monday",
    "TU": "Tuesday", "TUE": "Tuesday", "TUESDAY": "Tuesday",
    "WE": "Wednesday", "WED": "Wednesday", "WEDNESDAY": "Wednesday",
    "TH": "Thursday", "THU": "Thursday", "THURSDAY": "Thursday",
    "FR": "Friday", "FRI": "Friday", "FRIDAY": "Friday",
}

# ============================================================
# URL PARAM HELPERS
# ============================================================
def _get_query_params() -> dict:
    try:
        qp = st.query_params
        return {k: str(qp.get(k, "")) for k in qp.keys()}
    except Exception:
        try:
            qp = st.experimental_get_query_params()
            return {k: v[0] if isinstance(v, list) and v else str(v) for k, v in qp.items()}
        except Exception:
            return {}


def _clear_query_params():
    try:
        st.query_params.clear()
    except Exception:
        try:
            st.experimental_set_query_params()
        except Exception:
            pass

# ============================================================
# GRAPH / SHAREPOINT HELPERS
# ============================================================
def make_share_id(shared_url: str) -> str:
    b = base64.b64encode(shared_url.encode("utf-8")).decode("utf-8")
    b = b.rstrip("=").replace("/", "_").replace("+", "-")
    return "u!" + b


def graph_get(url: str, access_token: str) -> requests.Response:
    return requests.get(url, headers={"Authorization": f"Bearer {access_token}"}, timeout=60)


def resolve_shared_link(access_token: str, shared_url: str) -> dict:
    share_id = make_share_id(shared_url)
    meta_url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem"
    meta = graph_get(meta_url, access_token)
    if meta.status_code != 200:
        raise RuntimeError(f"Error resolving shared link: {meta.status_code}\n{meta.text}")
    return meta.json()


def download_excel_bytes_from_shared_link(access_token: str, shared_url: str) -> bytes:
    item = resolve_shared_link(access_token, shared_url)
    drive_id = item["parentReference"]["driveId"]
    item_id = item["id"]

    # If the shared link is a folder, pick the first Excel file inside.
    if "folder" in item:
        children_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children"
        children = graph_get(children_url, access_token)
        if children.status_code != 200:
            raise RuntimeError(f"Error reading folder contents: {children.status_code}\n{children.text}")

        files = children.json().get("value", [])
        excel_files = [
            f for f in files
            if str(f.get("name", "")).lower().endswith((".xlsx", ".xlsm", ".xls"))
        ]
        if not excel_files:
            raise RuntimeError("The SharePoint folder does not contain an Excel file (.xlsx, .xlsm, .xls).")

        # Choose most recently modified Excel file.
        excel_files = sorted(excel_files, key=lambda x: x.get("lastModifiedDateTime", ""), reverse=True)
        selected = excel_files[0]
        item_id = selected["id"]
        st.info(f"Excel selected from folder: {selected.get('name')}")

    content_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    file_r = graph_get(content_url, access_token)
    if file_r.status_code != 200:
        raise RuntimeError(f"Error downloading Excel file: {file_r.status_code}\n{file_r.text}")

    return file_r.content

# ============================================================
# DATA TRANSFORMATION
# ============================================================
def clean_text(x):
    if pd.isna(x):
        return ""
    return re.sub(r"\s+", " ", str(x).replace("\u00A0", " ")).strip()


def normalize_class(x):
    txt = clean_text(x)
    if txt == "" or txt.lower() == "no class":
        return DEFAULT_CLASS_WHEN_NO_CLASS
    return txt


def parse_date_from_cell(value, fallback_year=2026):
    if pd.isna(value):
        return pd.NaT

    if isinstance(value, (pd.Timestamp, datetime)):
        return pd.Timestamp(value).normalize()

    txt = clean_text(value)
    if not txt:
        return pd.NaT

    parsed = pd.to_datetime(txt, errors="coerce")
    if pd.notna(parsed):
        return pd.Timestamp(parsed).normalize()

    # Handles values like "3 Jan", "Jan 3", "03 de enero" if pandas cannot parse them.
    replacements = {
        "enero": "january", "janvier": "january", "jan": "january",
        "febrero": "february", "fevrier": "february", "février": "february",
        "marzo": "march", "mars": "march",
        "abril": "april", "avril": "april",
        "mayo": "may", "mai": "may",
        "junio": "june", "juin": "june",
        "julio": "july", "juillet": "july",
        "agosto": "august", "aout": "august", "août": "august",
        "septiembre": "september", "septembre": "september",
        "octubre": "october", "octobre": "october",
        "noviembre": "november", "novembre": "november",
        "diciembre": "december", "decembre": "december", "décembre": "december",
    }
    low = txt.lower().replace(" de ", " ")
    for k, v in replacements.items():
        low = low.replace(k, v)
    parsed = pd.to_datetime(f"{low} {fallback_year}", errors="coerce")
    return pd.Timestamp(parsed).normalize() if pd.notna(parsed) else pd.NaT


def committee_week_start(work_date):
    d = pd.Timestamp(work_date).normalize()
    if d < FIRST_COMMITTEE_WEEK_START:
        return FIRST_COMMITTEE_WEEK_START
    days_since = (d - FIRST_COMMITTEE_WEEK_START).days
    return FIRST_COMMITTEE_WEEK_START + pd.Timedelta(days=(days_since // 7) * 7)


@st.cache_data(ttl=300, show_spinner=False)
def load_regular_hours_report(excel_bytes: bytes) -> pd.DataFrame:
    raw = pd.read_excel(BytesIO(excel_bytes), sheet_name=SHEET_DATA, header=None)

    day_headers = raw.iloc[DAY_HEADER_ROW, DAY_COL_START:DAY_COL_END_EXCLUSIVE].tolist()
    date_headers = raw.iloc[DATE_ROW, DAY_COL_START:DAY_COL_END_EXCLUSIVE].tolist()

    rows = []
    data = raw.iloc[DATA_START_ROW:].copy()

    for _, r in data.iterrows():
        vendor = clean_text(r.iloc[COL_VENDOR_COMPANY]) if len(r) > COL_VENDOR_COMPANY else ""
        employee = clean_text(r.iloc[COL_EMPLOYEE_NAME]) if len(r) > COL_EMPLOYEE_NAME else ""
        worker_class = normalize_class(r.iloc[COL_EMPLOYEE_CLASS]) if len(r) > COL_EMPLOYEE_CLASS else DEFAULT_CLASS_WHEN_NO_CLASS
        rate = pd.to_numeric(r.iloc[COL_RATE], errors="coerce") if len(r) > COL_RATE else 0

        if vendor == "" and employee == "":
            continue

        for offset, col_idx in enumerate(range(DAY_COL_START, DAY_COL_END_EXCLUSIVE)):
            hours = pd.to_numeric(r.iloc[col_idx], errors="coerce") if len(r) > col_idx else 0
            if pd.isna(hours) or float(hours) == 0:
                continue

            day_code = clean_text(day_headers[offset]).upper()
            day_name = DAY_MAP.get(day_code, day_code)
            work_date = parse_date_from_cell(date_headers[offset])

            rows.append({
                "Vendor Company": vendor,
                "Employee Name": employee,
                "Type of Work": TYPE_OF_WORK_DEFAULT,
                "Employee Class": worker_class,
                "Work Date": work_date,
                "Day": day_name,
                "Hours": float(hours),
                "Rate": float(rate) if pd.notna(rate) else 0.0,
                "Gross Amount": float(hours) * (float(rate) if pd.notna(rate) else 0.0),
                "Committee Week Start": committee_week_start(work_date) if pd.notna(work_date) else pd.NaT,
            })

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    df["Work Date"] = pd.to_datetime(df["Work Date"], errors="coerce")
    df["Committee Week Start"] = pd.to_datetime(df["Committee Week Start"], errors="coerce")
    df["Week Label"] = df["Committee Week Start"].dt.strftime("%Y-%m-%d")
    return df


def sanitize_for_arrow(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.columns:
        if out[col].dtype == "object":
            out[col] = out[col].astype(str)
    return out

# ============================================================
# PDF REPORT
# ============================================================
def build_pdf_report(df, total_hours, total_gross, report_vendor, report_week):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    width, height = letter
    left = 40
    right = width - 40
    y = height - 45

    def money(x):
        return f"${float(x):,.2f}"

    c.setFont("Helvetica-Bold", 18)
    c.drawString(left, y, "CNET Regular Hours Report")
    c.setFont("Helvetica", 9)
    c.drawRightString(right, y, datetime.now().strftime("%Y-%m-%d %H:%M"))
    y -= 30

    c.setFont("Helvetica-Bold", 10)
    c.drawString(left, y, f"Vendor Company: {report_vendor}")
    y -= 15
    c.drawString(left, y, f"Committee Week: {report_week}")
    y -= 25

    c.setFont("Helvetica-Bold", 11)
    c.drawString(left, y, f"Total Hours: {total_hours:,.2f}")
    c.drawRightString(right, y, f"Gross Amount: {money(total_gross)}")
    y -= 30

    c.setFont("Helvetica-Bold", 9)
    headers = ["Employee", "Class", "Hours", "Rate", "Gross"]
    x_positions = [left, 250, 335, 410, right]
    for i, h in enumerate(headers):
        if i == 4:
            c.drawRightString(x_positions[i], y, h)
        else:
            c.drawString(x_positions[i], y, h)
    y -= 12
    c.line(left, y, right, y)
    y -= 12

    emp = (
        df.groupby(["Employee Name", "Employee Class"], dropna=False)
        .agg(Hours=("Hours", "sum"), Rate=("Rate", "mean"), Gross=("Gross Amount", "sum"))
        .reset_index()
        .sort_values("Employee Name")
    )

    c.setFont("Helvetica", 8)
    for _, row in emp.iterrows():
        if y < 50:
            c.showPage()
            y = height - 50
            c.setFont("Helvetica", 8)
        c.drawString(left, y, str(row["Employee Name"])[:34])
        c.drawString(250, y, str(row["Employee Class"])[:18])
        c.drawRightString(375, y, f"{row['Hours']:,.2f}")
        c.drawRightString(450, y, money(row["Rate"]))
        c.drawRightString(right, y, money(row["Gross"]))
        y -= 13

    c.save()
    buf.seek(0)
    return buf.getvalue()

# ============================================================
# MSAL LOGIN
# ============================================================
def get_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        token_cache=None,
    )

st.title("📊 CNET Regular Hours Report")

if not REDIRECT_URI:
    st.error("REDIRECT_URI is missing in Streamlit Secrets.")
    st.stop()

app = get_msal_app()
qp = _get_query_params()

if "token_result" not in st.session_state:
    if qp.get("code"):
        result = app.acquire_token_by_authorization_code(
            code=qp["code"],
            scopes=SCOPES,
            redirect_uri=REDIRECT_URI,
        )
        if "access_token" in result:
            st.session_state.token_result = result
            _clear_query_params()
            st.rerun()
        st.error("Could not obtain access_token.")
        st.code(result)
        st.stop()

    st.warning("You are not signed in to OneDrive/SharePoint.")
    auth_url = app.get_authorization_request_url(scopes=SCOPES, redirect_uri=REDIRECT_URI)
    st.link_button("Sign in to OneDrive", auth_url)
    st.stop()

token_result = st.session_state.token_result
st.success("✅ Connected to OneDrive/SharePoint")

c1, c2 = st.columns([1, 3])
with c1:
    if st.button("🔄 Refresh data"):
        st.session_state.pop("excel_bytes", None)
        load_regular_hours_report.clear()
        st.rerun()
with c2:
    if st.button("🔒 Sign out"):
        for k in ["token_result", "excel_bytes"]:
            st.session_state.pop(k, None)
        _clear_query_params()
        st.rerun()

try:
    if "excel_bytes" not in st.session_state:
        st.info("📥 Downloading Excel from SharePoint/OneDrive...")
        st.session_state.excel_bytes = download_excel_bytes_from_shared_link(
            token_result["access_token"],
            ONEDRIVE_SHARED_URL,
        )
    excel_bytes = st.session_state.excel_bytes
except Exception as e:
    st.error("Could not download the Excel file.")
    st.code(str(e))
    st.stop()

try:
    df_all = load_regular_hours_report(excel_bytes)
except Exception as e:
    st.error("Could not read the DATA sheet or transform the file.")
    st.code(str(e))
    st.stop()

if df_all.empty:
    st.warning("No regular hours were found. Verify sheet DATA, row 4 day headers, row 5 dates, and columns K:R hours.")
    st.stop()

# ============================================================
# FILTERS
# ============================================================
st.sidebar.header("Filters")

vendors = sorted(df_all["Vendor Company"].dropna().astype(str).unique().tolist())
sel_vendors = st.sidebar.multiselect("Vendor Company", vendors, default=[])

df = df_all.copy()
if sel_vendors:
    df = df[df["Vendor Company"].astype(str).isin(sel_vendors)]

weeks = sorted(df["Week Label"].dropna().astype(str).unique().tolist())
sel_weeks = st.sidebar.multiselect("Committee Week Start", weeks, default=[])
if sel_weeks:
    df = df[df["Week Label"].astype(str).isin(sel_weeks)]

classes = sorted(df["Employee Class"].dropna().astype(str).unique().tolist())
sel_classes = st.sidebar.multiselect("Employee Class", classes, default=[])
if sel_classes:
    df = df[df["Employee Class"].astype(str).isin(sel_classes)]

employees = sorted(df["Employee Name"].dropna().astype(str).unique().tolist())
sel_employees = st.sidebar.multiselect("Employee Name", employees, default=[])
if sel_employees:
    df = df[df["Employee Name"].astype(str).isin(sel_employees)]

# ============================================================
# KPIs
# ============================================================
total_hours = float(df["Hours"].fillna(0).sum())
total_gross = float(df["Gross Amount"].fillna(0).sum())
avg_rate = total_gross / total_hours if total_hours else 0
employee_count = df["Employee Name"].nunique()

st.subheader("📌 Executive Summary")
k1, k2, k3, k4 = st.columns(4)
k1.metric("Total Hours", f"{total_hours:,.2f}")
k2.metric("Gross Amount", f"${total_gross:,.2f}")
k3.metric("Average Rate", f"${avg_rate:,.2f}")
k4.metric("Employees", f"{employee_count:,}")

# ============================================================
# CHARTS
# ============================================================
st.subheader("📊 Hours and Gross Amount")

by_vendor = (
    df.groupby("Vendor Company", dropna=False)
    .agg(Hours=("Hours", "sum"), Gross=("Gross Amount", "sum"))
    .reset_index()
    .sort_values("Hours", ascending=False)
)

fig_vendor = go.Figure()
fig_vendor.add_trace(go.Bar(name="Hours", x=by_vendor["Vendor Company"], y=by_vendor["Hours"]))
fig_vendor.update_layout(title="Hours by Vendor Company", xaxis_title="Vendor Company", yaxis_title="Hours")
st.plotly_chart(fig_vendor, use_container_width=True)

by_day = (
    df.groupby(["Work Date", "Day"], dropna=False)
    .agg(Hours=("Hours", "sum"), Gross=("Gross Amount", "sum"))
    .reset_index()
    .sort_values("Work Date")
)
by_day["Date Label"] = by_day["Work Date"].dt.strftime("%Y-%m-%d")

fig_day = go.Figure()
fig_day.add_trace(go.Bar(name="Hours", x=by_day["Date Label"], y=by_day["Hours"]))
fig_day.add_trace(go.Scatter(name="Gross", x=by_day["Date Label"], y=by_day["Gross"], mode="lines+markers"))
fig_day.update_layout(title="Daily Hours and Gross Amount", xaxis_title="Date", yaxis_title="Value")
st.plotly_chart(fig_day, use_container_width=True)

# ============================================================
# EMPLOYEE SUMMARY
# ============================================================
st.subheader("🧾 Employee Summary")

emp_summary = (
    df.groupby(["Vendor Company", "Employee Name", "Employee Class", "Rate"], dropna=False)
    .agg(Hours=("Hours", "sum"), Gross=("Gross Amount", "sum"))
    .reset_index()
    .sort_values(["Vendor Company", "Employee Name"])
)

emp_show = emp_summary.copy()
emp_show["Rate"] = emp_show["Rate"].map(lambda x: f"${float(x):,.2f}")
emp_show["Gross"] = emp_show["Gross"].map(lambda x: f"${float(x):,.2f}")
st.dataframe(emp_show, use_container_width=True)

# ============================================================
# DETAILS
# ============================================================
with st.expander("Daily details"):
    details = df.copy().sort_values(["Work Date", "Vendor Company", "Employee Name"])
    details["Work Date"] = details["Work Date"].dt.strftime("%Y-%m-%d")
    details["Committee Week Start"] = details["Committee Week Start"].dt.strftime("%Y-%m-%d")
    details["Rate"] = details["Rate"].map(lambda x: f"${float(x):,.2f}")
    details["Gross Amount"] = details["Gross Amount"].map(lambda x: f"${float(x):,.2f}")
    st.dataframe(sanitize_for_arrow(details), use_container_width=True)

# ============================================================
# EXPORTS
# ============================================================
st.divider()
st.subheader("📤 Export")

report_vendor = ", ".join(sel_vendors) if sel_vendors else "All Vendors"
report_week = ", ".join(sel_weeks) if sel_weeks else "All Weeks"

csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "⬇️ Download CSV",
    data=csv_bytes,
    file_name="CNET_Regular_Hours_Report.csv",
    mime="text/csv",
)

pdf_bytes = build_pdf_report(df, total_hours, total_gross, report_vendor, report_week)
st.download_button(
    "⬇️ Download PDF",
    data=pdf_bytes,
    file_name="CNET_Regular_Hours_Report.pdf",
    mime="application/pdf",
)
