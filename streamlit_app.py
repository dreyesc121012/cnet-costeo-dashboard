import base64
import os
from io import BytesIO
from datetime import datetime
import re

import pandas as pd
import streamlit as st
import requests
import msal
import plotly.graph_objects as go

# PDF
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# ============================================================
# CONFIG (Secrets)
# ============================================================
CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
ONEDRIVE_SHARED_URL = st.secrets["ONEDRIVE_SHARED_URL"]

REDIRECT_URI = st.secrets.get("REDIRECT_URI", "").strip().rstrip("/")

TENANT_ID = st.secrets["TENANT_ID"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

SCOPES = ["User.Read", "Files.Read.All"]

# Excel config
SHEET_REAL = "Real Master"
SHEET_FIXED_124 = "Gasto Fijo"
SHEET_FIXED_9359 = "Gasto Fijo 9359"
HEADER_IDX = 6

# Exact column names
MONTH_COL = "Month"
YEAR_COL = "Year"

st.set_page_config(page_title="CNET Costing Dashboard", layout="wide")

# ============================================================
# HELPERS (Streamlit URL params)
# ============================================================
def _get_query_params() -> dict:
    try:
        qp = st.query_params
        out = {}
        for k in qp.keys():
            v = qp.get(k)
            if isinstance(v, list):
                out[k] = v[0] if v else ""
            else:
                out[k] = str(v) if v is not None else ""
        return out
    except Exception:
        pass

    try:
        qp = st.experimental_get_query_params()
        return {k: (v[0] if isinstance(v, list) and v else str(v)) for k, v in qp.items()}
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
# HELPERS (Graph download)
# ============================================================
def make_share_id(shared_url: str) -> str:
    b = base64.b64encode(shared_url.encode("utf-8")).decode("utf-8")
    b = b.rstrip("=").replace("/", "_").replace("+", "-")
    return "u!" + b


def graph_get(url: str, access_token: str) -> requests.Response:
    return requests.get(
        url,
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=60
    )


def download_excel_bytes_from_shared_link(access_token: str, shared_url: str) -> bytes:
    share_id = make_share_id(shared_url)

    meta_url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem"
    meta = graph_get(meta_url, access_token)
    if meta.status_code != 200:
        raise RuntimeError(
            f"Error resolving shared link: {meta.status_code}\n{meta.text}\n\n"
            f"TIP: Create a NEW link (Share -> Copy link) and replace ONEDRIVE_SHARED_URL."
        )

    meta_json = meta.json()
    item_id = meta_json["id"]
    drive_id = meta_json["parentReference"]["driveId"]

    content_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    file_r = graph_get(content_url, access_token)
    if file_r.status_code != 200:
        raise RuntimeError(f"Error downloading file: {file_r.status_code}\n{file_r.text}")

    return file_r.content


# ============================================================
# HELPERS (Excel parsing)
# ============================================================
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


@st.cache_data(ttl=300, show_spinner=False)
def read_real_master_from_bytes(excel_bytes: bytes) -> pd.DataFrame:
    raw = pd.read_excel(BytesIO(excel_bytes), sheet_name=SHEET_REAL, header=None)
    headers = make_unique_columns(raw.iloc[HEADER_IDX].tolist())

    df = raw.iloc[HEADER_IDX + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _load_fixed_total_by_sheet(excel_bytes: bytes, sheet_name: str) -> float:
    try:
        fx = pd.read_excel(BytesIO(excel_bytes), sheet_name=sheet_name, header=None)
    except Exception:
        return 0.0

    best_col = None
    best_count = -1

    for j in range(fx.shape[1]):
        s = pd.to_numeric(fx.iloc[:, j], errors="coerce")
        cnt = int(s.notna().sum())
        if cnt > best_count:
            best_count = cnt
            best_col = j

    if best_col is None or best_count <= 0:
        return 0.0

    amounts = pd.to_numeric(fx.iloc[:, best_col], errors="coerce")
    return float(amounts.fillna(0).sum())


def _load_fixed_breakdown_by_sheet(excel_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    try:
        fx = pd.read_excel(BytesIO(excel_bytes), sheet_name=sheet_name, header=None)
    except Exception:
        return pd.DataFrame(columns=["Category", "Amount"])

    best_col = None
    best_count = -1
    for j in range(fx.shape[1]):
        s = pd.to_numeric(fx.iloc[:, j], errors="coerce")
        cnt = int(s.notna().sum())
        if cnt > best_count:
            best_count = cnt
            best_col = j

    if best_col is None or best_count <= 0:
        return pd.DataFrame(columns=["Category", "Amount"])

    cat_col = best_col - 1 if best_col - 1 >= 0 else 0

    df_fixed = fx.iloc[:, [cat_col, best_col]].copy()
    df_fixed.columns = ["Category", "Amount"]

    df_fixed["Category"] = df_fixed["Category"].astype(str).str.strip()
    df_fixed["Amount"] = pd.to_numeric(df_fixed["Amount"], errors="coerce")

    df_fixed = df_fixed.dropna(subset=["Amount"])
    df_fixed = df_fixed[df_fixed["Category"].str.lower().ne("nan")]
    df_fixed = df_fixed[df_fixed["Category"] != ""]
    df_fixed = df_fixed[~df_fixed["Category"].str.lower().str.contains("gasto", na=False)]
    df_fixed = df_fixed[~df_fixed["Category"].str.lower().str.contains("fixed", na=False)]

    df_fixed = (
        df_fixed.groupby("Category", as_index=False)["Amount"]
        .sum()
        .sort_values("Amount", ascending=False)
        .reset_index(drop=True)
    )
    return df_fixed


@st.cache_data(ttl=300, show_spinner=False)
def load_fixed_data_from_bytes(excel_bytes: bytes):
    return {
        "12433087 Canada Inc": {
            "sheet": SHEET_FIXED_124,
            "total": _load_fixed_total_by_sheet(excel_bytes, SHEET_FIXED_124),
            "breakdown": _load_fixed_breakdown_by_sheet(excel_bytes, SHEET_FIXED_124),
        },
        "9359-6633 Quebec Inc": {
            "sheet": SHEET_FIXED_9359,
            "total": _load_fixed_total_by_sheet(excel_bytes, SHEET_FIXED_9359),
            "breakdown": _load_fixed_breakdown_by_sheet(excel_bytes, SHEET_FIXED_9359),
        },
    }


def _norm(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s.lower()


def find_col(df: pd.DataFrame, name: str):
    target = _norm(name)

    for c in df.columns:
        if _norm(c) == target:
            return c

    for c in df.columns:
        if target in _norm(c):
            return c

    return None


def safe_pct(x: float, base: float) -> float:
    return (x / base) if base not in (0, None) else 0.0


def sanitize_for_arrow(df: pd.DataFrame) -> pd.DataFrame:
    df2 = df.copy()
    for col in df2.columns:
        if df2[col].dtype == "object":
            df2[col] = df2[col].astype(str)
    return df2


def pick_building_col(df: pd.DataFrame):
    candidates = ["Building", "Building ID", "BuildinG ID", "Building Name", "Site", "Location", "Branch"]
    for name in candidates:
        c = find_col(df, name)
        if c:
            return c
    for c in df.columns:
        if "build" in _norm(c):
            return c
    return None


# ============================================================
# MONTH + YEAR -> TEXT LABELS (NO TIME AXIS)
# ============================================================
_MONTH_NUM = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}


def build_month_fields(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    out["_YearInt"] = pd.to_numeric(out[YEAR_COL], errors="coerce").astype("Int64")

    m = out[MONTH_COL].astype(str).str.strip().str.lower()
    m = m.str.replace(r"[^a-z]", "", regex=True)
    out["_MonthNum"] = m.map(_MONTH_NUM).astype("Int64")

    bad = out["_YearInt"].isna() | out["_MonthNum"].isna()
    out["_MonthKey"] = pd.NA
    out["_MonthText"] = pd.NA

    ok = ~bad
    out.loc[ok, "_MonthKey"] = (
        out.loc[ok, "_YearInt"].astype(int).astype(str)
        + "-"
        + out.loc[ok, "_MonthNum"].astype(int).astype(str).str.zfill(2)
    )
    out.loc[ok, "_MonthText"] = (
        out.loc[ok, MONTH_COL].astype(str).str.strip()
        + " "
        + out.loc[ok, "_YearInt"].astype(int).astype(str)
    )
    return out


# ============================================================
# FILTERS
# ============================================================
def add_filters(df: pd.DataFrame):
    st.sidebar.header("Executive Filters")

    filter_state = {
        "years": [],
        "months": [],
        "companies": [],
        "provinces": [],
        "clients": [],
        "projects": [],
        "buildings": [],
    }

    if MONTH_COL in df.columns and YEAR_COL in df.columns:
        df = build_month_fields(df)

        years = sorted([int(y) for y in df["_YearInt"].dropna().unique().tolist()])
        sel_years = st.sidebar.multiselect("Year", years, default=[])
        filter_state["years"] = sel_years
        if sel_years:
            df = df[df["_YearInt"].isin(sel_years)]

        month_table = (
            df[["_MonthKey", "_MonthText"]]
            .dropna()
            .drop_duplicates()
            .sort_values("_MonthKey")
        )
        sel_months = st.sidebar.multiselect("Month", month_table["_MonthText"].tolist(), default=[])
        filter_state["months"] = sel_months
        if sel_months:
            df = df[df["_MonthText"].isin(sel_months)]
    else:
        st.sidebar.warning("Month/Year columns were not found in the Excel file.")

    c_company = find_col(df, "Company")
    if c_company:
        sel = st.sidebar.multiselect("Company", sorted(df[c_company].dropna().astype(str).unique().tolist()))
        filter_state["companies"] = sel
        if sel:
            df = df[df[c_company].astype(str).isin(sel)]

    c_prov = find_col(df, "Province")
    if c_prov:
        sel = st.sidebar.multiselect("Province", sorted(df[c_prov].dropna().astype(str).unique().tolist()))
        filter_state["provinces"] = sel
        if sel:
            df = df[df[c_prov].astype(str).isin(sel)]

    c_client = find_col(df, "Client")
    if c_client:
        sel = st.sidebar.multiselect("Client", sorted(df[c_client].dropna().astype(str).unique().tolist()))
        filter_state["clients"] = sel
        if sel:
            df = df[df[c_client].astype(str).isin(sel)]

    c_proj = find_col(df, "Project Name")
    if c_proj:
        sel = st.sidebar.multiselect("Project (Project Name)", sorted(df[c_proj].dropna().astype(str).unique().tolist()))
        filter_state["projects"] = sel
        if sel:
            df = df[df[c_proj].astype(str).isin(sel)]

    c_bld = pick_building_col(df)
    if c_bld:
        uniq = sorted(df[c_bld].dropna().astype(str).unique().tolist())
        sel = st.sidebar.multiselect("Building", uniq)
        filter_state["buildings"] = sel
        if sel:
            df = df[df[c_bld].astype(str).isin(sel)]

    return df, filter_state


# ============================================================
# PDF REPORT - EXECUTIVE VERSION WITH LOGO
# ============================================================
def build_pdf_report(
    *,
    income, cost, gross,
    fixed_total_applied, fixed_sheet_name,
    net, mgmt_fee_total, royalty_5_total, royalty_3_total, new_total,
    p_cost, p_gross, p_fixed, p_net, p_mgmt, p_roy5, p_roy3, p_new,
    target, gross_margin, net_margin, final_margin,
    report_company="All Companies",
    report_building="All Buildings",
    report_month="All Periods",
    fig_waterfall=None, fig_gauge=None,
):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    width, height = letter

    # Layout
    left = 40
    right = width - 40
    top = height - 40
    bottom = 40
    page_num = 1

    def money(x):
        try:
            return f"${float(x):,.2f}"
        except Exception:
            return "$0.00"

    def pct(x):
        try:
            return f"{float(x):.1%}"
        except Exception:
            return "0.0%"

    def safe_text(value, default_text):
        txt = str(value).strip() if value is not None else ""
        return txt if txt else default_text

    def truncate_text(text, max_len=95):
        text = safe_text(text, "")
        return text if len(text) <= max_len else text[:max_len - 3] + "..."

    def footer():
        c.setStrokeColorRGB(0.8, 0.8, 0.8)
        c.line(left, bottom + 15, right, bottom + 15)
        c.setFont("Helvetica", 8)
        c.setFillColorRGB(0.4, 0.4, 0.4)
        c.drawString(left, bottom, "CNET Building Maintenance Services")
        c.drawRightString(right, bottom, f"Page {page_num}")

    def new_page():
        nonlocal page_num
        footer()
        c.showPage()
        page_num += 1
        return header()

    def space(y, needed=50):
        if y < bottom + needed:
            return new_page()
        return y

    def header():
        y = top

        try:
            logo_candidates = [
                "cnet_logo.png",
                os.path.join(os.getcwd(), "cnet_logo.png"),
                "/mount/src/work-orders/cnet_logo.png",
            ]
            for logo_path in logo_candidates:
                if os.path.exists(logo_path):
                    logo = ImageReader(logo_path)
                    c.drawImage(
                        logo,
                        left,
                        y - 40,
                        width=120,
                        height=40,
                        preserveAspectRatio=True,
                        mask='auto'
                    )
                    break
        except Exception:
            pass

        # Title
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica-Bold", 18)
        c.drawRightString(right, y, "Executive Summary")

        # Generated date
        c.setFont("Helvetica", 9)
        c.setFillColorRGB(0.4, 0.4, 0.4)
        c.drawRightString(right, y - 15, datetime.now().strftime("%Y-%m-%d %H:%M"))

        # Report filters
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica-Bold", 9)

        company_text = truncate_text(report_company, 90)
        building_text = truncate_text(report_building, 90)
        month_text = truncate_text(report_month, 90)

        c.drawString(left, y - 58, f"Company: {company_text}")
        c.drawString(left, y - 72, f"Building: {building_text}")
        c.drawString(left, y - 86, f"Report Period: {month_text}")

        c.setStrokeColorRGB(0.7, 0.7, 0.7)
        c.line(left, y - 98, right, y - 98)

        return y - 118

    y = header()

    # Executive Overview
    y = space(y, 90)
    c.setFillColorRGB(0.97, 0.97, 0.97)
    c.roundRect(left, y - 65, right - left, 65, 6, fill=1, stroke=0)

    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(left + 10, y - 15, "Executive Overview")

    c.setFont("Helvetica", 10)
    c.drawString(left + 10, y - 35, f"Revenue: {money(income)} | Final Margin: {pct(final_margin)}")
    c.drawString(left + 10, y - 50, f"Net: {money(net)} | Final Total: {money(new_total)}")

    y -= 90

    # KPI table
    rows = [
        ("Revenue", income, 1.0),
        ("Direct Costs", cost, p_cost),
        ("Gross Profit", gross, p_gross),
    ]

    if fixed_total_applied != 0:
        rows += [
            ("Fixed Expenses", fixed_total_applied, p_fixed),
            ("Net Profit", net, p_net),
        ]
    else:
        rows += [("Net Profit", net, p_net)]

    rows += [
        ("Management Fee", mgmt_fee_total, p_mgmt),
        ("Royalty 5%", royalty_5_total, p_roy5),
    ]

    if royalty_3_total != 0:
        rows += [("Royalty 3%", royalty_3_total, p_roy3)]

    rows += [("Final Total", new_total, p_new)]

    y = space(y, 140)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(left, y, "Key Performance Indicators")
    y -= 18

    c.setFont("Helvetica-Bold", 10)
    c.drawString(left, y, "Concept")
    c.drawRightString(380, y, "Amount")
    c.drawRightString(right, y, "%")
    y -= 14

    c.setFont("Helvetica", 10)

    for label, val, share in rows:
        y = space(y, 24)
        c.drawString(left, y, str(label))
        c.drawRightString(380, y, money(val))
        c.drawRightString(right, y, pct(share))
        y -= 16

    # Margins
    y -= 8
    y = space(y, 70)

    c.setFillColorRGB(0.97, 0.97, 0.97)
    c.roundRect(left, y - 52, right - left, 52, 6, fill=1, stroke=0)

    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left + 10, y - 16, "Margins")

    c.setFont("Helvetica", 10)
    c.drawString(
        left + 10,
        y - 36,
        f"Gross: {pct(gross_margin)} | Net: {pct(net_margin)} | Final: {pct(final_margin)}"
    )

    y -= 75

    # Charts
    def add_chart(fig, title, y):
        if fig is None:
            return y

        y = space(y, 250)

        c.setFont("Helvetica-Bold", 10)
        c.setFillColorRGB(0, 0, 0)
        c.drawString(left, y, title)
        y -= 10

        try:
            img = ImageReader(BytesIO(fig.to_image(format="png")))
            c.drawImage(
                img,
                left,
                y - 200,
                width=520,
                height=200,
                preserveAspectRatio=True,
                mask='auto'
            )
            y -= 220
        except Exception:
            c.setFont("Helvetica", 9)
            c.drawString(left, y, "Chart not available (install kaleido)")
            y -= 20

        return y

    y = add_chart(fig_gauge, "Margin Gauge", y)
    y = add_chart(fig_waterfall, "Financial Waterfall", y)

    footer()
    c.save()
    buf.seek(0)
    return buf.getvalue()


# ============================================================
# MSAL APP
# ============================================================
def get_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        token_cache=None,
    )


# ============================================================
# UI + LOGIN
# ============================================================
st.title("📊 CNET Financial Performance & Budget Control")

if not REDIRECT_URI:
    st.error("REDIRECT_URI is missing in Secrets. Example: https://cnet-dashboard.streamlit.app (no trailing slash).")
    st.stop()

app = get_msal_app()
qp = _get_query_params()

if "token_result" not in st.session_state:
    if qp.get("code"):
        try:
            result = app.acquire_token_by_authorization_code(
                code=qp["code"],
                scopes=SCOPES,
                redirect_uri=REDIRECT_URI,
            )
        except Exception as e:
            st.error(f"Login error: {e}")
            st.stop()

        if "access_token" in result:
            st.session_state.token_result = result
            _clear_query_params()
            st.rerun()
        else:
            st.error("Could not obtain access_token.")
            st.code(result)
            st.stop()

    st.warning("You are not signed in to OneDrive/SharePoint.")
    auth_url = app.get_authorization_request_url(scopes=SCOPES, redirect_uri=REDIRECT_URI)
    st.markdown("### 🔐 Sign in")
    st.link_button("Sign in to OneDrive", auth_url)
    st.caption(f"Auth URL (should contain /{TENANT_ID}/): {auth_url}")
    st.stop()

token_result = st.session_state.token_result

if "access_token" not in token_result:
    st.error("Could not obtain a valid token.")
    st.code(token_result)
    st.stop()

st.success("✅ Connected to OneDrive/SharePoint (active token)")

colA, colB = st.columns([1, 3])
with colA:
    if st.button("🔄 Refresh data"):
        st.session_state.pop("excel_bytes", None)
        read_real_master_from_bytes.clear()
        load_fixed_data_from_bytes.clear()
        st.rerun()

with colB:
    if st.button("🔒 Sign out"):
        for k in ["token_result", "excel_bytes"]:
            st.session_state.pop(k, None)
        _clear_query_params()
        st.rerun()


# ============================================================
# Download + Load
# ============================================================
try:
    if "excel_bytes" not in st.session_state:
        st.info("📥 Downloading Excel from SharePoint/OneDrive…")
        st.session_state.excel_bytes = download_excel_bytes_from_shared_link(
            token_result["access_token"],
            ONEDRIVE_SHARED_URL
        )
    excel_bytes = st.session_state.excel_bytes
except Exception as e:
    st.error("Could not download the file from OneDrive/SharePoint.")
    st.code(str(e))
    st.stop()

df_all = read_real_master_from_bytes(excel_bytes)
fixed_data = load_fixed_data_from_bytes(excel_bytes)

df, filter_state = add_filters(df_all.copy())


# ============================================================
# KPI base
# ============================================================
COL_INCOME = "Total to Bill"
COL_COST_REAL = "Total Cost Real"
COL_COST_BUDGET = "Total Cost Budget"
COL_COST_VAR = "Variation Total Cost (Budget vs Real)"
COL_MGMT = "Total Management Fee"
COL_ROY_5 = "Royalty CNET Group Inc 5%"
COL_ROY_3 = "Royalty CNET Master 3% BGIS"


# ============================================================
# CATEGORY SPECS
# ============================================================
CATEGORY_SPECS = {
    "Labor": {
        "real": "Total Labor Real",
        "budget": "Total Labor Budget",
        "var": "Variation Labor  (Budget vs Real)",
    },
    "PW": {
        "real": "Total PW Real",
        "budget": "Total PW Budget",
        "var": "Variation PW  (Budget vs Real)",
    },
    "Supplies": {
        "real": "Total Supplies Real",
        "budget": "Total Supplies Budget",
        "var": "Variation Supplies  (Budget vs Real)",
    },
    "Equipment": {
        "real": "Total Equipment Real",
        "budget": "Total Equipment Budget",
        "var": "Variation Equipment (Budget vs Real)",
    },
    "Total Cost": {
        "real": "Total Cost Real",
        "budget": "Total Cost Budget",
        "var": "Variation Total Cost (Budget vs Real)",
    },
}

# Resolve columns
c_income = find_col(df, COL_INCOME)
c_cost = find_col(df, COL_COST_REAL)
c_mgmt = find_col(df, COL_MGMT)
c_roy5 = find_col(df, COL_ROY_5)
c_roy3 = find_col(df, COL_ROY_3)

c_company = find_col(df, "Company")
c_client = find_col(df, "Client")

missing = [k for k, v in {
    COL_INCOME: c_income,
    COL_COST_REAL: c_cost,
    COL_MGMT: c_mgmt,
    COL_ROY_5: c_roy5,
}.items() if v is None]

if missing:
    st.error(f"Missing columns in 'Real Master': {missing}")
    with st.expander("Show detected columns"):
        st.write(df.columns.tolist())
    st.stop()

# Numeric conversion
for c in [c_income, c_cost, c_mgmt, c_roy5]:
    df[c] = pd.to_numeric(df[c], errors="coerce")
if c_roy3:
    df[c_roy3] = pd.to_numeric(df[c_roy3], errors="coerce")


# ============================================================
# FINAL BUSINESS RULES
# ============================================================
COMPANY_FIXED_124 = "12433087 Canada Inc"
COMPANY_FIXED_9359 = "9359-6633 Quebec Inc"
COMPANY_BG_QC = "9359-6633 Quebec Inc"
CLIENT_BGIS = "BGIS"

income = float(df[c_income].fillna(0).sum())
cost = float(df[c_cost].fillna(0).sum())
gross = income - cost

mgmt_fee_total = float(df[c_mgmt].fillna(0).sum())
royalty_5_total = float(df[c_roy5].fillna(0).sum())
royalty_3_total = float(df[c_roy3].fillna(0).sum()) if c_roy3 else 0.0

selected_companies_explicit = [str(x).strip() for x in filter_state.get("companies", [])]
selected_projects_explicit = [str(x).strip() for x in filter_state.get("projects", [])]

project_filter_active = len(selected_projects_explicit) > 0
company_filter_active = len(selected_companies_explicit) > 0

selected_company = selected_companies_explicit[0] if len(selected_companies_explicit) == 1 else None

apply_fixed = (
    company_filter_active
    and not project_filter_active
    and len(selected_companies_explicit) == 1
    and selected_company in {COMPANY_FIXED_124, COMPANY_FIXED_9359}
)

fixed_total = 0.0
fixed_sheet_name = ""
df_fixed_breakdown = pd.DataFrame(columns=["Category", "Amount"])

if apply_fixed and selected_company in fixed_data:
    fixed_total = float(fixed_data[selected_company]["total"])
    fixed_sheet_name = str(fixed_data[selected_company]["sheet"])
    df_fixed_breakdown = fixed_data[selected_company]["breakdown"].copy()

net = gross - fixed_total

apply_roy3 = False
if c_company and c_client:
    companies = df[c_company].dropna().astype(str).str.strip().unique().tolist()
    clients = df[c_client].dropna().astype(str).str.strip().unique().tolist()
    apply_roy3 = (
        len(companies) == 1
        and companies[0] == COMPANY_BG_QC
        and len(clients) == 1
        and clients[0] == CLIENT_BGIS
    )

if apply_roy3:
    new_total = net + mgmt_fee_total + royalty_5_total + royalty_3_total
else:
    new_total = net + mgmt_fee_total + royalty_5_total

# % of revenue
p_cost = safe_pct(cost, income)
p_gross = safe_pct(gross, income)
p_fixed = safe_pct(fixed_total, income)
p_net = safe_pct(net, income)
p_mgmt = safe_pct(mgmt_fee_total, income)
p_roy5 = safe_pct(royalty_5_total, income)
p_roy3 = safe_pct(royalty_3_total, income) if apply_roy3 else 0.0
p_new = safe_pct(new_total, income)


# ============================================================
# Traffic Light + Gauge
# ============================================================
st.subheader("📌 Executive Margin (KPIs + Traffic Light)")

target = st.slider("Target margin (%)", 0, 60, 25) / 100
yellow_zone = 0.05

gross_margin = gross / income if income else 0
net_margin = net / income if income else 0
final_margin = new_total / income if income else 0

def traffic_light(m, tgt):
    if m >= tgt + yellow_zone:
        return "🟢"
    elif m >= tgt - yellow_zone:
        return "🟡"
    return "🔴"


c1, c2 = st.columns(2)

c1.metric(
    "Gross Margin",
    f"{gross_margin:.1%}",
    f"{traffic_light(gross_margin, target)} vs {target:.0%}"
)

c2.metric(
    "Net Margin",
    f"{net_margin:.1%}",
    f"{traffic_light(net_margin, target)} vs {target:.0%}"
)

st.caption("Gauge: net margin")
gauge_max = 60
fig_gauge = go.Figure(go.Indicator(
    mode="gauge+number",
    value=float(net_margin * 100),
    number={"suffix": "%"},
    gauge={
        "axis": {"range": [0, gauge_max]},
        "threshold": {"line": {"width": 4}, "value": float(target * 100)},
        "steps": [
            {"range": [0, max(0, (target - yellow_zone) * 100)]},
            {"range": [max(0, (target - yellow_zone) * 100), (target + yellow_zone) * 100]},
            {"range": [(target + yellow_zone) * 100, gauge_max]},
        ],
    }
))
st.plotly_chart(fig_gauge, use_container_width=True)


# ============================================================
# KPI CARDS
# ============================================================
st.subheader("📊 KPIs (Executive)")

left_col, right_col = st.columns([2, 1])

with left_col:
    l1, l2 = st.columns(2)
    l1.metric("Revenue (Total to Bill)", f"${income:,.2f}")
    l2.metric("Costs (Total Cost Real)", f"${cost:,.2f}", f"{p_cost*100:,.2f}%")

    l3, l4 = st.columns(2)
    l3.metric("Gross (Revenue - Cost)", f"${gross:,.2f}", f"{p_gross*100:,.2f}%")

    if apply_fixed:
        l4.metric("Fixed Expenses (Gasto Fijo)", f"${fixed_total:,.2f}", f"{p_fixed*100:,.2f}%")
    else:
        l4.metric("Net", f"${net:,.2f}", f"{p_net*100:,.2f}%")

    if apply_fixed:
        l5, l6 = st.columns(2)
        l5.metric("Net (Gross - Fixed)", f"${net:,.2f}", f"{p_net*100:,.2f}%")
        l6.empty()

with right_col:
    st.metric("Total Management Fee", f"${mgmt_fee_total:,.2f}", f"{p_mgmt*100:,.2f}%")
    st.metric("Royalty (5%)", f"${royalty_5_total:,.2f}", f"{p_roy5*100:,.2f}%")
    st.metric(
        "Royalty (3%) BGIS",
        f"${royalty_3_total:,.2f}" if apply_roy3 else "$0.00",
        f"{p_roy3*100:,.2f}%"
    )


# ============================================================
# OPTIONAL KPI: Total Cost Budget vs Real + % used / variance
# ============================================================
tc_r = find_col(df, COL_COST_REAL)
tc_b = find_col(df, COL_COST_BUDGET)
tc_v = find_col(df, COL_COST_VAR)

if tc_r and tc_b and tc_v:
    df[tc_r] = pd.to_numeric(df[tc_r], errors="coerce")
    df[tc_b] = pd.to_numeric(df[tc_b], errors="coerce")
    df[tc_v] = pd.to_numeric(df[tc_v], errors="coerce")

    total_cost_real = float(df[tc_r].fillna(0).sum())
    total_cost_budget = float(df[tc_b].fillna(0).sum())
    total_cost_var = float(df[tc_v].fillna(0).sum())

    pct_used = safe_pct(total_cost_real, total_cost_budget)
    pct_under = safe_pct(max(0.0, total_cost_var), total_cost_budget)
    pct_over = safe_pct(max(0.0, total_cost_real - total_cost_budget), total_cost_budget)

    status_tc = "🟢 On track"
    if total_cost_var < 0:
        status_tc = "🔴 Over budget"
    elif total_cost_var > 0:
        status_tc = "🟢 Under budget"

    t1, t2, t3, t4 = st.columns(4)
    t1.metric("Total Cost Real", f"${total_cost_real:,.2f}")
    t2.metric("Total Cost Budget", f"${total_cost_budget:,.2f}")
    t3.metric("Variation (Budget - Real)", f"${total_cost_var:,.2f}", status_tc)
    t4.metric("% Budget Used", f"{pct_used*100:,.1f}%", f"Over: {pct_over*100:,.1f}% | Under: {pct_under*100:,.1f}%")


# ============================================================
# WATERFALL
# ============================================================
st.subheader("📉 Executive Waterfall")

wf_x = ["Revenue", "Costs", "Gross"]
wf_y = [income, -cost, gross]
wf_measure = ["absolute", "relative", "relative"]

if apply_fixed:
    wf_x += ["Fixed"]
    wf_y += [-fixed_total]
    wf_measure += ["relative"]

wf_x += ["Mgmt Fee", "Royalty 5%"]
wf_y += [mgmt_fee_total, royalty_5_total]
wf_measure += ["relative", "relative"]

if apply_roy3:
    wf_x += ["Royalty 3%"]
    wf_y += [royalty_3_total]
    wf_measure += ["relative"]

wf_x += ["New Total"]
wf_y += [new_total]
wf_measure += ["total"]

fig_waterfall = go.Figure(go.Waterfall(
    orientation="v",
    measure=wf_measure,
    x=wf_x,
    y=wf_y,
))
fig_waterfall.update_layout(title="Waterfall: Revenue → Costs → (Fixed) → Fees → New Total", showlegend=False)
st.plotly_chart(fig_waterfall, use_container_width=True)


# ============================================================
# CATEGORY BUDGET vs REAL BREAKDOWN
# ============================================================
st.subheader("🧩 Budget vs Real Breakdown (Categories)")

color_red = "#d93025"
color_green = "#188038"
color_gray = "#5f6368"

rows = []

for cat, spec in CATEGORY_SPECS.items():
    c_real = find_col(df, spec["real"])
    c_budget = find_col(df, spec["budget"])
    c_var = find_col(df, spec["var"])

    if not all([c_real, c_budget, c_var]):
        continue

    df[c_real] = pd.to_numeric(df[c_real], errors="coerce")
    df[c_budget] = pd.to_numeric(df[c_budget], errors="coerce")
    df[c_var] = pd.to_numeric(df[c_var], errors="coerce")

    real_total = float(df[c_real].fillna(0).sum())
    budget_total = float(df[c_budget].fillna(0).sum())
    var_total = float(df[c_var].fillna(0).sum())

    over_amt = max(0.0, real_total - budget_total)
    under_amt = max(0.0, budget_total - real_total)

    pct_of_budget = safe_pct(real_total, budget_total)
    pct_under_vs_budget = safe_pct(under_amt, budget_total)
    pct_over_vs_budget = safe_pct(over_amt, budget_total)

    if var_total < 0:
        status = "🔴 Over budget"
    elif var_total > 0:
        status = "🟢 Under budget"
    else:
        status = "⚪ On budget"

    rows.append({
        "Category": cat,
        "Real": real_total,
        "Budget": budget_total,
        "Variation (Budget - Real)": var_total,
        "% Budget Used": pct_of_budget,
        "Over %": pct_over_vs_budget,
        "Under %": pct_under_vs_budget,
        "Status": status,
    })

if not rows:
    st.warning("No category breakdown columns were found.")
else:
    df_cat = pd.DataFrame(rows)

    df_cat_show = df_cat.copy()
    df_cat_show["Real"] = df_cat_show["Real"].map(lambda x: f"${x:,.2f}")
    df_cat_show["Budget"] = df_cat_show["Budget"].map(lambda x: f"${x:,.2f}")
    df_cat_show["Variation (Budget - Real)"] = df_cat_show["Variation (Budget - Real)"].map(lambda x: f"${x:,.2f}")
    df_cat_show["% Budget Used"] = df_cat_show["% Budget Used"].map(lambda x: f"{x*100:,.1f}%")
    df_cat_show["Over %"] = df_cat_show["Over %"].map(lambda x: f"{x*100:,.1f}%")
    df_cat_show["Under %"] = df_cat_show["Under %"].map(lambda x: f"{x*100:,.1f}%")

    def highlight_variation(val):
        try:
            num = float(str(val).replace("$", "").replace(",", ""))
        except Exception:
            return ""
        if num < 0:
            return f"color: {color_red}; font-weight: 700;"
        elif num > 0:
            return f"color: {color_green}; font-weight: 700;"
        return f"color: {color_gray}; font-weight: 700;"

    def highlight_status(val):
        if "Over" in str(val):
            return f"color: {color_red}; font-weight: 700;"
        elif "Under" in str(val):
            return f"color: {color_green}; font-weight: 700;"
        return f"color: {color_gray}; font-weight: 700;"

    styled_cat = (
        df_cat_show.style
        .applymap(highlight_variation, subset=["Variation (Budget - Real)"])
        .applymap(highlight_status, subset=["Status"])
    )

    st.dataframe(styled_cat, use_container_width=True)

    fig_cat_all = go.Figure()
    fig_cat_all.add_trace(go.Bar(
        name="Budget",
        x=df_cat["Category"],
        y=df_cat["Budget"]
    ))
    fig_cat_all.add_trace(go.Bar(
        name="Real",
        x=df_cat["Category"],
        y=df_cat["Real"]
    ))

    fig_cat_all.update_layout(
        title="Budget vs Real by Category (Filtered)",
        barmode="group",
        xaxis_title="Category",
        yaxis_title="Amount",
        height=500,
    )

    fig_cat_all.update_xaxes(type="category")
    st.plotly_chart(fig_cat_all, use_container_width=True)


# ============================================================
# FIXED EXPENSES BREAKDOWN
# ============================================================
st.markdown("---")
st.subheader("🏢 Fixed Expenses Breakdown (Gasto Fijo)")

if apply_fixed:
    if df_fixed_breakdown is None or df_fixed_breakdown.empty:
        st.warning(f"No fixed-expense breakdown found in sheet '{fixed_sheet_name}'.")
    else:
        total_fx = float(df_fixed_breakdown["Amount"].fillna(0).sum())
        st.metric(f"Total Fixed Expenses (from sheet: {fixed_sheet_name})", f"${total_fx:,.2f}")

        df_fixed_show = df_fixed_breakdown.copy()
        df_fixed_show["Amount"] = df_fixed_show["Amount"].map(lambda v: f"{float(v):,.0f}")
        st.dataframe(df_fixed_show, use_container_width=True)

        fig_fx = go.Figure()
        fig_fx.add_trace(
            go.Bar(
                x=df_fixed_breakdown["Category"],
                y=df_fixed_breakdown["Amount"],
                text=df_fixed_breakdown["Amount"].map(lambda v: f"{float(v):,.0f}"),
                textposition="outside",
            )
        )
        fig_fx.update_layout(
            title=f"Fixed Expenses Breakdown - {fixed_sheet_name}",
            xaxis_title="Category",
            yaxis_title="Amount ($)",
            height=520,
        )
        fig_fx.update_xaxes(type="category", categoryorder="total descending")
        st.plotly_chart(fig_fx, use_container_width=True)
else:
    st.info(
        "Fixed expenses breakdown is shown only when Company filter is explicitly selected as exactly "
        "'12433087 Canada Inc' or '9359-6633 Quebec Inc', and no Project filter is active."
    )


# ============================================================
# PROJECT PROFIT / LOSS
# ============================================================
st.subheader("🧾 Project Profit / Loss (Filtered)")

pcol = find_col(df, "Project Name")
c_income2 = find_col(df, COL_INCOME)
c_total_cost_real = find_col(df, COL_COST_REAL)

if not pcol:
    st.info("Project column (Project Name) not found.")
elif not c_income2:
    st.info("Revenue column (Total to Bill) not found.")
elif not c_total_cost_real:
    st.info("Total Cost Real column not found (needed for Project P/L).")
else:
    df[c_income2] = pd.to_numeric(df[c_income2], errors="coerce")
    df[c_total_cost_real] = pd.to_numeric(df[c_total_cost_real], errors="coerce")

    p = (
        df.groupby(pcol, dropna=False)
        .agg(
            Revenue=(c_income2, "sum"),
            TotalCostReal=(c_total_cost_real, "sum"),
        )
        .reset_index()
    )
    p[pcol] = p[pcol].astype(str).replace({"nan": "None"})
    p["Profit/Loss"] = p["Revenue"] - p["TotalCostReal"]
    p["Margin %"] = p.apply(lambda r: safe_pct(r["Profit/Loss"], r["Revenue"]), axis=1)
    p = p.sort_values("Profit/Loss")

    def _color_pl(v):
        try:
            v = float(v)
        except Exception:
            return ""
        return "color: red; font-weight: 700;" if v < 0 else "color: green; font-weight: 700;"

    p_show = p.copy()
    p_show["Revenue"] = p_show["Revenue"].map(lambda x: f"${float(x):,.2f}")
    p_show["TotalCostReal"] = p_show["TotalCostReal"].map(lambda x: f"${float(x):,.2f}")

    sty = (
        p_show.style
        .format({"Profit/Loss": "${:,.2f}", "Margin %": "{:.1%}"})
        .applymap(_color_pl, subset=["Profit/Loss"])
        .applymap(
            lambda v: "color: red; font-weight: 700;" if float(v) < 0 else "color: green; font-weight: 700;",
            subset=["Margin %"]
        )
    )
    st.dataframe(sty, use_container_width=True)


# ============================================================
# MONTHLY BREAKDOWN
# ============================================================
st.subheader("🗓️ Monthly Breakdown (Filtered)")

if MONTH_COL in df.columns and YEAR_COL in df.columns:
    if "_MonthKey" not in df.columns or "_MonthText" not in df.columns:
        df = build_month_fields(df)

    selected_years = sorted([int(y) for y in df["_YearInt"].dropna().unique().tolist()])
    title_year_part = (
        f"Years {', '.join(map(str, selected_years))}"
        if len(selected_years) != 1
        else f"Year {selected_years[0]}"
    )

    df_m = df.copy()

    for col in [COL_INCOME, COL_COST_REAL, COL_MGMT, COL_ROY_5, COL_ROY_3]:
        c = find_col(df_m, col)
        if c:
            df_m[c] = pd.to_numeric(df_m[c], errors="coerce")

    mi = find_col(df_m, COL_INCOME)
    mc = find_col(df_m, COL_COST_REAL)
    mm = find_col(df_m, COL_MGMT)
    mr5 = find_col(df_m, COL_ROY_5)
    mr3 = find_col(df_m, COL_ROY_3)

    ok = df_m["_MonthKey"].notna() & df_m["_MonthText"].notna()
    if ok.any() and all([mi, mc, mm, mr5]):
        g = (
            df_m[ok]
            .groupby(["_MonthKey", "_MonthText"], dropna=True)
            .agg(
                Income=(mi, "sum"),
                Cost=(mc, "sum"),
                MgmtFee=(mm, "sum"),
                Royalty5=(mr5, "sum"),
                Royalty3=(mr3, "sum") if (mr3 and apply_roy3) else (mr5, lambda s: 0.0),
            )
            .reset_index()
            .sort_values("_MonthKey")
        )

        g["Gross"] = g["Income"] - g["Cost"]

        if apply_fixed:
            g["Fixed"] = fixed_total
            g["Net"] = g["Gross"] - g["Fixed"]
        else:
            g["Net"] = g["Gross"]

        if apply_roy3:
            g["New Total"] = g["Net"] + g["MgmtFee"] + g["Royalty5"] + g["Royalty3"]
        else:
            g["New Total"] = g["Net"] + g["MgmtFee"] + g["Royalty5"]

        cols_show = ["Month", "Income", "Cost", "Gross"]
        if apply_fixed:
            cols_show += ["Fixed", "Net"]
        else:
            cols_show += ["Net"]
        cols_show += ["MgmtFee", "Royalty5"]
        if apply_roy3:
            cols_show += ["Royalty3"]
        cols_show += ["New Total"]

        g_show = g.rename(columns={"_MonthText": "Month"}).copy()
        for c in [col for col in cols_show if col != "Month"]:
            g_show[c] = g_show[c].map(lambda x: f"${float(x):,.2f}")

        st.dataframe(g_show[cols_show], use_container_width=True)

        x_text = g["_MonthText"].tolist()

        fig_month = go.Figure()
        fig_month.add_trace(go.Bar(name="Income", x=x_text, y=g["Income"]))
        fig_month.add_trace(go.Bar(name="Cost", x=x_text, y=g["Cost"]))
        fig_month.add_trace(go.Scatter(name="New Total", x=x_text, y=g["New Total"], mode="lines+markers"))
        fig_month.update_layout(
            title=f"Month-to-Month (Filtered) - {title_year_part}: Income vs Cost + New Total",
            barmode="group",
            xaxis_title="Month",
            yaxis_title="Amount",
        )
        fig_month.update_xaxes(type="category", categoryorder="array", categoryarray=x_text)

        st.plotly_chart(fig_month, use_container_width=True)
    else:
        st.info("Could not build the monthly breakdown. Please verify Month/Year values and KPI columns.")
else:
    st.info("Month/Year columns were not found in the filtered dataframe.")


# ============================================================
# EXPORT PDF
# ============================================================
st.divider()
st.subheader("📄 Export Executive Report (PDF)")

report_company = ", ".join(filter_state.get("companies", [])) if filter_state.get("companies") else "All Companies"
report_building = ", ".join(filter_state.get("buildings", [])) if filter_state.get("buildings") else "All Buildings"
report_month = ", ".join(filter_state.get("months", [])) if filter_state.get("months") else "All Periods"

pdf_bytes = build_pdf_report(
    income=income,
    cost=cost,
    gross=gross,
    fixed_total_applied=fixed_total,
    fixed_sheet_name=fixed_sheet_name if fixed_sheet_name else "N/A",
    net=net,
    mgmt_fee_total=mgmt_fee_total,
    royalty_5_total=royalty_5_total,
    royalty_3_total=(royalty_3_total if apply_roy3 else 0.0),
    new_total=new_total,
    p_cost=p_cost,
    p_gross=p_gross,
    p_fixed=p_fixed,
    p_net=p_net,
    p_mgmt=p_mgmt,
    p_roy5=p_roy5,
    p_roy3=p_roy3,
    p_new=p_new,
    target=target,
    gross_margin=gross_margin,
    net_margin=net_margin,
    final_margin=final_margin,
    report_company=report_company,
    report_building=report_building,
    report_month=report_month,
    fig_waterfall=fig_waterfall,
    fig_gauge=fig_gauge,
)

st.download_button(
    "⬇️ Download Executive PDF",
    data=pdf_bytes,
    file_name="CNET_Executive_Report.pdf",
    mime="application/pdf",
)


# ============================================================
# TABLES
# ============================================================
st.subheader("Summary")

summary_rows = [
    ["Revenue", income, 1.0],
    ["Costs", cost, p_cost],
    ["Gross", gross, p_gross],
]

if apply_fixed:
    summary_rows += [[f"Fixed Expenses ({fixed_sheet_name})", fixed_total, p_fixed]]

summary_rows += [
    ["Net", net, p_net],
    ["Total Management Fee", mgmt_fee_total, p_mgmt],
    ["Royalty (5%)", royalty_5_total, p_roy5],
]

if apply_roy3:
    summary_rows += [["Royalty (3%) BGIS", royalty_3_total, p_roy3]]

summary_rows += [["New Total", new_total, p_new]]

summary = pd.DataFrame(summary_rows, columns=["Concept", "Amount", "% of Revenue"])
summary["Amount"] = summary["Amount"].map(lambda x: f"${x:,.2f}")
summary["% of Revenue"] = summary["% of Revenue"].map(lambda x: f"{x*100:,.2f}%")
st.dataframe(summary, use_container_width=True)

with st.expander("Real Master details (filtered)"):
    st.dataframe(sanitize_for_arrow(df), use_container_width=True)

with st.expander("Real Master details (unfiltered)"):
    st.dataframe(sanitize_for_arrow(df_all), use_container_width=True)
