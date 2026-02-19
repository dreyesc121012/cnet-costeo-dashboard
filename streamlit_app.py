import base64
from io import BytesIO
from datetime import datetime

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
SHEET_FIXED = "Gasto Fijo"
HEADER_IDX = 6

# Exact column names
MONTH_COL = "Month"   # text: "January", "February", etc.
YEAR_COL  = "Year"    # numeric or text year like 2026

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

@st.cache_data(ttl=300, show_spinner=False)
def load_fixed_total_from_bytes(excel_bytes: bytes) -> float:
    fx = pd.read_excel(BytesIO(excel_bytes), sheet_name=SHEET_FIXED, header=None)
    amounts = pd.to_numeric(fx.iloc[:, 2], errors="coerce")
    return float(amounts.fillna(0).sum())

def find_col(df: pd.DataFrame, name: str):
    if name in df.columns:
        return name
    n = name.strip().lower()
    for c in df.columns:
        if str(c).strip().lower() == n:
            return c
    for c in df.columns:
        if n in str(c).strip().lower():
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
        if "build" in str(c).strip().lower():
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
    """
    Adds:
      _YearInt   -> int year
      _MonthNum  -> 1..12
      _MonthKey  -> YYYY-MM (for sorting)
      _MonthText -> 'January 2026' (for display)
    """
    out = df.copy()

    out["_YearInt"] = pd.to_numeric(out[YEAR_COL], errors="coerce").astype("Int64")

    m = out[MONTH_COL].astype(str).str.strip().str.lower()
    m = m.str.replace(r"[^a-z]", "", regex=True)  # keep letters only
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
# FILTERS (Year + Month + Company + Province + Client + Project + Building)
# ============================================================
def add_filters(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("Executive Filters")

    if MONTH_COL in df.columns and YEAR_COL in df.columns:
        df = build_month_fields(df)

        years = sorted([int(y) for y in df["_YearInt"].dropna().unique().tolist()])
        sel_years = st.sidebar.multiselect("Year", years, default=[])
        if sel_years:
            df = df[df["_YearInt"].isin(sel_years)]

        month_table = (
            df[["_MonthKey", "_MonthText"]]
            .dropna()
            .drop_duplicates()
            .sort_values("_MonthKey")
        )
        sel_months = st.sidebar.multiselect("Month", month_table["_MonthText"].tolist(), default=[])
        if sel_months:
            df = df[df["_MonthText"].isin(sel_months)]
    else:
        st.sidebar.warning("Month/Year columns were not found in the Excel file.")

    # Company
    c_company = find_col(df, "Company")
    if c_company:
        sel = st.sidebar.multiselect("Company", sorted(df[c_company].dropna().unique()))
        if sel:
            df = df[df[c_company].isin(sel)]

    # Province
    c_prov = find_col(df, "Province")
    if c_prov:
        sel = st.sidebar.multiselect("Province", sorted(df[c_prov].dropna().unique()))
        if sel:
            df = df[df[c_prov].isin(sel)]

    # Client
    c_client = find_col(df, "Client")
    if c_client:
        sel = st.sidebar.multiselect("Client", sorted(df[c_client].dropna().unique()))
        if sel:
            df = df[df[c_client].isin(sel)]

    # Project Name
    c_proj = find_col(df, "Project Name")
    if c_proj:
        sel = st.sidebar.multiselect("Project (Project Name)", sorted(df[c_proj].dropna().unique()))
        if sel:
            df = df[df[c_proj].isin(sel)]

    # Building
    c_bld = pick_building_col(df)
    if c_bld:
        uniq = sorted(df[c_bld].dropna().astype(str).unique().tolist())
        sel = st.sidebar.multiselect("Building", uniq)
        if sel:
            df = df[df[c_bld].astype(str).isin(sel)]

    return df

# ============================================================
# PDF REPORT
# ============================================================
def build_pdf_report(
    *,
    income, cost, gross, fixed_total, net, mgmt_fee_total, royalty_total, new_total,
    p_cost, p_gross, p_fixed, p_net, p_mgmt, p_roy, p_new,
    target, gross_margin, net_margin, final_margin,
    fig_waterfall=None, fig_gauge=None,
):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    width, height = letter

    y = height - 50
    c.setFont("Helvetica-Bold", 16)
    c.drawString(40, y, "CNET Costing & Net - Executive Summary")
    y -= 16
    c.setFont("Helvetica", 9)
    c.drawString(40, y, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    y -= 10
    c.drawString(40, y, f"Target Margin: {target:.0%}")
    y -= 20

    rows = [
        ("Revenue (Total to Bill)", income, 1.0),
        ("Costs (Total Cost Month)", cost, p_cost),
        ("Gross (Revenue - Cost)", gross, p_gross),
        ("Fixed Expenses (Gasto Fijo)", fixed_total, p_fixed),
        ("Net (Gross - Fixed)", net, p_net),
        ("Management Fee", mgmt_fee_total, p_mgmt),
        ("Royalty 5%", royalty_total, p_roy),
        ("New Total", new_total, p_new),
    ]

    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "KPIs")
    y -= 14

    c.setFont("Helvetica", 10)
    c.drawString(40, y, "Concept")
    c.drawRightString(360, y, "Amount")
    c.drawRightString(520, y, "% of Revenue")
    y -= 12

    for label, val, pct in rows:
        c.drawString(40, y, label[:55])
        c.drawRightString(360, y, f"${val:,.2f}")
        c.drawRightString(520, y, f"{pct:.1%}")
        y -= 14

    y -= 8
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "Executive Margins")
    y -= 14
    c.setFont("Helvetica", 10)
    c.drawString(40, y, f"Gross Margin: {gross_margin:.1%}")
    y -= 14
    c.drawString(40, y, f"Net Margin: {net_margin:.1%}")
    y -= 14
    c.drawString(40, y, f"Final Margin (after fees): {final_margin:.1%}")
    y -= 18

    def add_plotly_image(fig, title, y_top):
        if fig is None:
            return y_top
        try:
            img_bytes = fig.to_image(format="png")  # requires kaleido
            img = ImageReader(BytesIO(img_bytes))
            c.setFont("Helvetica-Bold", 11)
            c.drawString(40, y_top, title)
            y_top -= 10
            c.drawImage(img, 40, y_top - 220, width=520, height=220, preserveAspectRatio=True, mask='auto')
            return y_top - 235
        except Exception:
            c.setFont("Helvetica", 9)
            c.drawString(40, y_top, f"{title} (could not export image - install kaleido)")
            return y_top - 15

    y = add_plotly_image(fig_gauge, "Gauge - Final Margin", y)
    y = add_plotly_image(fig_waterfall, "Waterfall - Revenue → Cost → Fixed → Fees → Total", y)

    c.showPage()
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
st.title("📊 CNET Costing & Net Dashboard")

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

# Header actions
colA, colB = st.columns([1, 3])
with colA:
    if st.button("🔄 Refresh data"):
        st.session_state.pop("excel_bytes", None)
        read_real_master_from_bytes.clear()
        load_fixed_total_from_bytes.clear()
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
fixed_total = load_fixed_total_from_bytes(excel_bytes)

# Apply filters (Year breakdown comes from the Year column)
df = add_filters(df_all.copy())

# ============================================================
# KPI base
# ============================================================
COL_INCOME = "Total to Bill"
COL_COST   = "Total Cost Month"
COL_MGMT   = "Total Management Fee"
COL_ROY    = "Royalty CNET Group Inc 5%"

# ============================================================
# CATEGORY SPECS (Budget vs Real uses Budget - Real in your Excel)
# ============================================================
CATEGORY_SPECS = {
    "Labor": {
        "real": "Total Labor Real",
        "budget": "Total Labor Budget",
        "var": "Variation Labor  (Budget vs Real)",   # Budget - Real
    },
    "PW": {
        "real": "Total PW Real",
        "budget": "Total PW Budget",
        "var": "Variation PW  (Budget vs Real)",      # Budget - Real
    },
    "Supplies": {
        "real": "Total Supplies Real",
        "budget": "Total Supplies Budget",
        "var": "Variation Supplies  (Budget vs Real)",# Budget - Real
    },
    "Equipment": {
        "real": "Total Equipment Real",
        "budget": "Total Equipment Budget",
        "var": "Variation Equipment (Budget vs Real)",# Budget - Real
    },
    "Total Cost": {
        "real": "Total Cost Real",
        "budget": "Total Cost Budget",
        "var": "Variation Total Cost (Budget vs Real)",# Budget - Real
    },
}

c_income = find_col(df, COL_INCOME)
c_cost   = find_col(df, COL_COST)
c_mgmt   = find_col(df, COL_MGMT)
c_roy    = find_col(df, COL_ROY)

missing = [k for k, v in {
    COL_INCOME: c_income,
    COL_COST: c_cost,
    COL_MGMT: c_mgmt,
    COL_ROY: c_roy,
}.items() if v is None]

if missing:
    st.error(f"Missing columns in 'Real Master': {missing}")
    with st.expander("Show detected columns"):
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

p_cost  = safe_pct(cost, income)
p_gross = safe_pct(gross, income)
p_fixed = safe_pct(fixed_total, income)
p_net   = safe_pct(net, income)
p_mgmt  = safe_pct(mgmt_fee_total, income)
p_roy   = safe_pct(royalty_total, income)
p_new   = safe_pct(new_total, income)

# ============================================================
# Traffic Light + Gauge
# ============================================================
st.subheader("📌 Executive Margin (KPIs + Traffic Light)")

target = st.slider("Target margin (%)", 0, 60, 25) / 100
yellow_zone = 0.05

gross_margin = gross / income if income else 0
net_margin   = net / income if income else 0
final_margin = new_total / income if income else 0

def traffic_light(m, tgt):
    if m >= tgt + yellow_zone:
        return "🟢"
    elif m >= tgt - yellow_zone:
        return "🟡"
    return "🔴"

c1, c2, c3 = st.columns(3)
c1.metric("Gross Margin", f"{gross_margin:.1%}", f"{traffic_light(gross_margin, target)} vs {target:.0%}")
c2.metric("Net Margin", f"{net_margin:.1%}", f"{traffic_light(net_margin, target)} vs {target:.0%}")
c3.metric("Final Margin (after fees)", f"{final_margin:.1%}", f"{traffic_light(final_margin, target)} vs {target:.0%}")

st.caption("Gauge: Final margin (after fees)")
gauge_max = 60
fig_gauge = go.Figure(go.Indicator(
    mode="gauge+number",
    value=float(final_margin * 100),
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
k1, k2, k3, k4 = st.columns(4)
k1.metric("Revenue (Total to Bill)", f"${income:,.2f}")
k2.metric("Costs (Total Cost Month)", f"${cost:,.2f}", f"{p_cost*100:,.2f}%")
k3.metric("Gross (Revenue - Cost)", f"${gross:,.2f}", f"{p_gross*100:,.2f}%")
k4.metric("Fixed Expenses (Gasto Fijo)", f"${fixed_total:,.2f}", f"{p_fixed*100:,.2f}%")

k5, k6, k7, k8 = st.columns(4)
k5.metric("Net (Gross - Fixed)", f"${net:,.2f}", f"{p_net*100:,.2f}%")
k6.metric("Total Management Fee", f"${mgmt_fee_total:,.2f}", f"{p_mgmt*100:,.2f}%")
k7.metric("Royalty (5%)", f"${royalty_total:,.2f}", f"{p_roy*100:,.2f}%")
k8.metric("New Total", f"${new_total:,.2f}", f"{p_new*100:,.2f}%")

# ============================================================
# OPTIONAL KPI: Total Cost Budget vs Real + % used / variance
# ============================================================
tc_r = find_col(df, "Total Cost Real")
tc_b = find_col(df, "Total Cost Budget")
tc_v = find_col(df, "Variation Total Cost (Budget vs Real)")  # Budget - Real

if tc_r and tc_b and tc_v:
    df[tc_r] = pd.to_numeric(df[tc_r], errors="coerce")
    df[tc_b] = pd.to_numeric(df[tc_b], errors="coerce")
    df[tc_v] = pd.to_numeric(df[tc_v], errors="coerce")

    total_cost_real = float(df[tc_r].fillna(0).sum())
    total_cost_budget = float(df[tc_b].fillna(0).sum())
    total_cost_var = float(df[tc_v].fillna(0).sum())  # Budget - Real

    pct_used = safe_pct(total_cost_real, total_cost_budget)  # Real / Budget
    pct_under = safe_pct(max(0.0, total_cost_var), total_cost_budget)  # Under % (Budget-Real)/Budget
    pct_over = safe_pct(max(0.0, total_cost_real - total_cost_budget), total_cost_budget)  # Over % (Real-Budget)/Budget

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
fig_waterfall = go.Figure(go.Waterfall(
    orientation="v",
    measure=["absolute", "relative", "relative", "relative", "relative", "total"],
    x=["Revenue", "Costs", "Gross", "Fixed", "Mgmt+Royalty", "New Total"],
    y=[income, -cost, gross, -fixed_total, (mgmt_fee_total + royalty_total), new_total],
))
fig_waterfall.update_layout(title="Waterfall: Revenue → Costs → Fixed → +Fees → New Total", showlegend=False)
st.plotly_chart(fig_waterfall, use_container_width=True)

# ============================================================
# ✅ CATEGORY BUDGET vs REAL BREAKDOWN (FILTERED) — Budget - Real
# ============================================================
st.subheader("🧩 Budget vs Real Breakdown (Categories)")

cat = st.selectbox("Select category", list(CATEGORY_SPECS.keys()), index=0)
spec = CATEGORY_SPECS[cat]

c_real = find_col(df, spec["real"])
c_budget = find_col(df, spec["budget"])
c_var = find_col(df, spec["var"])  # Budget - Real

missing_cat = [k for k, v in {
    spec["real"]: c_real,
    spec["budget"]: c_budget,
    spec["var"]: c_var,
}.items() if v is None]

if missing_cat:
    st.error(f"Missing columns for '{cat}': {missing_cat}")
    with st.expander("Show detected columns"):
        st.write(df.columns.tolist())
else:
    df[c_real] = pd.to_numeric(df[c_real], errors="coerce")
    df[c_budget] = pd.to_numeric(df[c_budget], errors="coerce")
    df[c_var] = pd.to_numeric(df[c_var], errors="coerce")

    real_total = float(df[c_real].fillna(0).sum())
    budget_total = float(df[c_budget].fillna(0).sum())
    var_total = float(df[c_var].fillna(0).sum())  # Budget - Real

    # Over/Under amounts
    over_amt = max(0.0, real_total - budget_total)     # Real - Budget (when over)
    under_amt = max(0.0, budget_total - real_total)    # Budget - Real (when under) = max(0, var_total)

    # % consumption + % over/under
    pct_of_budget = safe_pct(real_total, budget_total)               # Real / Budget
    pct_under_vs_budget = safe_pct(under_amt, budget_total)          # Under / Budget
    pct_over_vs_budget  = safe_pct(over_amt, budget_total)           # Over / Budget

    # Status (Budget - Real)
    status = "🟢 On track"
    if var_total < 0:
        status = "🔴 Over budget"
    elif var_total > 0:
        status = "🟢 Under budget"

    a1, a2, a3, a4 = st.columns(4)
    a1.metric(f"{cat} - Real", f"${real_total:,.2f}")
    a2.metric(f"{cat} - Budget", f"${budget_total:,.2f}")
    a3.metric("Variation (Budget - Real)", f"${var_total:,.2f}", f"Over: ${over_amt:,.2f} | Under: ${under_amt:,.2f}")
    a4.metric("% of Budget Used", f"{pct_of_budget*100:,.1f}%", f"Over: {pct_over_vs_budget*100:,.1f}% | Under: {pct_under_vs_budget*100:,.1f}%")

    fig_cat = go.Figure()
    fig_cat.add_trace(go.Bar(name="Budget", x=["Budget"], y=[budget_total]))
    fig_cat.add_trace(go.Bar(name="Real", x=["Real"], y=[real_total]))
    fig_cat.update_layout(
        title=f"{cat}: Budget vs Real (Filtered)",
        barmode="group",
        xaxis_title="",
        yaxis_title="Amount",
    )
    st.plotly_chart(fig_cat, use_container_width=True)

# ============================================================
# ✅ BUILDING PROFIT / LOSS (RED if LOSS, GREEN if PROFIT)
# Uses: Revenue (Total to Bill) - Total Cost Real
# ============================================================
st.subheader("🏢 Building Profit / Loss (Filtered)")

bcol = pick_building_col(df)
c_income2 = find_col(df, COL_INCOME)
c_total_cost_real = find_col(df, "Total Cost Real")

if not bcol:
    st.info("Building column not found.")
elif not c_income2:
    st.info("Revenue column (Total to Bill) not found.")
elif not c_total_cost_real:
    st.info("Total Cost Real column not found (needed for building P/L).")
else:
    df[c_income2] = pd.to_numeric(df[c_income2], errors="coerce")
    df[c_total_cost_real] = pd.to_numeric(df[c_total_cost_real], errors="coerce")

    b = (
        df.groupby(bcol, dropna=False)
          .agg(
              Revenue=(c_income2, "sum"),
              TotalCostReal=(c_total_cost_real, "sum"),
          )
          .reset_index()
    )
    b["Profit/Loss"] = b["Revenue"] - b["TotalCostReal"]
    b["Margin %"] = b.apply(lambda r: safe_pct(r["Profit/Loss"], r["Revenue"]), axis=1)

    # Sort: worst first
    b = b.sort_values("Profit/Loss")

    def _color_pl(v):
        try:
            v = float(v)
        except Exception:
            return ""
        return "color: red; font-weight: 700;" if v < 0 else "color: green; font-weight: 700;"

    b_show = b.copy()
    b_show["Revenue"] = b_show["Revenue"].map(lambda x: f"${float(x):,.2f}")
    b_show["TotalCostReal"] = b_show["TotalCostReal"].map(lambda x: f"${float(x):,.2f}")

    sty = (
        b_show.style
        .format({"Profit/Loss": "${:,.2f}", "Margin %": "{:.1%}"})
        .applymap(_color_pl, subset=["Profit/Loss"])
        .applymap(lambda v: "color: red; font-weight: 700;" if float(v) < 0 else "color: green; font-weight: 700;",
                  subset=["Margin %"])
    )
    st.dataframe(sty, use_container_width=True)

# ============================================================
# ✅ MONTHLY BREAKDOWN (FILTERED) — uses Year column, NO TIME AXIS
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

    for col in [COL_INCOME, COL_COST, COL_MGMT, COL_ROY]:
        c = find_col(df_m, col)
        if c:
            df_m[c] = pd.to_numeric(df_m[c], errors="coerce")

    mi = find_col(df_m, COL_INCOME)
    mc = find_col(df_m, COL_COST)
    mm = find_col(df_m, COL_MGMT)
    mr = find_col(df_m, COL_ROY)

    ok = df_m["_MonthKey"].notna() & df_m["_MonthText"].notna()
    if ok.any() and all([mi, mc, mm, mr]):
        g = (
            df_m[ok]
            .groupby(["_MonthKey", "_MonthText"], dropna=True)
            .agg(
                Income=(mi, "sum"),
                Cost=(mc, "sum"),
                MgmtFee=(mm, "sum"),
                Royalty=(mr, "sum"),
            )
            .reset_index()
            .sort_values("_MonthKey")
        )

        g["Gross"] = g["Income"] - g["Cost"]
        g["Net (Gross - Fixed)"] = g["Gross"] - fixed_total
        g["New Total"] = g["Net (Gross - Fixed)"] + g["MgmtFee"] + g["Royalty"]

        g_show = g.rename(columns={"_MonthText": "Month"}).copy()
        for c in ["Income", "Cost", "Gross", "MgmtFee", "Royalty", "Net (Gross - Fixed)", "New Total"]:
            g_show[c] = g_show[c].map(lambda x: f"${float(x):,.2f}")
        st.dataframe(
            g_show[["Month", "Income", "Cost", "Gross", "MgmtFee", "Royalty", "Net (Gross - Fixed)", "New Total"]],
            use_container_width=True
        )

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

pdf_bytes = build_pdf_report(
    income=income, cost=cost, gross=gross, fixed_total=fixed_total, net=net,
    mgmt_fee_total=mgmt_fee_total, royalty_total=royalty_total, new_total=new_total,
    p_cost=p_cost, p_gross=p_gross, p_fixed=p_fixed, p_net=p_net, p_mgmt=p_mgmt, p_roy=p_roy, p_new=p_new,
    target=target, gross_margin=gross_margin, net_margin=net_margin, final_margin=final_margin,
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
summary = pd.DataFrame([
    ["Revenue", income, 1.0],
    ["Costs", cost, p_cost],
    ["Gross", gross, p_gross],
    ["Fixed Expenses", fixed_total, p_fixed],
    ["Net", net, p_net],
    ["Total Management Fee", mgmt_fee_total, p_mgmt],
    ["Royalty (5%)", royalty_total, p_roy],
    ["New Total", new_total, p_new],
], columns=["Concept", "Amount", "% of Revenue"])

summary["Amount"] = summary["Amount"].map(lambda x: f"${x:,.2f}")
summary["% of Revenue"] = summary["% of Revenue"].map(lambda x: f"{x*100:,.2f}%")
st.dataframe(summary, use_container_width=True)

with st.expander("Real Master details (filtered)"):
    st.dataframe(sanitize_for_arrow(df), use_container_width=True)

with st.expander("Real Master details (unfiltered)"):
    st.dataframe(sanitize_for_arrow(df_all), use_container_width=True)
