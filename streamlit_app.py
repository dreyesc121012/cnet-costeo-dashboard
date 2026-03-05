import base64
from io import BytesIO
from datetime import datetime
import re
import os

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
# LOGO CONFIG (local file in repo)
# Put your file in the repo root as: cnet_logo.png
# ============================================================
LOGO_FILE = "cnet_logo.png"

def _logo_path() -> str:
    """Return absolute path for the logo if it exists, else empty string."""
    try:
        p = os.path.join(os.getcwd(), LOGO_FILE)
        return p if os.path.exists(p) else ""
    except Exception:
        return ""

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

# ============================================================
# EXCEL CONFIG
# ============================================================
SHEET_REAL = "Real Master"
HEADER_IDX = 6

# Fixed expenses sheets by company (PER YOUR REQUEST)
COMPANY_FIXED_SHEETS = {
    "12433087 Canada Inc": "Gasto Fijo",
    "9359-6633 Quebec Inc": "Gasto Fijo 9359",
}

# Exact column names
MONTH_COL = "Month"   # text: "January", "February", etc.
YEAR_COL  = "Year"    # numeric or text year like 2026

st.set_page_config(page_title="Financial Performance & Budget Control", layout="wide")

# ============================================================
# HEADER (Logo + Title)
# ============================================================
lp = _logo_path()
h1, h2 = st.columns([1, 4])
with h1:
    if lp:
        st.image(lp, use_container_width=True)
with h2:
    st.markdown("## Financial Performance & Budget Control")
    st.caption("Real vs Budget • Executive KPIs • Margin Control")

st.divider()

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

def _pick_amount_column(df: pd.DataFrame):
    """
    Find an amount/cost column. Prefer columns whose name suggests 'amount',
    otherwise pick the column with most numeric values.
    """
    for c in df.columns:
        cname = str(c).strip().lower()
        if any(k in cname for k in ["amount", "importe", "monto", "valor", "total", "cost", "costo"]):
            return c

    best_col = None
    best_score = -1
    for c in df.columns:
        s = pd.to_numeric(df[c], errors="coerce")
        score = int(s.notna().sum())
        if score > best_score:
            best_score = score
            best_col = c
    return best_col

def _pick_name_column(df: pd.DataFrame, amount_col):
    """
    Pick the best text column to use as Name:
    - Excludes amount column
    - Prefers columns with MOST non-empty text values (not numeric)
    """
    cols = [c for c in df.columns if c != amount_col]
    if not cols:
        return None

    best_col = None
    best_score = -1

    for c in cols:
        s = df[c]

        sn = pd.to_numeric(s, errors="coerce")
        numeric_ratio = sn.notna().mean() if len(s) else 1.0
        if numeric_ratio > 0.60:
            continue

        txt = s.astype(str).str.strip()
        txt = txt.replace({"nan": "", "None": "", "NaT": ""})
        score = int(txt.ne("").sum())

        if score > best_score:
            best_score = score
            best_col = c

    return best_col

@st.cache_data(ttl=300, show_spinner=False)
def load_fixed_detail_from_bytes(excel_bytes: bytes, sheet_name: str) -> tuple[float, pd.DataFrame]:
    """
    Returns (fixed_total, fixed_detail_df) for a given fixed-expense sheet.
    Output table: Category (optional) + Name + Amount
    """
    # -------------------------------------------------
    # 1) TRY WITH HEADERS
    # -------------------------------------------------
    try:
        fx = pd.read_excel(BytesIO(excel_bytes), sheet_name=sheet_name)
        fx = fx.dropna(how="all")
        if fx.empty:
            return 0.0, fx

        amt_col = _pick_amount_column(fx)
        if amt_col is None:
            raise ValueError("No amount column detected.")

        fx["_Amount"] = pd.to_numeric(fx[amt_col], errors="coerce").fillna(0.0)
        fixed_total = float(fx["_Amount"].sum())

        detail = fx[fx["_Amount"] != 0].copy()
        if detail.empty:
            return fixed_total, pd.DataFrame(columns=["Name", "Amount"])

        # Build candidate text columns
        text_cols = []
        for c in detail.columns:
            if c in (amt_col, "_Amount"):
                continue
            s = detail[c]
            sn = pd.to_numeric(s, errors="coerce")
            if (sn.notna().mean() if len(s) else 1.0) <= 0.60:
                text_cols.append(c)

        def clean_text(series):
            t = series.astype(str).str.strip()
            t = t.replace({"nan": "", "None": "", "NaT": ""})
            return t

        # Pick best name column
        name_col = _pick_name_column(detail, amt_col)

        out = pd.DataFrame()

        if name_col is not None:
            nm = clean_text(detail[name_col])
            empty_ratio = (nm.eq("")).mean()

            # If too empty, combine first 2 text cols
            if empty_ratio > 0.50 and len(text_cols) >= 2:
                a = clean_text(detail[text_cols[0]])
                b = clean_text(detail[text_cols[1]])
                nm2 = (a + " - " + b).str.strip(" -")
                nm = nm.mask(nm.eq(""), nm2)
            elif empty_ratio > 0.50 and len(text_cols) >= 1:
                a = clean_text(detail[text_cols[0]])
                nm = nm.mask(nm.eq(""), a)

            out["Name"] = nm
        else:
            if len(text_cols) >= 2:
                a = clean_text(detail[text_cols[0]])
                b = clean_text(detail[text_cols[1]])
                out["Name"] = (a + " - " + b).str.strip(" -")
            elif len(text_cols) == 1:
                out["Name"] = clean_text(detail[text_cols[0]])
            else:
                out["Name"] = ["(Unnamed)"] * len(detail)

        # Optional Category: first other text col with enough non-empty values
        cat_col = None
        for c in text_cols:
            if name_col is not None and c == name_col:
                continue
            ct = clean_text(detail[c])
            if ct.ne("").sum() >= max(3, int(0.20 * len(ct))):
                cat_col = c
                break

        if cat_col is not None:
            out.insert(0, "Category", clean_text(detail[cat_col]))

        out["Amount"] = detail["_Amount"].values

        # Final cleanup
        out["Name"] = out["Name"].astype(str).replace({"nan": "", "None": ""}).str.strip()
        out = out[out["Name"].ne("")].copy()

        out["Amount"] = pd.to_numeric(out["Amount"], errors="coerce").fillna(0.0)
        out = out.sort_values("Amount", ascending=False).reset_index(drop=True)

        return fixed_total, out

    except Exception:
        pass

    # -------------------------------------------------
    # 2) FALLBACK RAW (header=None)
    # -------------------------------------------------
    fx_raw = pd.read_excel(BytesIO(excel_bytes), sheet_name=sheet_name, header=None)
    fx_raw = fx_raw.dropna(how="all")
    if fx_raw.empty:
        return 0.0, fx_raw

    # Amount col: prefer index 2
    idx_amount = 2 if fx_raw.shape[1] >= 3 else (fx_raw.shape[1] - 1)
    fx_raw["_Amount"] = pd.to_numeric(fx_raw.iloc[:, idx_amount], errors="coerce").fillna(0.0)
    fixed_total = float(fx_raw["_Amount"].sum())

    detail = fx_raw[fx_raw["_Amount"] != 0].copy()
    if detail.empty:
        return fixed_total, pd.DataFrame(columns=["Name", "Amount"])

    candidate_idxs = [i for i in range(detail.shape[1]) if i != idx_amount]

    def score_text_col(idx):
        s = detail.iloc[:, idx].astype(str).str.strip()
        s = s.replace({"nan": "", "None": "", "NaT": ""})
        sn = pd.to_numeric(s, errors="coerce")
        numeric_ratio = sn.notna().mean() if len(s) else 1.0
        if numeric_ratio > 0.60:
            return -1
        return int(s.ne("").sum())

    best_name_idx = None
    best_score = -1
    for idx in candidate_idxs:
        sc = score_text_col(idx)
        if sc > best_score:
            best_score = sc
            best_name_idx = idx

    def clean_series(idx):
        s = detail.iloc[:, idx].astype(str).str.strip()
        s = s.replace({"nan": "", "None": "", "NaT": ""})
        return s

    out = pd.DataFrame()

    if best_name_idx is not None and best_score > 0:
        nm = clean_series(best_name_idx)
        empty_ratio = (nm.eq("")).mean()

        # If too empty, fill blanks using second best text column
        if empty_ratio > 0.50 and len(candidate_idxs) >= 2:
            scores = [(idx, score_text_col(idx)) for idx in candidate_idxs if idx != best_name_idx]
            scores = sorted(scores, key=lambda x: x[1], reverse=True)
            if scores and scores[0][1] > 0:
                nm2 = clean_series(scores[0][0])
                nm = nm.mask(nm.eq(""), nm2)

        out["Name"] = nm
    else:
        out["Name"] = ["(Unnamed)"] * len(detail)

    out["Amount"] = detail["_Amount"].values
    out["Name"] = out["Name"].astype(str).replace({"nan": "", "None": ""}).str.strip()
    out = out[out["Name"].ne("")].copy()

    out["Amount"] = pd.to_numeric(out["Amount"], errors="coerce").fillna(0.0)
    out = out.sort_values("Amount", ascending=False).reset_index(drop=True)

    return fixed_total, out

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
# PDF REPORT (dynamic rows: Fixed appears only when applied)
# + LOGO IN PDF HEADER
# ============================================================
def build_pdf_report(
    *,
    income, cost, gross,
    fixed_total_applied, fixed_total_full,
    net, mgmt_fee_total, royalty_5_total, royalty_3_total, new_total,
    p_cost, p_gross, p_fixed, p_net, p_mgmt, p_roy5, p_roy3, p_new,
    target, gross_margin, net_margin, final_margin,
    fig_waterfall=None, fig_gauge=None,
):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    width, height = letter

    y_top = height - 40
    lp = _logo_path()
    if lp:
        try:
            c.drawImage(
                ImageReader(lp),
                40, y_top - 30,
                width=140, height=30,
                preserveAspectRatio=True,
                mask="auto"
            )
        except Exception:
            pass

    c.setFont("Helvetica-Bold", 16)
    c.drawString(200, y_top - 6, "Financial Performance & Budget Control")
    c.setFont("Helvetica", 9)
    c.drawString(200, y_top - 20, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    c.drawString(200, y_top - 32, f"Target Margin: {target:.0%}")

    y = y_top - 60

    rows = [
        ("Revenue (Total to Bill)", income, 1.0),
        ("Costs (Total Cost Real)", cost, p_cost),
        ("Gross (Revenue - Cost)", gross, p_gross),
    ]

    if fixed_total_applied != 0.0:
        rows += [
            ("Fixed Expenses (Gasto Fijo)", fixed_total_applied, p_fixed),
            ("Net (Gross - Fixed)", net, p_net),
        ]
    else:
        rows += [("Net", net, p_net)]

    rows += [
        ("Management Fee", mgmt_fee_total, p_mgmt),
        ("Royalty (5%)", royalty_5_total, p_roy5),
    ]

    if royalty_3_total != 0.0:
        rows += [("Royalty (3%) BGIS", royalty_3_total, p_roy3)]

    rows += [("New Total", new_total, p_new)]

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

    def add_plotly_image(fig, title, y_top_local):
        if fig is None:
            return y_top_local
        try:
            img_bytes = fig.to_image(format="png")
            img = ImageReader(BytesIO(img_bytes))
            c.setFont("Helvetica-Bold", 11)
            c.drawString(40, y_top_local, title)
            y_top_local -= 10
            c.drawImage(img, 40, y_top_local - 220, width=520, height=220, preserveAspectRatio=True, mask='auto')
            return y_top_local - 235
        except Exception:
            c.setFont("Helvetica", 9)
            c.drawString(40, y_top_local, f"{title} (could not export image - install kaleido)")
            return y_top_local - 15

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

if not _logo_path():
    st.warning("⚠️ Logo file not found. Please ensure 'cnet_logo.png' is in the repo root.")

colA, colB = st.columns([1, 3])
with colA:
    if st.button("🔄 Refresh data"):
        st.session_state.pop("excel_bytes", None)
        read_real_master_from_bytes.clear()
        load_fixed_detail_from_bytes.clear()
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
df = add_filters(df_all.copy())

# ============================================================
# KPI base
# ============================================================
COL_INCOME = "Total to Bill"
COL_COST_REAL   = "Total Cost Real"
COL_COST_BUDGET = "Total Cost Budget"
COL_COST_VAR    = "Variation Total Cost (Budget vs Real)"
COL_MGMT   = "Total Management Fee"
COL_ROY_5  = "Royalty CNET Group Inc 5%"
COL_ROY_3  = "Royalty CNET Master 3% BGIS"

# Resolve columns
c_income = find_col(df, COL_INCOME)
c_cost   = find_col(df, COL_COST_REAL)
c_mgmt   = find_col(df, COL_MGMT)
c_roy5   = find_col(df, COL_ROY_5)
c_roy3   = find_col(df, COL_ROY_3)

c_company = find_col(df, "Company")
c_client  = find_col(df, "Client")

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

for c in [c_income, c_cost, c_mgmt, c_roy5]:
    df[c] = pd.to_numeric(df[c], errors="coerce")
if c_roy3:
    df[c_roy3] = pd.to_numeric(df[c_roy3], errors="coerce")

COMPANY_BG_QC = "9359-6633 Quebec Inc"
CLIENT_BGIS   = "BGIS"

income = float(df[c_income].fillna(0).sum())
cost   = float(df[c_cost].fillna(0).sum())
gross  = income - cost

mgmt_fee_total   = float(df[c_mgmt].fillna(0).sum())
royalty_5_total  = float(df[c_roy5].fillna(0).sum())
royalty_3_total  = float(df[c_roy3].fillna(0).sum()) if c_roy3 else 0.0

# Fixed expenses logic
apply_fixed = False
fixed_total = 0.0
fixed_total_full = 0.0
fixed_detail_df = pd.DataFrame()
fixed_sheet_used = ""

companies = []
if c_company:
    companies = df[c_company].dropna().astype(str).unique().tolist()

if len(companies) == 1:
    comp_selected = companies[0].strip()
    if comp_selected in COMPANY_FIXED_SHEETS:
        apply_fixed = True
        fixed_sheet_used = COMPANY_FIXED_SHEETS[comp_selected]
        fixed_total, fixed_detail_df = load_fixed_detail_from_bytes(excel_bytes, fixed_sheet_used)
        fixed_total_full = fixed_total

net = gross - (fixed_total if apply_fixed else 0.0)

# Royalty 3% logic
apply_roy3 = False
clients = []
if c_client:
    clients = df[c_client].dropna().astype(str).unique().tolist()

only_company_bg_qc = (len(companies) == 1 and companies[0] == COMPANY_BG_QC)
only_client_bgis   = (len(clients) == 1 and clients[0].strip().lower() == CLIENT_BGIS.lower())
apply_roy3 = (only_company_bg_qc or only_client_bgis)

if apply_roy3:
    new_total = net + mgmt_fee_total + royalty_5_total + royalty_3_total
else:
    new_total = net + mgmt_fee_total + royalty_5_total

p_cost  = safe_pct(cost, income)
p_gross = safe_pct(gross, income)
p_fixed = safe_pct(fixed_total if apply_fixed else 0.0, income)
p_net   = safe_pct(net, income)
p_mgmt  = safe_pct(mgmt_fee_total, income)
p_roy5  = safe_pct(royalty_5_total, income)
p_roy3  = safe_pct(royalty_3_total, income) if apply_roy3 else 0.0
p_new   = safe_pct(new_total, income)

# ============================================================
# Executive Margin
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
# KPIs (Executive)
# ============================================================
st.subheader("📊 KPIs (Executive)")

left, right = st.columns([2, 1], gap="large")

# --- LEFT SIDE KPIs ---
with left:
    # 1) Revenue
    st.metric("Revenue (Total to Bill)", f"${income:,.2f}")

    # 2) Costs
    st.metric("Costs (Total Cost Real)", f"${cost:,.2f}", f"{p_cost*100:,.2f}%")

    # 3) Gross
    st.metric("Gross (Revenue - Cost)", f"${gross:,.2f}", f"{p_gross*100:,.2f}%")

    # 4) Fixed (only if applied)
    if apply_fixed:
        st.metric("Fixed Expenses (Gasto Fijo)", f"${fixed_total:,.2f}", f"{p_fixed*100:,.2f}%")

    # 5) Net (always)
    net_label = "Net (Gross - Fixed)" if apply_fixed else "Net"
    st.metric(net_label, f"${net:,.2f}", f"{p_net*100:,.2f}%")

# --- RIGHT SIDE KPIs ---
with right:
    st.metric("Total Management Fee", f"${mgmt_fee_total:,.2f}", f"{p_mgmt*100:,.2f}%")
    st.metric("Royalty (5%)", f"${royalty_5_total:,.2f}", f"{p_roy5*100:,.2f}%")

    # Royalty 3% only if applies (else show $0.00)
    roy3_value = royalty_3_total if apply_roy3 else 0.0
    roy3_delta = (p_roy3 * 100) if apply_roy3 else 0.0
    st.metric("Royalty (3%) BGIS", f"${roy3_value:,.2f}", f"{roy3_delta:,.2f}%")
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
# ✅ FIXED EXPENSES DETAIL (only when applied) — SHOW NAME + AMOUNT
# ============================================================
if apply_fixed:
    st.divider()
    st.subheader("📌 Fixed Expenses Detail (Gasto Fijo)")

    left, right = st.columns([1, 2])
    with left:
        st.metric("Fixed Expenses Total", f"${fixed_total:,.2f}")
        if fixed_sheet_used:
            st.caption(f"Source sheet: {fixed_sheet_used}")

    with right:
        if fixed_detail_df is None or fixed_detail_df.empty:
            st.info("No fixed-expenses detail rows found (or the sheet is empty).")
        else:
            show_fx = fixed_detail_df.copy()
            if "Amount" in show_fx.columns:
                show_fx["Amount"] = pd.to_numeric(show_fx["Amount"], errors="coerce").fillna(0.0)
                show_fx = show_fx.sort_values("Amount", ascending=False)
            st.dataframe(sanitize_for_arrow(show_fx), use_container_width=True)

# ============================================================
# EXPORT PDF
# ============================================================
st.divider()
st.subheader("📄 Export Executive Report (PDF)")

pdf_bytes = build_pdf_report(
    income=income, cost=cost, gross=gross,
    fixed_total_applied=(fixed_total if apply_fixed else 0.0),
    fixed_total_full=fixed_total_full,
    net=net,
    mgmt_fee_total=mgmt_fee_total,
    royalty_5_total=royalty_5_total,
    royalty_3_total=(royalty_3_total if apply_roy3 else 0.0),
    new_total=new_total,
    p_cost=p_cost, p_gross=p_gross, p_fixed=p_fixed, p_net=p_net,
    p_mgmt=p_mgmt, p_roy5=p_roy5, p_roy3=p_roy3, p_new=p_new,
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
# SUMMARY TABLE
# ============================================================
st.subheader("Summary")

summary_rows = [
    ["Revenue", income, 1.0],
    ["Costs", cost, p_cost],
    ["Gross", gross, p_gross],
]
if apply_fixed:
    summary_rows += [["Fixed Expenses (Gasto Fijo)", fixed_total, p_fixed]]

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
