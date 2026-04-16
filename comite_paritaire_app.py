import base64
from io import BytesIO
from datetime import timedelta, datetime
import unicodedata

import pandas as pd
import requests
import streamlit as st
import msal


# ============================================================
# PAGE
# ============================================================
st.set_page_config(page_title="Comité Paritaire QC", layout="wide")
st.title("Comité Paritaire Québec - Weekly Report")


# ============================================================
# CONFIG (SECRETS)
# ============================================================
CLIENT_ID = str(st.secrets["CLIENT_ID"]).strip()
CLIENT_SECRET = str(st.secrets["CLIENT_SECRET"]).strip()
TENANT_ID = str(st.secrets["TENANT_ID"]).strip()
REDIRECT_URI = str(st.secrets["REDIRECT_URI"]).strip().rstrip("/")
ONEDRIVE_FOLDER_URL = str(st.secrets["ONEDRIVE_FOLDER_URL"]).strip()
ALLOWED_DOMAIN = str(st.secrets.get("ALLOWED_DOMAIN", "")).strip().lower()
DOMAIN_HINT = str(st.secrets.get("DOMAIN_HINT", "")).strip()
LOGIN_HINT = str(st.secrets.get("LOGIN_HINT", "")).strip()

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.Read.All", "Sites.Read.All"]


# ============================================================
# BUSINESS RULES
# ============================================================
COMITE_CLASS_A_RATE = 21.57
REER_PER_HOUR = 0.45

ROW_ORDER = [
    ("Régulier", "regular_hours"),
    ("Suppl.", "suppl_hours"),
    ("Congé", "conge_hours"),
    ("Congé Travaillé", "conge_trav_hours"),
    ("Maladie", "maladie_hours"),
]


# ============================================================
# GENERIC HELPERS
# ============================================================
def normalize_text(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.replace("\u00A0", " ")
    s = " ".join(s.split())
    return s


def get_query_params_compat() -> dict:
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
        try:
            qp = st.experimental_get_query_params()
            out = {}
            for k, v in qp.items():
                if isinstance(v, list):
                    out[k] = v[0] if v else ""
                else:
                    out[k] = str(v) if v is not None else ""
            return out
        except Exception:
            return {}


def clear_query_params_compat():
    try:
        st.query_params.clear()
    except Exception:
        try:
            st.experimental_set_query_params()
        except Exception:
            pass


def format_money(x) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"


def safe_text_series(s: pd.Series) -> pd.Series:
    out = s.astype(str).str.replace("\u00A0", " ", regex=False).str.strip()
    return out.replace(
        {
            "nan": "",
            "None": "",
            "none": "",
            "NULL": "",
            "null": "",
            "<NA>": "",
        }
    ).fillna("")


# ============================================================
# MSAL / AUTH
# ============================================================
def get_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        token_cache=None,
    )


def get_me(access_token: str) -> dict:
    r = requests.get(
        "https://graph.microsoft.com/v1.0/me",
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=60,
    )
    if r.status_code != 200:
        raise RuntimeError(f"Graph /me error {r.status_code}\n{r.text}")
    return r.json()


def get_user_email(me: dict) -> str:
    return (me.get("mail") or me.get("userPrincipalName") or "").strip().lower()


def is_allowed_user(me: dict) -> bool:
    email = get_user_email(me)
    if not ALLOWED_DOMAIN:
        return False
    return email.endswith(f"@{ALLOWED_DOMAIN}")


# ============================================================
# GRAPH / ONEDRIVE HELPERS
# ============================================================
def graph_get(url: str, access_token: str) -> requests.Response:
    return requests.get(
        url,
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=60,
    )


def graph_get_json(url: str, access_token: str) -> dict:
    r = graph_get(url, access_token)
    if r.status_code != 200:
        raise RuntimeError(f"Graph error {r.status_code}\n{r.text}")
    return r.json()


def make_share_id(shared_url: str) -> str:
    b = base64.b64encode(shared_url.encode("utf-8")).decode("utf-8")
    b = b.rstrip("=").replace("/", "_").replace("+", "-")
    return "u!" + b


def resolve_shared_link(access_token: str, shared_url: str) -> dict:
    share_id = make_share_id(shared_url)
    meta_url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem"
    meta = graph_get(meta_url, access_token)
    if meta.status_code != 200:
        raise RuntimeError(
            f"Error resolving shared link: {meta.status_code}\n{meta.text}\n\n"
            "TIP: Use SharePoint/OneDrive → Share → Copy link (within your organization)."
        )
    return meta.json()


def download_item_bytes(access_token: str, drive_id: str, item_id: str) -> bytes:
    content_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    r = graph_get(content_url, access_token)
    if r.status_code != 200:
        raise RuntimeError(f"Error downloading file: {r.status_code}\n{r.text}")
    return r.content


def list_children_all(access_token: str, drive_id: str, folder_item_id: str):
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_item_id}/children?$top=200"
    all_items = []
    while url:
        data = graph_get_json(url, access_token)
        all_items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return all_items


def is_excel_name(name: str) -> bool:
    n = (name or "").lower()
    return n.endswith(".xlsx") or n.endswith(".xlsm") or n.endswith(".xls")


def is_folder_item(item: dict) -> bool:
    return "folder" in item


# ============================================================
# DATA LOADING / MAPPING
# ============================================================
def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    return df


def pick_column(df: pd.DataFrame, candidates: list[str], fallback_idx: int | None = None):
    norm_cols = {normalize_text(c): c for c in df.columns}

    for cand in candidates:
        cand_norm = normalize_text(cand)
        if cand_norm in norm_cols:
            return norm_cols[cand_norm]

    for cand in candidates:
        cand_norm = normalize_text(cand)
        for col in df.columns:
            if cand_norm in normalize_text(col):
                return col

    if fallback_idx is not None and fallback_idx < len(df.columns):
        return df.columns[fallback_idx]

    return None


def load_selected_excel_files(access_token: str, drive_id: str, selected_files: list[dict], month_name_map: dict) -> pd.DataFrame:
    dfs = []

    for file_info in selected_files:
        try:
            file_bytes = download_item_bytes(access_token, drive_id, file_info["id"])
            excel_file = pd.ExcelFile(BytesIO(file_bytes))

            sheet_to_use = None
            for s in excel_file.sheet_names:
                if normalize_text(s) == "data":
                    sheet_to_use = s
                    break

            if sheet_to_use is None:
                st.warning(
                    f"Could not read {file_info['name']}: sheet 'Data' not found. "
                    f"Available sheets: {excel_file.sheet_names}"
                )
                continue

            df = excel_file.parse(sheet_to_use)
            df = clean_columns(df)
            df["source_file"] = file_info["name"]
            df["source_month_folder"] = month_name_map.get(file_info["id"], "")
            dfs.append(df)

        except Exception as e:
            st.warning(f"Could not read {file_info['name']}: {e}")

    if dfs:
        return pd.concat(dfs, ignore_index=True)

    return pd.DataFrame()


def build_required_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    Expected layout from your sheet:
    H = total hours worked (number)
    I = hourly rate
    J = hours Suppl
    K = hours Congé
    L = hours Congé Travaillé
    M = Hours Maladie
    N = Total to pay
    R = type of work
    S = Vendor Company
    T = Employee
    """
    out = pd.DataFrame()

    date_col = pick_column(df, ["date"], fallback_idx=4)
    province_col = pick_column(df, ["province"], fallback_idx=6)
    hours_col = pick_column(df, ["total hours worked (number)", "total hours worked"], fallback_idx=7)
    hourly_rate_col = pick_column(df, ["hourly rate", "hourly_rate"], fallback_idx=8)

    suppl_col = pick_column(df, ["hours suppl"], fallback_idx=9)
    conge_col = pick_column(df, ["hours congé", "hours conge"], fallback_idx=10)
    conge_trav_col = pick_column(df, ["hours congé travaillé", "hours conge travaille"], fallback_idx=11)
    maladie_col = pick_column(df, ["hours maladie", "maladie"], fallback_idx=12)

    total_pay_col = pick_column(df, ["total to pay", "total_pay"], fallback_idx=13)
    type_work_col = pick_column(df, ["type of work", "type_of_work"], fallback_idx=17)
    vendor_col = pick_column(df, ["vendor company", "vendor_company"], fallback_idx=18)
    employee_col = pick_column(df, ["employee"], fallback_idx=19)

    col_debug = {
        "date": date_col,
        "province": province_col,
        "hours": hours_col,
        "hourly_rate": hourly_rate_col,
        "hours_suppl": suppl_col,
        "hours_conge": conge_col,
        "hours_conge_travaille": conge_trav_col,
        "hours_maladie": maladie_col,
        "total_pay": total_pay_col,
        "type_of_work": type_work_col,
        "vendor_company": vendor_col,
        "employee": employee_col,
    }

    def col_or_blank(col_name):
        if col_name is None:
            return pd.Series([""] * len(df))
        data = df[col_name]
        if isinstance(data, pd.DataFrame):
            return data.bfill(axis=1).iloc[:, 0]
        return data

    out["date"] = col_or_blank(date_col)
    out["province"] = col_or_blank(province_col)
    out["hours"] = col_or_blank(hours_col)
    out["hourly_rate"] = col_or_blank(hourly_rate_col)
    out["suppl_raw"] = col_or_blank(suppl_col)
    out["conge_raw"] = col_or_blank(conge_col)
    out["conge_trav_raw"] = col_or_blank(conge_trav_col)
    out["maladie_raw"] = col_or_blank(maladie_col)
    out["total_pay"] = col_or_blank(total_pay_col)
    out["type_of_work"] = col_or_blank(type_work_col)
    out["vendor_company"] = col_or_blank(vendor_col)
    out["employee"] = col_or_blank(employee_col)

    out["source_file"] = df["source_file"] if "source_file" in df.columns else ""
    out["source_month_folder"] = df["source_month_folder"] if "source_month_folder" in df.columns else ""

    return out, col_debug


# ============================================================
# COMMITTEE LOGIC
# ============================================================
def assign_committee_week(date_value: pd.Timestamp, start_date: pd.Timestamp, num_weeks: int = 6):
    for i in range(num_weeks):
        week_start = start_date + timedelta(days=i * 7)
        week_end = week_start + timedelta(days=6)
        if week_start <= date_value <= week_end:
            return week_start, week_end
    return None, None


def calculate_weekly_committee_hours(total_pay: float, raw_hours_sum: float, hourly_rate_min: float) -> float:
    try:
        total_pay = float(total_pay)
    except Exception:
        total_pay = 0.0

    try:
        raw_hours_sum = float(raw_hours_sum)
    except Exception:
        raw_hours_sum = 0.0

    try:
        hourly_rate_min = float(hourly_rate_min)
    except Exception:
        hourly_rate_min = 0.0

    # Flat case / lower than Class A:
    # use weekly total pay / 21.57
    if hourly_rate_min > 0 and hourly_rate_min < COMITE_CLASS_A_RATE:
        return round(total_pay / COMITE_CLASS_A_RATE, 2)

    # Flat weekly case with only ~1 hour but high payment
    if raw_hours_sum <= 1 and total_pay > COMITE_CLASS_A_RATE:
        return round(total_pay / COMITE_CLASS_A_RATE, 2)

    # Otherwise keep real summed hours
    return round(raw_hours_sum, 2)


def build_weekly_summary(df: pd.DataFrame) -> pd.DataFrame:
    grouped = (
        df.groupby(["vendor_company", "employee", "week_label"], dropna=False)
        .agg(
            raw_hours_sum=("hours", "sum"),
            suppl_hours_raw=("suppl_raw", "sum"),
            conge_hours_raw=("conge_raw", "sum"),
            conge_trav_hours_raw=("conge_trav_raw", "sum"),
            maladie_hours_raw=("maladie_raw", "sum"),
            total_pay=("total_pay", "sum"),
            hourly_rate_min=("hourly_rate", "min"),
            source_month_folder=("source_month_folder", "first"),
            source_file=("source_file", "first"),
            province=("province", "first"),
        )
        .reset_index()
        .sort_values(["vendor_company", "employee", "week_label"])
    )

    grouped["committee_hours"] = grouped.apply(
        lambda r: calculate_weekly_committee_hours(
            r["total_pay"],
            r["raw_hours_sum"],
            r["hourly_rate_min"],
        ),
        axis=1,
    )

    grouped["suppl_hours"] = pd.to_numeric(grouped["suppl_hours_raw"], errors="coerce").fillna(0)
    grouped["conge_hours"] = pd.to_numeric(grouped["conge_hours_raw"], errors="coerce").fillna(0)
    grouped["conge_trav_hours"] = pd.to_numeric(grouped["conge_trav_hours_raw"], errors="coerce").fillna(0)
    grouped["maladie_hours"] = pd.to_numeric(grouped["maladie_hours_raw"], errors="coerce").fillna(0)

    special_total = (
        grouped["suppl_hours"] +
        grouped["conge_hours"] +
        grouped["conge_trav_hours"] +
        grouped["maladie_hours"]
    )

    grouped["regular_hours"] = (grouped["committee_hours"] - special_total).clip(lower=0)
    grouped["reer"] = grouped["committee_hours"] * REER_PER_HOUR
    grouped["total_with_reer"] = grouped["total_pay"] + grouped["reer"]

    grouped["regular_hours"] = (grouped["committee_hours"] - special_total).clip(lower=0)
    grouped["reer"] = grouped["committee_hours"] * REER_PER_HOUR
    grouped["total_with_reer"] = grouped["total_pay"] + grouped["reer"]

    numeric_cols = grouped.select_dtypes(include=["number"]).columns
    grouped[numeric_cols] = grouped[numeric_cols].round(2)

    return grouped


# ============================================================
# EXPORT HELPERS
# ============================================================
def create_employee_report_blocks(weekly_df: pd.DataFrame, vendor_company: str):
    report_data = []

    vendor_df = weekly_df[weekly_df["vendor_company"] == vendor_company].copy()
    employees = sorted(vendor_df["employee"].dropna().unique().tolist())
    week_labels = sorted(vendor_df["week_label"].dropna().unique().tolist())

    for employee in employees:
        emp_df = vendor_df[vendor_df["employee"] == employee].copy()

        block = {
            "employee": employee,
            "vendor_company": vendor_company,
            "weeks": week_labels,
            "rows": [],
            "week_pay_totals": [],
            "total_hours": 0.0,
            "total_pay": 0.0,
            "reer_amount": 0.0,
            "total_with_reer": 0.0,
        }

        total_hours_employee = 0.0
        total_pay_employee = 0.0

        for label, col_name in ROW_ORDER:
            row_values = []
            row_total = 0.0

            for wk in week_labels:
                val = emp_df.loc[emp_df["week_label"] == wk, col_name].sum()
                val = round(float(val), 2)
                row_values.append(val)
                row_total += val

            total_hours_employee += row_total
            block["rows"].append(
                {
                    "label": label,
                    "week_values": row_values,
                    "row_total": round(row_total, 2),
                }
            )

        for wk in week_labels:
            pay_val = emp_df.loc[emp_df["week_label"] == wk, "total_pay"].sum()
            block["week_pay_totals"].append(round(float(pay_val), 2))
            total_pay_employee += float(pay_val)

        block["total_hours"] = round(total_hours_employee, 2)
        block["total_pay"] = round(total_pay_employee, 2)
        block["reer_amount"] = round(total_hours_employee * REER_PER_HOUR, 2)
        block["total_with_reer"] = round(block["total_pay"] + block["reer_amount"], 2)

        report_data.append(block)

    return report_data


def export_committee_report(weekly_df: pd.DataFrame, start_date_value) -> BytesIO:
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        title_fmt = workbook.add_format({
            "bold": True, "font_size": 16, "align": "center", "valign": "vcenter",
            "border": 1, "bg_color": "#1F4E78", "font_color": "white"
        })
        subtitle_fmt = workbook.add_format({
            "bold": True, "font_size": 11, "align": "center", "valign": "vcenter",
            "border": 1, "bg_color": "#D9EAF7"
        })
        section_fmt = workbook.add_format({
            "bold": True, "align": "center", "valign": "vcenter",
            "border": 1, "bg_color": "#EDEDED"
        })
        header_fmt = workbook.add_format({
            "bold": True, "align": "center", "valign": "vcenter",
            "border": 1, "bg_color": "#DCE6F1"
        })
        row_label_fmt = workbook.add_format({
            "border": 1, "align": "left", "valign": "vcenter"
        })
        num_fmt = workbook.add_format({
            "border": 1, "align": "right", "valign": "vcenter", "num_format": "0.00"
        })
        money_fmt = workbook.add_format({
            "border": 1, "align": "right", "valign": "vcenter", "num_format": "$#,##0.00"
        })
        total_fmt = workbook.add_format({
            "bold": True, "border": 1, "align": "right", "valign": "vcenter",
            "bg_color": "#FFF2CC", "num_format": "0.00"
        })
        total_money_fmt = workbook.add_format({
            "bold": True, "border": 1, "align": "right", "valign": "vcenter",
            "bg_color": "#FFF2CC", "num_format": "$#,##0.00"
        })
        emp_fmt = workbook.add_format({
            "bold": True, "border": 1, "align": "left", "valign": "vcenter",
            "bg_color": "#FCE4D6"
        })

        vendors = sorted(weekly_df["vendor_company"].dropna().unique().tolist())
        summary_rows = []

        for vendor in vendors:
            sheet_name = vendor[:31] if vendor else "Committee Report"
            ws = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = ws

            ws.set_column("A:A", 6)
            ws.set_column("B:B", 35)
            ws.set_column("C:Z", 12)
            ws.set_column("AA:AC", 16)

            vendor_df = weekly_df[weekly_df["vendor_company"] == vendor].copy()
            report_blocks = create_employee_report_blocks(weekly_df, vendor)
            week_labels = sorted(vendor_df["week_label"].dropna().unique().tolist())

            row = 0
            ws.merge_range(row, 0, row, 10, f"Versement de {pd.to_datetime(start_date_value).strftime('%b. %Y')}", title_fmt)
            row += 1
            ws.merge_range(row, 0, row, 10, vendor, subtitle_fmt)
            row += 2

            grand_total_reer = 0.0
            grand_total_pay = 0.0
            grand_total_with_reer = 0.0

            for idx, block in enumerate(report_blocks, start=1):
                ws.write(row, 0, idx, section_fmt)
                ws.write(row, 1, block["employee"], emp_fmt)
                row += 1

                ws.write(row, 0, "Classes", header_fmt)

                col = 1
                for wk in week_labels:
                    ws.merge_range(row, col, row, col + 1, wk, header_fmt)
                    col += 2

                ws.write(row, col, "Cal. Heures", header_fmt)
                row += 1

                ws.write(row, 0, "", header_fmt)
                col = 1
                for _ in week_labels:
                    ws.write(row, col, "A", header_fmt)
                    ws.write(row, col + 1, "B", header_fmt)
                    col += 2
                ws.write(row, col, "REER", header_fmt)
                row += 1

                for class_row in block["rows"]:
                    ws.write(row, 0, class_row["label"], row_label_fmt)

                    col = 1
                    for val in class_row["week_values"]:
                        ws.write_number(row, col, val, num_fmt)
                        ws.write_number(row, col + 1, 0, num_fmt)
                        col += 2

                    ws.write_number(row, col, class_row["row_total"], num_fmt)
                    row += 1

                ws.write(row, 0, "", row_label_fmt)
                col = 1
                for pay_val in block["week_pay_totals"]:
                    ws.write_number(row, col, pay_val, money_fmt)
                    ws.write_number(row, col + 1, 0, money_fmt)
                    col += 2
                ws.write_number(row, col, block["total_hours"], total_fmt)
                row += 1

                ws.write(row, 1, "Total gains", row_label_fmt)
                ws.write_number(row, 2, block["total_pay"], total_money_fmt)
                row += 1

                ws.write(row, 1, "REER", row_label_fmt)
                ws.write_number(row, 2, REER_PER_HOUR, num_fmt)
                ws.write_number(row, 3, block["reer_amount"], total_money_fmt)
                row += 1

                ws.write(row, 1, "Total gains including REER", row_label_fmt)
                ws.write_number(row, 2, block["total_with_reer"], total_money_fmt)
                row += 2

                grand_total_reer += block["reer_amount"]
                grand_total_pay += block["total_pay"]
                grand_total_with_reer += block["total_with_reer"]

                summary_rows.append({
                    "Vendor Company": vendor,
                    "Employee": block["employee"],
                    "Total Hours": block["total_hours"],
                    "Total Pay": block["total_pay"],
                    "REER": block["reer_amount"],
                    "Total with REER": block["total_with_reer"],
                })

            top_col = 13
            ws.write(0, top_col, "TOTAL REER DE TOUTES LES EMPLOYÉS", header_fmt)
            ws.write_number(0, top_col + 1, grand_total_reer, total_money_fmt)

            ws.write(1, top_col, "TOTAL DES GAINS DE TOUTES LES EMPLOYÉS INCLUANT MONTANTS REER", header_fmt)
            ws.write_number(1, top_col + 1, grand_total_with_reer, total_money_fmt)

            levy = round(grand_total_with_reer * 0.01, 2)
            ws.write(2, top_col, "X 1% (EMPLOYEUR ET EMPLOYÉS)", header_fmt)
            ws.write_number(2, top_col + 1, levy, total_money_fmt)

            ws.write(3, top_col, "AJUST. MOIS PRÉCÉDENTS", header_fmt)
            ws.write_number(3, top_col + 1, 0, total_money_fmt)

            ws.write(4, top_col, "PRÉLÈVEMENT TOTAL DÛ", header_fmt)
            ws.write_number(4, top_col + 1, levy, total_money_fmt)

        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary", index=False)

    output.seek(0)
    return output


# ============================================================
# AUTH FLOW
# ============================================================
if not ALLOWED_DOMAIN:
    st.error("Missing ALLOWED_DOMAIN in Streamlit secrets.")
    st.stop()

app = get_msal_app()
params = get_query_params_compat()

if "token_result" not in st.session_state:
    code = params.get("code")

    if code:
        result = app.acquire_token_by_authorization_code(
            code=code,
            scopes=SCOPES,
            redirect_uri=REDIRECT_URI,
        )

        if "access_token" in result:
            st.session_state.token_result = result
            clear_query_params_compat()
            st.rerun()
        else:
            st.error("Could not obtain access token.")
            st.code(str(result))
            st.stop()

    st.warning("You are not signed in to Microsoft 365 / SharePoint.")

    extra_qp = {}
    if DOMAIN_HINT:
        extra_qp["domain_hint"] = DOMAIN_HINT
    if LOGIN_HINT:
        extra_qp["login_hint"] = LOGIN_HINT

    auth_url = app.get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
        prompt="select_account",
        response_mode="query",
        extra_query_parameters=extra_qp,
    )

    st.link_button("🔐 Sign in with Microsoft (Company)", auth_url)
    st.caption(f"Redirect URI used: {REDIRECT_URI}")
    st.stop()

token_result = st.session_state.token_result
access_token = token_result.get("access_token", "")

if not access_token:
    st.error("No access token found. Please sign in again.")
    st.session_state.pop("token_result", None)
    st.stop()

try:
    me = get_me(access_token)
    signed_in_email = get_user_email(me)
except Exception as e:
    st.error("Could not validate signed-in user.")
    st.code(str(e))
    st.session_state.pop("token_result", None)
    st.stop()

if not is_allowed_user(me):
    st.error("Access denied. This dashboard is restricted to company users only.")
    st.write("Signed in as:", signed_in_email if signed_in_email else "(unknown user)")
    st.session_state.pop("token_result", None)
    st.stop()


# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.success(f"Logged in as {signed_in_email}")
st.success(f"✅ Signed in as {signed_in_email}")

if st.button("🚪 Sign out"):
    st.session_state.pop("token_result", None)
    clear_query_params_compat()
    st.rerun()

st.sidebar.header("📁 SharePoint Source")
st.sidebar.caption("Select month folder(s), then choose Excel files inside them.")


# ============================================================
# RESOLVE ROOT FOLDER
# ============================================================
try:
    meta = resolve_shared_link(access_token, ONEDRIVE_FOLDER_URL)
except Exception as e:
    st.error("Could not resolve the SharePoint/OneDrive folder link.")
    st.code(str(e))
    st.stop()

drive_id = meta["parentReference"]["driveId"]
root_item_id = meta["id"]

if "folder" not in meta:
    st.error("ONEDRIVE_FOLDER_URL must be a folder link.")
    st.stop()


# ============================================================
# LIST MONTH FOLDERS
# ============================================================
try:
    root_children = list_children_all(access_token, drive_id, root_item_id)
except Exception as e:
    st.error("Could not list folders from the root folder.")
    st.code(str(e))
    st.stop()

month_folders = [x for x in root_children if is_folder_item(x)]
month_folders.sort(key=lambda x: normalize_text(x.get("name", "")))

if not month_folders:
    st.warning("No month folders were found inside the selected root folder.")
    st.stop()

month_folder_names = [f["name"] for f in month_folders]

selected_month_names = st.sidebar.multiselect(
    "Select month folder(s)",
    month_folder_names,
    default=month_folder_names[:1] if month_folder_names else [],
)

if not selected_month_names:
    st.info("Please select at least one month folder.")
    st.stop()

selected_month_folders = [f for f in month_folders if f["name"] in selected_month_names]


# ============================================================
# LIST EXCEL FILES INSIDE SELECTED MONTH FOLDERS
# ============================================================
all_excel_files = []
month_name_map = {}

for folder_info in selected_month_folders:
    folder_id = folder_info["id"]
    folder_name = folder_info["name"]

    try:
        children = list_children_all(access_token, drive_id, folder_id)
    except Exception as e:
        st.warning(f"Could not list files inside '{folder_name}': {e}")
        continue

    excels = [x for x in children if is_excel_name(x.get("name", ""))]
    excels.sort(key=lambda x: normalize_text(x.get("name", "")))

    for item in excels:
        item_copy = dict(item)
        display_name = f"{folder_name} | {item_copy['name']}"
        item_copy["display_name"] = display_name
        all_excel_files.append(item_copy)
        month_name_map[item_copy["id"]] = folder_name

all_excel_files.sort(key=lambda x: normalize_text(x["display_name"]))

if not all_excel_files:
    st.warning("No Excel files found inside the selected month folder(s).")
    st.stop()

excel_display_names = [f["display_name"] for f in all_excel_files]

selected_excel_display_names = st.sidebar.multiselect(
    "Select Excel file(s)",
    excel_display_names,
    default=excel_display_names,
)

if not selected_excel_display_names:
    st.info("Please select at least one Excel file.")
    st.stop()

selected_excel_files = [f for f in all_excel_files if f["display_name"] in selected_excel_display_names]


# ============================================================
# LOAD EXCEL DATA
# ============================================================
raw_df = load_selected_excel_files(access_token, drive_id, selected_excel_files, month_name_map)

if raw_df.empty:
    st.error("No valid data could be loaded from the selected Excel files.")
    st.stop()


# ============================================================
# BUILD FINAL DATAFRAME
# ============================================================
df, col_debug = build_required_dataframe(raw_df)

required_cols = [
    "date", "province", "employee", "hours",
    "total_pay", "type_of_work", "vendor_company"
]
missing = [c for c in required_cols if c not in df.columns]

if missing:
    st.error(f"Missing required columns: {missing}")
    st.write("Detected columns:", list(raw_df.columns))
    st.stop()


# ============================================================
# DATA CLEANING
# ============================================================
df["date"] = pd.to_datetime(df["date"], errors="coerce")
df["province"] = safe_text_series(df["province"]).str.upper()
df["employee"] = safe_text_series(df["employee"])
df["vendor_company"] = safe_text_series(df["vendor_company"])
df["type_of_work"] = safe_text_series(df["type_of_work"])

for c in [
    "hours", "hourly_rate", "suppl_raw", "conge_raw",
    "conge_trav_raw", "maladie_raw", "total_pay"
]:
    df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

df = df.dropna(subset=["date"]).copy()


# ============================================================
# FILTERS
# ============================================================
st.sidebar.header("🧷 Committee Filters")

all_provinces = sorted([p for p in df["province"].dropna().unique().tolist() if p])
selected_provinces = st.sidebar.multiselect("Province", all_provinces, default=all_provinces)

all_vendors = sorted([v for v in df["vendor_company"].dropna().unique().tolist() if v])
selected_vendors = st.sidebar.multiselect("Vendor Company", all_vendors, default=all_vendors)

all_employees = sorted([e for e in df["employee"].dropna().unique().tolist() if e])
selected_employees = st.sidebar.multiselect("Name Employee", all_employees, default=all_employees)

all_types = sorted([t for t in df["type_of_work"].dropna().unique().tolist() if t])
selected_types = st.sidebar.multiselect("Type of work", all_types, default=all_types)

default_start = datetime(2026, 1, 4).date()
start_date = st.sidebar.date_input("First committee week start date", value=default_start)
num_weeks = st.sidebar.number_input("Number of weeks", min_value=1, max_value=12, value=4)


# ============================================================
# APPLY FILTERS
# ============================================================
filtered_df = df.copy()

if selected_provinces:
    filtered_df = filtered_df[filtered_df["province"].isin(selected_provinces)]
if selected_vendors:
    filtered_df = filtered_df[filtered_df["vendor_company"].isin(selected_vendors)]
if selected_employees:
    filtered_df = filtered_df[filtered_df["employee"].isin(selected_employees)]
if selected_types:
    filtered_df = filtered_df[filtered_df["type_of_work"].isin(selected_types)]

start_date_dt = pd.to_datetime(start_date)

filtered_df[["week_start", "week_end"]] = filtered_df["date"].apply(
    lambda x: pd.Series(assign_committee_week(x, start_date_dt, num_weeks))
)

filtered_df = filtered_df[filtered_df["week_start"].notna()].copy()
filtered_df["week_label"] = filtered_df["week_end"].dt.strftime("%Y-%m-%d")

if filtered_df.empty:
    st.warning("No data available for the selected filters.")
    st.stop()


# ============================================================
# BUILD WEEKLY SUMMARY
# ============================================================
weekly_summary = build_weekly_summary(filtered_df)


# ============================================================
# METRICS
# ============================================================
col1, col2, col3, col4 = st.columns(4)
col1.metric("Rows", len(filtered_df))
col2.metric("Employees", weekly_summary["employee"].nunique())
col3.metric("Committee Hours", f"{weekly_summary['committee_hours'].sum():,.2f}")
col4.metric("Total Pay", format_money(weekly_summary["total_pay"].sum()))


# ============================================================
# PREVIEW SOURCE
# ============================================================
preview_cols = [
    "source_month_folder",
    "source_file",
    "date",
    "employee",
    "province",
    "vendor_company",
    "type_of_work",
    "hours",
    "hourly_rate",
    "suppl_raw",
    "conge_raw",
    "conge_trav_raw",
    "maladie_raw",
    "total_pay",
    "week_label",
]
preview_cols = [c for c in preview_cols if c in filtered_df.columns]

st.subheader("Filtered source data")

display_source = filtered_df[preview_cols].copy()
numeric_cols = display_source.select_dtypes(include=["number"]).columns
for col in numeric_cols:
    display_source[col] = display_source[col].round(2)

st.dataframe(display_source, use_container_width=True)


# ============================================================
# WEEKLY EMPLOYEE SUMMARY
# ============================================================
st.subheader("Employee summary")

display_summary = weekly_summary.copy()

numeric_cols = display_summary.select_dtypes(include=["number"]).columns
for col in numeric_cols:
    display_summary[col] = display_summary[col].round(2)

st.dataframe(display_summary, use_container_width=True)


# ============================================================
# EXPORT
# ============================================================
report_file = export_committee_report(weekly_summary, start_date)

st.download_button(
    label="Download Comité Excel Report",
    data=report_file,
    file_name="comite_paritaire_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)


# ============================================================
# DIAGNOSTICS
# ============================================================
with st.expander("Diagnostics", expanded=False):
    st.write("Redirect URI used:", REDIRECT_URI)
    st.write("Allowed domain:", ALLOWED_DOMAIN)
    st.write("Signed in as:", signed_in_email)
    st.write("Root folder resolved name:", meta.get("name"))
    st.write("Selected month folders:", selected_month_names)
    st.write("Selected excel display names:", selected_excel_display_names)
    st.write("Column mapping used:", col_debug)
    st.write("Final source columns:", list(filtered_df.columns))
    st.write("Weekly summary columns:", list(weekly_summary.columns))
