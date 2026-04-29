import base64
import os
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
st.set_page_config(page_title="CNET Regular Hours Report", layout="wide")

LOGO_PATH = "cnet_logo.png"

top_left, top_right = st.columns([1, 4])
with top_left:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=220)
with top_right:
    st.title("CNET Regular Hours Report")

# ============================================================
# CONFIG - SECRETS
# ============================================================
CLIENT_ID = str(st.secrets["CLIENT_ID"]).strip()
CLIENT_SECRET = str(st.secrets["CLIENT_SECRET"]).strip()
TENANT_ID = str(st.secrets["TENANT_ID"]).strip()
REDIRECT_URI = str(st.secrets["REDIRECT_URI"]).strip().rstrip("/")

ONEDRIVE_FOLDER_URL = str(
    st.secrets.get(
        "ONEDRIVE_FOLDER_URL",
        "https://groupcastillo.sharepoint.com/:f:/s/GroupCastilloTeamSite/IgDJ46w1V3YWT7e0yB8CKkD9AenZh0xzbn8pNRRGuDcIpPw?e=s4L0Z9",
    )
).strip()

ALLOWED_DOMAIN = str(st.secrets.get("ALLOWED_DOMAIN", "")).strip().lower()
DOMAIN_HINT = str(st.secrets.get("DOMAIN_HINT", "")).strip()
LOGIN_HINT = str(st.secrets.get("LOGIN_HINT", "")).strip()

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.Read.All", "Sites.Read.All"]

# ============================================================
# BUSINESS RULES
# ============================================================
SHEET_DATA = "DATA"
TYPE_OF_WORK_DEFAULT = "REGULAR"
DEFAULT_CLASS_WHEN_NO_CLASS = "Class A"

FIRST_COMMITTEE_WEEK_START_DEFAULT = pd.Timestamp("2026-01-04")
REER_PER_HOUR = 0.45
OVERTIME_WEEKLY_THRESHOLD = 40.0
OVERTIME_MULTIPLIER = 1.5

# DATA sheet layout
COL_VENDOR_COMPANY = 0       # A
COL_EMPLOYEE_NAME = 1        # B
COL_EMPLOYEE_CLASS = 8       # I
COL_WEEK_RANGE = 10          # K, example: 03/01 - 09/01
COL_RATE = 19                # T

DAY_COL_START = 11           # L
DAY_COL_END_EXCLUSIVE = 18   # R included
DAY_HEADER_ROW = 3           # Excel row 4
DATA_START_ROW = 4           # Excel row 5

# INPUT / IMPUT sheet layout
INPUT_COL_EMPLOYEE_NAME = 1   # B
INPUT_COL_EMPLOYEE_CLASS = 8  # I
INPUT_COL_DATE = 11           # L = FECHA
INPUT_COL_V = 12              # M = V
INPUT_COL_SD = 13             # N = SD
INPUT_COL_H = 14              # O = H

ROW_ORDER = [
    ("Régulier", "regular_hours"),
    ("Overtime", "overtime_hours"),
    ("Suppl.", "suppl_hours"),
    ("Congé", "conge_hours"),
    ("Congé Travaillé", "conge_trav_hours"),
    ("Maladie", "maladie_hours"),
]

DAY_MAP = {
    "SA": "Saturday", "SAT": "Saturday", "SATURDAY": "Saturday", "SAM": "Saturday", "SAMEDI": "Saturday",
    "SU": "Sunday", "SUN": "Sunday", "SUNDAY": "Sunday", "DIM": "Sunday", "DIMANCHE": "Sunday",
    "MO": "Monday", "MON": "Monday", "MONDAY": "Monday", "LUN": "Monday", "LUNDI": "Monday",
    "TU": "Tuesday", "TUE": "Tuesday", "TUESDAY": "Tuesday", "MAR": "Tuesday", "MARDI": "Tuesday",
    "WE": "Wednesday", "WED": "Wednesday", "WEDNESDAY": "Wednesday", "MER": "Wednesday", "MERCREDI": "Wednesday",
    "TH": "Thursday", "THU": "Thursday", "THURSDAY": "Thursday", "JEU": "Thursday", "JEUDI": "Thursday",
    "FR": "Friday", "FRI": "Friday", "FRIDAY": "Friday", "VEN": "Friday", "VENDREDI": "Friday",
}

# ============================================================
# HELPERS
# ============================================================
def normalize_text(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.replace("\u00A0", " ")
    s = " ".join(s.split())
    return s


def clean_text(x) -> str:
    if pd.isna(x):
        return ""
    return " ".join(str(x).replace("\u00A0", " ").strip().split())


def safe_text_series(s: pd.Series) -> pd.Series:
    out = s.astype(str).str.replace("\u00A0", " ", regex=False).str.strip()
    return out.replace(
        {"nan": "", "None": "", "none": "", "NULL": "", "null": "", "<NA>": ""}
    ).fillna("")


def to_num_series(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0.0)

    cleaned = (
        s.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def format_money(x) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"


def dataframe_with_2_decimals(df: pd.DataFrame):
    numeric_cols = df.select_dtypes(include=["number"]).columns
    column_config = {
        col: st.column_config.NumberColumn(col, format="%.2f")
        for col in numeric_cols
    }
    st.dataframe(df, use_container_width=True, column_config=column_config)


def normalize_class(x) -> str:
    txt = clean_text(x)
    if txt == "" or normalize_text(txt) == "no class":
        return DEFAULT_CLASS_WHEN_NO_CLASS
    return txt


def parse_week_range_cell(value, fallback_year=2026):
    txt = clean_text(value)
    if not txt:
        return pd.NaT, pd.NaT

    txt = txt.replace("–", "-").replace("—", "-")
    parts = [p.strip() for p in txt.split("-")]

    if len(parts) < 2:
        return pd.NaT, pd.NaT

    start_txt = parts[0]
    end_txt = parts[1]

    start_dt = pd.to_datetime(
        f"{start_txt}/{fallback_year}",
        format="%d/%m/%Y",
        errors="coerce",
    )
    end_dt = pd.to_datetime(
        f"{end_txt}/{fallback_year}",
        format="%d/%m/%Y",
        errors="coerce",
    )

    return start_dt, end_dt


def parse_input_date(value, fallback_year=2026):
    """
    Robust parser for INPUT/IMPUT FECHA.
    Forces year 2026 when text like 10-Jan has no year.
    """
    if pd.isna(value):
        return pd.NaT

    if isinstance(value, (pd.Timestamp, datetime)):
        parsed = pd.Timestamp(value)
        if parsed.year < 2000:
            parsed = pd.Timestamp(year=fallback_year, month=parsed.month, day=parsed.day)
        return parsed.normalize()

    txt = clean_text(value)
    if not txt:
        return pd.NaT

    txt_norm = txt.replace("–", "-").replace("—", "-").strip()
    has_year = any(str(y) in txt_norm for y in range(2000, 2051))

    # If the FECHA cell says 10-Jan, 17-Jan, etc., force the payroll year.
    if not has_year:
        for candidate in (f"{txt_norm}-{fallback_year}", f"{txt_norm}/{fallback_year}"):
            parsed = pd.to_datetime(candidate, dayfirst=True, errors="coerce")
            if pd.notna(parsed):
                return pd.Timestamp(parsed).normalize()

    parsed = pd.to_datetime(txt_norm, dayfirst=True, errors="coerce")
    if pd.notna(parsed):
        parsed = pd.Timestamp(parsed)
        if not has_year or parsed.year < 2000:
            parsed = pd.Timestamp(year=fallback_year, month=parsed.month, day=parsed.day)
        return parsed.normalize()

    return pd.NaT
def assign_committee_week(date_value: pd.Timestamp, start_date: pd.Timestamp, num_weeks: int = 24):
    d = pd.to_datetime(date_value)

    if pd.isna(d):
        return None, None

    for i in range(num_weeks):
        week_start = start_date + timedelta(days=i * 7)
        week_end = week_start + timedelta(days=6)

        if week_start <= d <= week_end:
            return week_start, week_end

    return None, None


# ============================================================
# QUERY PARAM HELPERS
# ============================================================
def get_query_params_compat() -> dict:
    try:
        qp = st.query_params
        out = {}

        for k in qp.keys():
            v = qp.get(k)
            out[k] = v[0] if isinstance(v, list) and v else str(v) if v is not None else ""

        return out

    except Exception:
        try:
            qp = st.experimental_get_query_params()
            return {
                k: (v[0] if isinstance(v, list) and v else str(v))
                for k, v in qp.items()
            }
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
    if not ALLOWED_DOMAIN:
        return True

    email = get_user_email(me)
    return email.endswith(f"@{ALLOWED_DOMAIN}")


# ============================================================
# GRAPH / SHAREPOINT HELPERS
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
            "TIP: Use SharePoint/OneDrive → Share → Copy link within your organization."
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
    return not n.startswith("~$") and (
        n.endswith(".xlsx") or n.endswith(".xlsm") or n.endswith(".xls")
    )


def is_folder_item(item: dict) -> bool:
    return "folder" in item


# ============================================================
# SPECIAL HOURS INPUT / IMPUT
# ============================================================
def find_input_sheet_name(excel_file: pd.ExcelFile):
    for sheet_name in excel_file.sheet_names:
        if normalize_text(sheet_name) in {"input", "imput"}:
            return sheet_name
    return None


def normalize_week_range_key(value) -> str:
    """
    Normalizes FECHA / Week number text like:
    03/01 - 09/01, 03/01-09/01, 03/01 – 09/01
    so DATA column K and IMPUT column L can match exactly.
    """
    txt = clean_text(value)
    txt = txt.replace("–", "-").replace("—", "-")
    parts = [p.strip() for p in txt.split("-")]
    if len(parts) >= 2:
        return f"{parts[0]} - {parts[1]}".lower()
    return txt.lower()


# ============================================================
# SPECIAL HOURS INPUT / IMPUT
# ============================================================
def find_input_sheet_name(excel_file: pd.ExcelFile):
    for sheet_name in excel_file.sheet_names:
        if normalize_text(sheet_name) in {"input", "imput"}:
            return sheet_name
    return None


def build_special_hours_lookup(file_bytes: bytes, excel_file: pd.ExcelFile) -> dict:
    """
    Reads Input/Imput sheet.

    Expected columns:
    B = Employee
    I = Class
    L = FECHA or Week range
    M = V  -> Congé
    N = SD -> Maladie
    O = H  -> Congé Travaillé

    IMPORTANT:
    The match is employee + class + week range.
    It DOES NOT ignore class, because Antony Class A and Class B can have different hours.
    """

    sheet_name = find_input_sheet_name(excel_file)

    if sheet_name is None:
        return {
            "by_employee": {},
            "by_week_range": {},
            "sheet_found": None,
            "rows_found": 0,
        }

    try:
        input_raw = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, header=None)
    except Exception:
        return {
            "by_employee": {},
            "by_week_range": {},
            "sheet_found": sheet_name,
            "rows_found": 0,
        }

    by_employee = {}
    by_week_range = {}
    rows_found = 0

    for idx in range(1, input_raw.shape[0]):
        row = input_raw.iloc[idx]

        employee = clean_text(row.iloc[INPUT_COL_EMPLOYEE_NAME]) if len(row) > INPUT_COL_EMPLOYEE_NAME else ""
        employee_class = normalize_class(row.iloc[INPUT_COL_EMPLOYEE_CLASS]) if len(row) > INPUT_COL_EMPLOYEE_CLASS else DEFAULT_CLASS_WHEN_NO_CLASS
        fecha_raw = row.iloc[INPUT_COL_DATE] if len(row) > INPUT_COL_DATE else ""
        fecha_key = normalize_week_range_key(fecha_raw)
        date_value = parse_input_date(fecha_raw)

        if not employee:
            continue

        v_hours = pd.to_numeric(row.iloc[INPUT_COL_V], errors="coerce") if len(row) > INPUT_COL_V else 0.0
        sd_hours = pd.to_numeric(row.iloc[INPUT_COL_SD], errors="coerce") if len(row) > INPUT_COL_SD else 0.0
        h_hours = pd.to_numeric(row.iloc[INPUT_COL_H], errors="coerce") if len(row) > INPUT_COL_H else 0.0

        v_hours = float(v_hours) if pd.notna(v_hours) else 0.0
        sd_hours = float(sd_hours) if pd.notna(sd_hours) else 0.0
        h_hours = float(h_hours) if pd.notna(h_hours) else 0.0

        if v_hours == 0 and sd_hours == 0 and h_hours == 0:
            continue

        rows_found += 1

        values = {"V": v_hours, "SD": sd_hours, "H": h_hours}

        # Main matching method: Employee + Class + FECHA text / week range.
        if fecha_key:
            key_range = (
                normalize_text(employee),
                normalize_text(employee_class),
                fecha_key,
            )
            if key_range not in by_week_range:
                by_week_range[key_range] = {"V": 0.0, "SD": 0.0, "H": 0.0}
            by_week_range[key_range]["V"] += values["V"]
            by_week_range[key_range]["SD"] += values["SD"]
            by_week_range[key_range]["H"] += values["H"]

        # Backup matching method: Employee + Class + parsed date.
        if pd.notna(date_value):
            date_key = pd.Timestamp(date_value).normalize()
            key_date = (
                normalize_text(employee),
                normalize_text(employee_class),
                date_key,
            )
            if key_date not in by_employee:
                by_employee[key_date] = {"V": 0.0, "SD": 0.0, "H": 0.0}
            by_employee[key_date]["V"] += values["V"]
            by_employee[key_date]["SD"] += values["SD"]
            by_employee[key_date]["H"] += values["H"]

    return {
        "by_employee": by_employee,
        "by_week_range": by_week_range,
        "sheet_found": sheet_name,
        "rows_found": rows_found,
    }


def get_special_hours(special_lookup: dict, employee: str, employee_class: str, lookup_value, code: str) -> float:
    """
    Gets V, SD, or H hours from INPUT/IMPUT.

    Priority:
    1) employee + class + week range text, for example 03/01 - 09/01
    2) employee + class + parsed date

    It NEVER ignores class. This avoids Class A taking Class B hours.
    """
    employee_key = normalize_text(employee)
    class_key = normalize_text(employee_class)

    by_week_range = special_lookup.get("by_week_range", {})
    week_key = normalize_week_range_key(lookup_value)
    exact_range_key = (employee_key, class_key, week_key)
    if exact_range_key in by_week_range:
        return float(by_week_range[exact_range_key].get(code, 0.0))

    by_employee = special_lookup.get("by_employee", {})
    parsed_date = parse_input_date(lookup_value)
    if pd.notna(parsed_date):
        exact_date_key = (employee_key, class_key, pd.Timestamp(parsed_date).normalize())
        if exact_date_key in by_employee:
            return float(by_employee[exact_date_key].get(code, 0.0))

    return 0.0

# ============================================================
# DATA LOADING
# ============================================================
def load_selected_excel_files_regular(
    access_token: str,
    drive_id: str,
    selected_files: list[dict],
    month_name_map: dict,
) -> pd.DataFrame:

    all_rows = []
    diagnostics = []

    for file_info in selected_files:
        file_name = file_info.get("name", "")
        source_month = month_name_map.get(file_info["id"], "")

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
                    f"Could not read {file_name}: sheet DATA not found. "
                    f"Available sheets: {excel_file.sheet_names}"
                )
                continue

            special_hours_lookup = build_special_hours_lookup(file_bytes, excel_file)
            raw = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_to_use, header=None)

            day_headers = raw.iloc[DAY_HEADER_ROW, DAY_COL_START:DAY_COL_END_EXCLUSIVE].tolist()
            data = raw.iloc[DATA_START_ROW:].copy()

            diagnostics.append({
                "source_file": file_name,
                "source_month_folder": source_month,
                "sheet": sheet_to_use,
                "day_headers_L_R": [clean_text(x) for x in day_headers],
                "input_sheet_found": special_hours_lookup.get("sheet_found"),
                "input_rows_found": special_hours_lookup.get("rows_found", 0),
            })

            for _, r in data.iterrows():
                vendor = clean_text(r.iloc[COL_VENDOR_COMPANY]) if len(r) > COL_VENDOR_COMPANY else ""
                employee = clean_text(r.iloc[COL_EMPLOYEE_NAME]) if len(r) > COL_EMPLOYEE_NAME else ""
                employee_class = normalize_class(r.iloc[COL_EMPLOYEE_CLASS]) if len(r) > COL_EMPLOYEE_CLASS else DEFAULT_CLASS_WHEN_NO_CLASS
                week_range_text = clean_text(r.iloc[COL_WEEK_RANGE]) if len(r) > COL_WEEK_RANGE else ""

                week_start_from_excel, week_end_from_excel = parse_week_range_cell(
                    week_range_text,
                    fallback_year=2026,
                )

                rate = pd.to_numeric(r.iloc[COL_RATE], errors="coerce") if len(r) > COL_RATE else 0.0
                rate = float(rate) if pd.notna(rate) else 0.0

                if not vendor and not employee:
                    continue

                if pd.isna(week_start_from_excel):
                    continue

                # DATA column K must match INPUT/IMPUT column L (FECHA / Week range).
                # Example: DATA K = 03/01 - 09/01 matches IMPUT L = 03/01 - 09/01.
                special_lookup_date = week_range_text

                # Read all L:R cells for this DATA row.
                # We calculate regular numeric hours as the sum of numeric cells ONLY.
                # IMPORTANT:
                # Sometimes Excel displays V / SD / H, but pandas receives the numeric cached value.
                # In that case, regular_hours can incorrectly include the special hours.
                # To fix it:
                #   1) Always read the V / SD / H value from Input/Imput once.
                #   2) If no visible letter was detected in DATA, subtract that special value from regular_hours.
                week_values = []
                for col_idx in range(DAY_COL_START, DAY_COL_END_EXCLUSIVE):
                    cell_value = r.iloc[col_idx] if len(r) > col_idx else ""
                    week_values.append(cell_value)

                week_letters = [clean_text(v).upper() for v in week_values]
                visible_special_detected = any(letter in {"V", "SD", "H"} for letter in week_letters)

                regular_hours = 0.0
                regular_numeric_values = []

                for v in week_values:
                    txt = clean_text(v).upper()

                    # If pandas actually sees V / SD / H, never count it as regular.
                    if txt in {"V", "SD", "H"}:
                        continue

                    numeric_value = pd.to_numeric(v, errors="coerce")
                    if pd.notna(numeric_value):
                        regular_numeric_values.append(float(numeric_value))
                        regular_hours += float(numeric_value)

                suppl_hours = 0.0
                conge_hours = 0.0
                conge_trav_hours = 0.0
                maladie_hours = 0.0

                # Pull special values once from Input/Imput using employee + class + lookup date.
                input_v_hours = get_special_hours(
                    special_hours_lookup,
                    employee,
                    employee_class,
                    special_lookup_date,
                    "V",
                )
                input_h_hours = get_special_hours(
                    special_hours_lookup,
                    employee,
                    employee_class,
                    special_lookup_date,
                    "H",
                )
                input_sd_hours = get_special_hours(
                    special_hours_lookup,
                    employee,
                    employee_class,
                    special_lookup_date,
                    "SD",
                )

                # FINAL RULES:
                # INPUT column M (V)  -> Congé
                # INPUT column N (SD) -> Maladie
                # INPUT column O (H)  -> Congé Travaillé
                if "V" in week_letters or input_v_hours > 0:
                    conge_hours += input_v_hours

                if "SD" in week_letters or input_sd_hours > 0:
                    maladie_hours += input_sd_hours

                if "H" in week_letters or input_h_hours > 0:
                    conge_trav_hours += input_h_hours
                special_hours_total = conge_hours + maladie_hours + conge_trav_hours + suppl_hours

                # Critical correction:
                # If pandas did not see visible V / SD / H, but Input/Imput has special hours,
                # those special hours may already be included in the numeric cached values.
                # Subtract them from regular_hours so regular shows only real worked numeric hours.
                if not visible_special_detected and special_hours_total > 0:
                    regular_hours = max(0.0, regular_hours - special_hours_total)

                total_hours_for_week = (
                    regular_hours
                    + suppl_hours
                    + conge_hours
                    + conge_trav_hours
                    + maladie_hours
                )

                if total_hours_for_week == 0:
                    continue

                all_rows.append({
                    "source_month_folder": source_month,
                    "source_file": file_name,
                    "excel_week_range": week_range_text,
                    "excel_week_start": week_start_from_excel,
                    "excel_week_end": week_end_from_excel,
                    "special_lookup_date": special_lookup_date,
                    "date": week_start_from_excel,
                    "vendor_company": vendor,
                    "employee": employee,
                    "employee_class": employee_class,
                    "type_of_work": TYPE_OF_WORK_DEFAULT,
                    "day": "Week Total",
                    "excel_cell_value": " | ".join([clean_text(v) for v in week_values if clean_text(v)]),
                    "regular_numeric_values": " | ".join([str(x) for x in regular_numeric_values]),
                    "visible_special_detected": visible_special_detected,
                    "input_v_hours": input_v_hours,
                    "input_h_hours": input_h_hours,
                    "input_sd_hours": input_sd_hours,
                    "special_hours_total": special_hours_total,
                    "hours": total_hours_for_week,
                    "hourly_rate": rate,
                    "regular_hours": regular_hours,
                    "suppl_hours": suppl_hours,
                    "conge_hours": conge_hours,
                    "conge_trav_hours": conge_trav_hours,
                    "maladie_hours": maladie_hours,
                })

        except Exception as e:
            st.warning(f"Could not read {file_name}: {e}")

    df = pd.DataFrame(all_rows)
    st.session_state["regular_loader_diagnostics"] = diagnostics
    return df


# ============================================================
# WEEKLY SUMMARY
# ============================================================
def build_weekly_summary(df: pd.DataFrame) -> pd.DataFrame:
    grouped = (
        df.groupby(["vendor_company", "employee", "employee_class", "week_label"], dropna=False)
        .agg(
            regular_hours_original=("regular_hours", "sum"),
            suppl_hours=("suppl_hours", "sum"),
            conge_hours=("conge_hours", "sum"),
            conge_trav_hours=("conge_trav_hours", "sum"),
            maladie_hours=("maladie_hours", "sum"),
            hourly_rate=("hourly_rate", "mean"),
            source_month_folder=("source_month_folder", "first"),
            source_file=("source_file", "first"),
        )
        .reset_index()
        .sort_values(["vendor_company", "employee", "week_label"])
    )

    grouped["overtime_hours"] = (
        grouped["regular_hours_original"] - OVERTIME_WEEKLY_THRESHOLD
    ).clip(lower=0)

    grouped["regular_hours"] = grouped["regular_hours_original"].clip(
        upper=OVERTIME_WEEKLY_THRESHOLD
    )

    grouped["regular_pay"] = grouped["regular_hours"] * grouped["hourly_rate"]
    grouped["overtime_pay"] = grouped["overtime_hours"] * grouped["hourly_rate"] * OVERTIME_MULTIPLIER
    grouped["suppl_pay"] = grouped["suppl_hours"] * grouped["hourly_rate"]
    grouped["conge_pay"] = grouped["conge_hours"] * grouped["hourly_rate"]
    grouped["conge_trav_pay"] = grouped["conge_trav_hours"] * grouped["hourly_rate"]
    grouped["maladie_pay"] = grouped["maladie_hours"] * grouped["hourly_rate"]

    grouped["total_pay"] = (
        grouped["regular_pay"]
        + grouped["overtime_pay"]
        + grouped["suppl_pay"]
        + grouped["conge_pay"]
        + grouped["conge_trav_pay"]
        + grouped["maladie_pay"]
    )

    grouped["total_hours"] = (
        grouped["regular_hours"]
        + grouped["overtime_hours"]
        + grouped["suppl_hours"]
        + grouped["conge_hours"]
        + grouped["conge_trav_hours"]
        + grouped["maladie_hours"]
    )

    grouped["reer"] = grouped["total_hours"] * REER_PER_HOUR
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
        emp_class = emp_df["employee_class"].dropna().astype(str).iloc[0] if not emp_df.empty else "Class A"

        block = {
            "employee": employee,
            "employee_class": emp_class,
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
                val = emp_df.loc[emp_df["week_label"] == wk, col_name].sum() if col_name in emp_df.columns else 0.0
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


def export_regular_hours_report(weekly_df: pd.DataFrame, start_date_value) -> BytesIO:
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        title_fmt = workbook.add_format({"bold": True, "font_size": 16, "align": "center", "valign": "vcenter", "border": 1, "bg_color": "#1F4E78", "font_color": "white"})
        subtitle_fmt = workbook.add_format({"bold": True, "font_size": 11, "align": "center", "valign": "vcenter", "border": 1, "bg_color": "#D9EAF7"})
        section_fmt = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter", "border": 1, "bg_color": "#EDEDED"})
        header_fmt = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter", "border": 1, "bg_color": "#DCE6F1"})
        row_label_fmt = workbook.add_format({"border": 1, "align": "left", "valign": "vcenter"})
        num_fmt = workbook.add_format({"border": 1, "align": "right", "valign": "vcenter", "num_format": "0.00"})
        money_fmt = workbook.add_format({"border": 1, "align": "right", "valign": "vcenter", "num_format": "$#,##0.00"})
        total_fmt = workbook.add_format({"bold": True, "border": 1, "align": "right", "valign": "vcenter", "bg_color": "#FFF2CC", "num_format": "0.00"})
        total_money_fmt = workbook.add_format({"bold": True, "border": 1, "align": "right", "valign": "vcenter", "bg_color": "#FFF2CC", "num_format": "$#,##0.00"})
        emp_fmt = workbook.add_format({"bold": True, "border": 1, "align": "left", "valign": "vcenter", "bg_color": "#FCE4D6"})

        vendors = sorted(weekly_df["vendor_company"].dropna().unique().tolist())
        summary_rows = []

        for vendor in vendors:
            sheet_name = str(vendor)[:31] if vendor else "Regular Report"
            ws = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = ws

            ws.set_column("A:A", 16)
            ws.set_column("B:B", 35)
            ws.set_column("C:Z", 14)
            ws.set_column("AA:AC", 18)

            if os.path.exists(LOGO_PATH):
                ws.insert_image(0, 0, LOGO_PATH, {"x_scale": 0.35, "y_scale": 0.35})

            row = 0
            ws.merge_range(row, 2, row, 10, f"Regular Hours Report - {pd.to_datetime(start_date_value).strftime('%b. %Y')}", title_fmt)
            row += 1
            ws.merge_range(row, 2, row, 10, str(vendor), subtitle_fmt)
            row += 2

            vendor_df = weekly_df[weekly_df["vendor_company"] == vendor].copy()
            report_blocks = create_employee_report_blocks(weekly_df, vendor)
            week_labels = sorted(vendor_df["week_label"].dropna().unique().tolist())

            grand_total_reer = 0.0
            grand_total_pay = 0.0
            grand_total_with_reer = 0.0

            for idx, block in enumerate(report_blocks, start=1):
                ws.write(row, 0, idx, section_fmt)
                ws.write(row, 1, block["employee"], emp_fmt)
                ws.write(row, 2, block["employee_class"], emp_fmt)
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
                        ws.write_number(row, col, round(float(val), 2), num_fmt)
                        ws.write_number(row, col + 1, 0, num_fmt)
                        col += 2
                    ws.write_number(row, col, round(float(class_row["row_total"]), 2), num_fmt)
                    row += 1

                ws.write(row, 0, "", row_label_fmt)
                col = 1
                for pay_val in block["week_pay_totals"]:
                    ws.write_number(row, col, round(float(pay_val), 2), money_fmt)
                    ws.write_number(row, col + 1, 0, money_fmt)
                    col += 2
                ws.write_number(row, col, round(float(block["total_hours"]), 2), total_fmt)
                row += 1

                ws.write(row, 1, "Total gains", row_label_fmt)
                ws.write_number(row, 2, round(float(block["total_pay"]), 2), total_money_fmt)
                row += 1

                ws.write(row, 1, "REER", row_label_fmt)
                ws.write_number(row, 2, REER_PER_HOUR, num_fmt)
                ws.write_number(row, 3, round(float(block["reer_amount"]), 2), total_money_fmt)
                row += 1

                ws.write(row, 1, "Total gains including REER", row_label_fmt)
                ws.write_number(row, 2, round(float(block["total_with_reer"]), 2), total_money_fmt)
                row += 2

                grand_total_reer += float(block["reer_amount"])
                grand_total_pay += float(block["total_pay"])
                grand_total_with_reer += float(block["total_with_reer"])

                summary_rows.append({
                    "Vendor Company": vendor,
                    "Employee": block["employee"],
                    "Employee Class": block["employee_class"],
                    "Total Hours": round(float(block["total_hours"]), 2),
                    "Total Pay": round(float(block["total_pay"]), 2),
                    "REER": round(float(block["reer_amount"]), 2),
                    "Total with REER": round(float(block["total_with_reer"]), 2),
                })

            levy = round(grand_total_with_reer * 0.01, 2)
            prelevement_total_du_vendor = round(grand_total_with_reer + levy, 2)
            top_col = 13
            ws.write(0, top_col, "TOTAL DES GAINS", header_fmt)
            ws.write_number(0, top_col + 1, round(grand_total_pay, 2), total_money_fmt)
            ws.write(1, top_col, "TOTAL REER DE TOUS LES EMPLOYÉS", header_fmt)
            ws.write_number(1, top_col + 1, round(grand_total_reer, 2), total_money_fmt)
            ws.write(2, top_col, "TOTAL DES GAINS INCLUANT REER", header_fmt)
            ws.write_number(2, top_col + 1, round(grand_total_with_reer, 2), total_money_fmt)
            ws.write(3, top_col, "X 1% (EMPLOYEUR ET EMPLOYÉS)", header_fmt)
            ws.write_number(3, top_col + 1, levy, total_money_fmt)
            ws.write(4, top_col, "PRÉLÈVEMENT TOTAL DÛ", header_fmt)
            ws.write_number(4, top_col + 1, prelevement_total_du_vendor, total_money_fmt)

        summary_df = pd.DataFrame(summary_rows)
        if not summary_df.empty:
            numeric_cols = summary_df.select_dtypes(include=["number"]).columns
            summary_df[numeric_cols] = summary_df[numeric_cols].round(2)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

    output.seek(0)
    return output


# ============================================================
# AUTH FLOW
# ============================================================
app = get_msal_app()
params = get_query_params_compat()

if "token_result" not in st.session_state:
    code = params.get("code")

    if code:
        result = app.acquire_token_by_authorization_code(code=code, scopes=SCOPES, redirect_uri=REDIRECT_URI)

        if "access_token" in result:
            st.session_state.token_result = result
            clear_query_params_compat()
            st.rerun()

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

st.sidebar.success(f"Logged in as {signed_in_email}")
st.success(f"✅ Signed in as {signed_in_email}")

if st.button("🚪 Sign out"):
    st.session_state.pop("token_result", None)
    clear_query_params_compat()
    st.rerun()


# ============================================================
# RESOLVE ROOT FOLDER
# ============================================================
st.sidebar.header("📁 SharePoint Source")
st.sidebar.caption("Select folder(s), then choose Excel files inside them.")

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

try:
    root_children = list_children_all(access_token, drive_id, root_item_id)
except Exception as e:
    st.error("Could not list folders from the root folder.")
    st.code(str(e))
    st.stop()

folders = [x for x in root_children if is_folder_item(x)]
folders.sort(key=lambda x: normalize_text(x.get("name", "")))

root_excel_files = [x for x in root_children if is_excel_name(x.get("name", ""))]

all_excel_files = []
folder_name_map = {}

if folders:
    folder_names = [f["name"] for f in folders]

    selected_folder_names = st.sidebar.multiselect(
        "Select folder(s)",
        folder_names,
        default=folder_names[:1],
    )

    selected_folders = [f for f in folders if f["name"] in selected_folder_names]

    for folder_info in selected_folders:
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
            item_copy["display_name"] = f"{folder_name} | {item_copy['name']}"
            all_excel_files.append(item_copy)
            folder_name_map[item_copy["id"]] = folder_name

for item in root_excel_files:
    item_copy = dict(item)
    item_copy["display_name"] = f"Root | {item_copy['name']}"
    all_excel_files.append(item_copy)
    folder_name_map[item_copy["id"]] = "Root"

all_excel_files.sort(key=lambda x: normalize_text(x["display_name"]))

if not all_excel_files:
    st.warning("No Excel files found inside the selected folder(s).")
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
# LOAD DATA
# ============================================================
raw_df = load_selected_excel_files_regular(access_token, drive_id, selected_excel_files, folder_name_map)

if raw_df.empty:
    st.error("No valid data could be loaded from the selected Excel files.")
    st.stop()


# ============================================================
# CLEANING + FILTERS
# ============================================================
df = raw_df.copy()

df["date"] = pd.to_datetime(df["date"], errors="coerce")
df["vendor_company"] = safe_text_series(df["vendor_company"])
df["employee"] = safe_text_series(df["employee"])
df["employee_class"] = safe_text_series(df["employee_class"]).replace({"": DEFAULT_CLASS_WHEN_NO_CLASS})
df["type_of_work"] = TYPE_OF_WORK_DEFAULT

for c in [
    "hours",
    "hourly_rate",
    "regular_hours",
    "suppl_hours",
    "conge_hours",
    "conge_trav_hours",
    "maladie_hours",
]:
    df[c] = to_num_series(df[c])

df = df.dropna(subset=["date"]).copy()

if df.empty:
    st.warning("No rows with valid dates were found. Check week number in column K.")
    st.stop()

st.sidebar.header("🧷 Report Filters")

all_vendors = sorted([v for v in df["vendor_company"].dropna().unique().tolist() if v])
selected_vendors = st.sidebar.multiselect("Vendor Company", all_vendors, default=all_vendors)

all_classes = sorted([v for v in df["employee_class"].dropna().unique().tolist() if v])
selected_classes = st.sidebar.multiselect("Employee Class", all_classes, default=all_classes)

all_employees = sorted([e for e in df["employee"].dropna().unique().tolist() if e])
selected_employees = st.sidebar.multiselect("Name Employee", all_employees, default=all_employees)

all_types = sorted([t for t in df["type_of_work"].dropna().unique().tolist() if t])
selected_types = st.sidebar.multiselect("Type of work", all_types, default=all_types)

start_date = st.sidebar.date_input("First committee week start date", value=FIRST_COMMITTEE_WEEK_START_DEFAULT.date())
num_weeks = st.sidebar.number_input("Number of weeks", min_value=1, max_value=24, value=4)

filtered_df = df.copy()

if selected_vendors:
    filtered_df = filtered_df[filtered_df["vendor_company"].isin(selected_vendors)]
if selected_classes:
    filtered_df = filtered_df[filtered_df["employee_class"].isin(selected_classes)]
if selected_employees:
    filtered_df = filtered_df[filtered_df["employee"].isin(selected_employees)]
if selected_types:
    filtered_df = filtered_df[filtered_df["type_of_work"].isin(selected_types)]

start_date_dt = pd.to_datetime(start_date)

filtered_df[["week_start", "week_end"]] = filtered_df["date"].apply(
    lambda x: pd.Series(assign_committee_week(x, start_date_dt, num_weeks))
)

filtered_df["week_start"] = pd.to_datetime(filtered_df["week_start"], errors="coerce")
filtered_df["week_end"] = pd.to_datetime(filtered_df["week_end"], errors="coerce")

filtered_df = filtered_df[filtered_df["week_start"].notna()].copy()
filtered_df["week_label"] = filtered_df["week_end"].dt.strftime("%Y-%m-%d")

if filtered_df.empty:
    st.warning("No data available for the selected filters.")
    st.stop()

weekly_summary = build_weekly_summary(filtered_df)


# ============================================================
# TOP SUMMARY
# ============================================================
total_gains_all = round(float(weekly_summary["total_pay"].sum()), 2)
total_reer_all = round(float(weekly_summary["reer"].sum()), 2)
total_with_reer_all = round(float(weekly_summary["total_with_reer"].sum()), 2)
levy_1pct = round(total_with_reer_all * 0.01, 2)
prelevement_total_du = round(total_with_reer_all + levy_1pct, 2)
total_hours_all = round(float(weekly_summary["total_hours"].sum()), 2)

col1, col2, col3, col4, col5 = st.columns(5)

col1.metric("TOTAL HOURS", f"{total_hours_all:,.2f}")
col2.metric("TOTAL DES GAINS", format_money(total_gains_all))
col3.metric("TOTAL REER", format_money(total_reer_all))
col4.metric("TOTAL GAINS + REER", format_money(total_with_reer_all))
col5.metric("PRÉLÈVEMENT TOTAL DÛ", format_money(prelevement_total_du))


# ============================================================
# EMPLOYEE SUMMARY
# ============================================================
st.subheader("Employee summary")

summary_view_cols = [
    "vendor_company",
    "employee",
    "employee_class",
    "week_label",
    "regular_hours",
    "overtime_hours",
    "suppl_hours",
    "conge_hours",
    "conge_trav_hours",
    "maladie_hours",
    "total_hours",
    "hourly_rate",
    "regular_pay",
    "overtime_pay",
    "suppl_pay",
    "conge_pay",
    "conge_trav_pay",
    "maladie_pay",
    "total_pay",
    "reer",
    "total_with_reer",
]

dataframe_with_2_decimals(weekly_summary[[c for c in summary_view_cols if c in weekly_summary.columns]])


# ============================================================
# SOURCE PREVIEW
# ============================================================
st.subheader("Filtered source data")

preview_cols = [
    "source_month_folder",
    "source_file",
    "excel_week_range",
    "excel_week_start",
    "excel_week_end",
    "special_lookup_date",
    "date",
    "day",
    "vendor_company",
    "employee",
    "employee_class",
    "type_of_work",
    "excel_cell_value",
    "regular_numeric_values",
    "visible_special_detected",
    "input_v_hours",
    "input_h_hours",
    "input_sd_hours",
    "special_hours_total",
    "hours",
    "regular_hours",
    "suppl_hours",
    "conge_hours",
    "conge_trav_hours",
    "maladie_hours",
    "hourly_rate",
    "week_label",
]

dataframe_with_2_decimals(filtered_df[[c for c in preview_cols if c in filtered_df.columns]])


# ============================================================
# EXPORT
# ============================================================
report_file = export_regular_hours_report(weekly_summary, start_date)

st.download_button(
    label="Download Regular Hours Excel Report",
    data=report_file,
    file_name="regular_hours_report.xlsx",
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
    st.write("Selected excel display names:", selected_excel_display_names)
    st.write("Loader diagnostics:", st.session_state.get("regular_loader_diagnostics", []))
    st.write("Final source columns:", list(filtered_df.columns))
    st.write("Weekly summary columns:", list(weekly_summary.columns))
    st.write("Special hours rule:", "V and H go to Congé. SD goes to Maladie. The value is taken once from Input/Imput, not multiplied by the number of letters.")
    st.write("Overtime condition:", "regular worked hours over 40 in the same committee week are overtime at 1.5x")
