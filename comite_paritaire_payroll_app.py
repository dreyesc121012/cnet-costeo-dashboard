# ============================================================
# 🔥 IMPORTS
# ============================================================
import pandas as pd
import streamlit as st
from io import BytesIO
import requests
import base64

# ============================================================
# 🔥 CONFIG
# ============================================================
st.set_page_config(page_title="CNET Regular Hours Report", layout="wide")
st.title("CNET Regular Hours Report")

# ============================================================
# 🔐 AUTH (ya lo tienes funcionando)
# ============================================================
access_token = st.session_state.get("access_token", None)

if not access_token:
    st.warning("Not connected to OneDrive")
    st.stop()

# ============================================================
# 🔗 SHAREPOINT FOLDER
# ============================================================
FOLDER_URL = "https://groupcastillo.sharepoint.com/:f:/s/GroupCastilloTeamSite/IgDJ46w1V3YWT7e0yB8CKkD9AenZh0xzbn8pNRRGuDcIpPw?e=s4L0Z9"

def make_share_id(url):
    b = base64.b64encode(url.encode()).decode().rstrip("=")
    return "u!" + b.replace("/", "_").replace("+", "-")

def get_json(url):
    return requests.get(url, headers={"Authorization": f"Bearer {access_token}"}).json()

# ============================================================
# 📂 RESOLVE FOLDER
# ============================================================
meta = get_json(f"https://graph.microsoft.com/v1.0/shares/{make_share_id(FOLDER_URL)}/driveItem")

drive_id = meta["parentReference"]["driveId"]
folder_id = meta["id"]

files = get_json(f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}/children")["value"]

excel_files = [f for f in files if f["name"].endswith(".xlsx")]

if not excel_files:
    st.error("No Excel files found")
    st.stop()

file = excel_files[0]

# ============================================================
# 📥 DOWNLOAD FILE
# ============================================================
file_bytes = requests.get(
    f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file['id']}/content",
    headers={"Authorization": f"Bearer {access_token}"}
).content

# ============================================================
# 📊 READ EXCEL
# ============================================================
xls = pd.ExcelFile(BytesIO(file_bytes))

df = pd.read_excel(xls, sheet_name="DATA", header=None)
input_df = pd.read_excel(xls, sheet_name="Input", header=None)

# ============================================================
# 📅 WEEK PARSER (03/01 - 09/01)
# ============================================================
def parse_week(text):
    try:
        start = text.split("-")[0].strip()
        return pd.to_datetime(start + "/2026", format="%d/%m/%Y")
    except:
        return None

# ============================================================
# 🔥 BUILD DATASET
# ============================================================
rows = []

for i in range(4, len(df)):

    vendor = df.iloc[i, 0]
    employee = df.iloc[i, 1]
    emp_class = df.iloc[i, 8]
    week_text = df.iloc[i, 10]
    rate = df.iloc[i, 19]

    if pd.isna(employee):
        continue

    if str(emp_class).strip() == "No Class":
        emp_class = "Class A"

    week_start = parse_week(str(week_text))

    for j in range(11, 18):

        value = df.iloc[i, j]

        if pd.isna(value):
            continue

        work_date = week_start + pd.Timedelta(days=(j - 11))

        row = {
            "vendor": vendor,
            "employee": employee,
            "class": emp_class,
            "date": work_date,
            "rate": rate,
            "regular": 0,
            "vacation": 0,
            "sick": 0,
            "holiday": 0
        }

        # ====================================================
        # 🔥 LOGICA NUEVA (V / SD / H)
        # ====================================================
        if value == "V":
            hrs = input_df[input_df.iloc[:,12] == work_date].iloc[:,13].sum()
            row["vacation"] = hrs

        elif value == "SD":
            hrs = input_df[input_df.iloc[:,12] == work_date].iloc[:,14].sum()
            row["sick"] = hrs

        elif value == "H":
            hrs = input_df[input_df.iloc[:,12] == work_date].iloc[:,15].sum()
            row["holiday"] = hrs

        else:
            row["regular"] = float(value)

        rows.append(row)

# ============================================================
# 📊 DATAFRAME FINAL
# ============================================================
final = pd.DataFrame(rows)

# ============================================================
# 📊 SUMMARY
# ============================================================
summary = final.groupby(["vendor","employee","class"]).agg({
    "regular":"sum",
    "vacation":"sum",
    "sick":"sum",
    "holiday":"sum",
}).reset_index()

summary["total_hours"] = summary["regular"] + summary["vacation"] + summary["sick"] + summary["holiday"]

# ============================================================
# 📺 UI
# ============================================================
st.subheader("Employee Summary")
st.dataframe(summary)

# ============================================================
# 📤 EXPORT
# ============================================================
output = BytesIO()
summary.to_excel(output, index=False)

st.download_button(
    "Download Report",
    output.getvalue(),
    file_name="report.xlsx"
)
