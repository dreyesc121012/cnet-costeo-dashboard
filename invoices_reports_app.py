import io
import re
import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

st.set_page_config(page_title="CNET Invoices Reports", layout="wide")

# =========================
# Helpers
# =========================
def safe_num(x):
    try:
        if pd.isna(x):
            return 0.0
        s = str(x).replace("$", "").replace(",", "").strip()
        return float(s) if s else 0.0
    except:
        return 0.0

def contains_ci(text, needle):
    return needle.lower() in str(text or "").lower()

# =========================
# Login + Download
# =========================
def download_excel():
    session = requests.Session()

    # 1) GET login page to get CSRF
    login_page = session.get("https://app.master.cnetfranchise.com/login")
    soup = BeautifulSoup(login_page.text, "lxml")

    token = soup.select_one('input[name="_csrf_token"]')["value"]

    # 2) POST login
    payload = {
        "_csrf_token": token,
        "_username": st.secrets["SYSTEM_USERNAME"],
        "_password": st.secrets["SYSTEM_PASSWORD"],
        "_submit": "Login"
    }

    session.post("https://app.master.cnetfranchise.com/login_check", data=payload)

    # 3) Download Excel
    file_response = session.get(st.secrets["EXPORT_EXCEL_URL"])
    return file_response.content

# =========================
# UI
# =========================
st.title("CNET - Invoice Reports")

with st.sidebar:
    report = st.radio(
        "Select report",
        ["Resume Without Fees", "Validation", "Jeff-validation"]
    )
    run = st.button("Download + Process")

if not run:
    st.stop()

file = download_excel()
df = pd.read_excel(io.BytesIO(file))

st.success("Excel downloaded successfully")

st.dataframe(df.head())
