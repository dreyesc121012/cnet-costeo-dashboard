import base64
from io import BytesIO
import pandas as pd
import streamlit as st
import requests
import msal
import plotly.express as px

# CONFIG

CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
TENANT_ID = st.secrets["TENANT_ID"]
REDIRECT_URI = st.secrets["REDIRECT_URI"]

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read","Files.Read.All"]

st.set_page_config(page_title="Invoices Control",layout="wide")
st.title("📑 Invoice Category Control Dashboard")

# MSAL LOGIN

def get_msal_app():
return msal.ConfidentialClientApplication(
CLIENT_ID,
authority=AUTHORITY,
client_credential=CLIENT_SECRET,
)

# Convert Share URL

def make_share_id(shared_url):
b = base64.b64encode(shared_url.encode()).decode()
b = b.rstrip("=").replace("/","_").replace("+","-")
return "u!"+b

def download_excel(access_token,url):

```
share_id = make_share_id(url)  

meta = requests.get(  
    f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem",  
    headers={"Authorization":f"Bearer {access_token}"}  
)  

item = meta.json()  

drive = item["parentReference"]["driveId"]  
item_id = item["id"]  

file = requests.get(  
    f"https://graph.microsoft.com/v1.0/drives/{drive}/items/{item_id}/content",  
    headers={"Authorization":f"Bearer {access_token}"}  
)  

return file.content  
```

# LOGIN

app = get_msal_app()

if "token" not in st.session_state:

```
auth_url = app.get_authorization_request_url(  
    SCOPES,  
    redirect_uri=REDIRECT_URI  
)  

st.link_button("Login to OneDrive",auth_url)  
st.stop()  
```

access_token = st.session_state["token"]

# SELECT FILE

st.header("Excel Source")

excel_url = st.text_input("Paste OneDrive Excel Share Link")

if st.button("Load Excel"):

```
bytes_file = download_excel(access_token,excel_url)  

xls = pd.ExcelFile(BytesIO(bytes_file))  

payments = pd.read_excel(xls,"2025 Summary PAYMENTS")  
invoicing = pd.read_excel(xls,"Invoicing")  

payments = payments.rename(columns={  
    "Building Address":"Address",  
    "Category":"Category",  
    "Amount without taxes":"Actual"  
})  

actual = payments.groupby(["Address","Category"])["Actual"].sum().reset_index()  

budget = invoicing.rename(columns={"Building Address":"Address"})  

merged = actual.merge(budget,on="Address",how="left")  

st.subheader("Actual vs Budget")  

st.dataframe(merged,use_container_width=True)  

fig = px.bar(  
    merged
```

