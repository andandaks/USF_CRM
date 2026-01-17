import streamlit as st
import pandas as pd
import urllib
import sys
from sqlalchemy import create_engine, text
import msal
import requests
import base64
from datetime import datetime
from io import BytesIO

# --- CONFIGURATION ---
try:
    # Database Config
    DB_SERVER = st.secrets["database"]["server"]
    DB_NAME = st.secrets["database"]["name"]
    DB_USER = st.secrets["database"]["user"]
    DB_PASSWORD = st.secrets["database"]["password"]
    
    # Azure / OneDrive Config
    CLIENT_ID = st.secrets["azure"]["client_id"]
    CLIENT_SECRET = st.secrets["azure"]["client_secret"]
    TENANT_ID = st.secrets["azure"]["tenant_id"]
    AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
    SHARE_LINK = st.secrets["azure"]["sharepoint_link"]
except Exception:
    st.error("Missing Secrets! Please configure .streamlit/secrets.toml")
    st.stop()

# --- DATABASE CONNECTION (SQLAlchemy) ---
def get_db_engine():
    # 1. Detect Environment & Select Driver
    if sys.platform == "linux":
        # Streamlit Cloud (Linux) -> Use FreeTDS
        driver = 'FreeTDS'
        # FreeTDS requires specific syntax in the connection string for Azure
        params = urllib.parse.quote_plus(
            f"DRIVER={{{driver}}};SERVER={DB_SERVER};PORT=1433;DATABASE={DB_NAME};"
            f"UID={DB_USER};PWD={DB_PASSWORD};TDS_Version=7.4;"
        )
    else:
        # Local (Windows) -> Use ODBC Driver 18 (as in your snippet)
        driver = 'ODBC Driver 18 for SQL Server'
        params = urllib.parse.quote_plus(
            f"DRIVER={{{driver}}};SERVER={DB_SERVER};DATABASE={DB_NAME};"
            f"UID={DB_USER};PWD={DB_PASSWORD};Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
        )

    # 2. Create SQLAlchemy Engine
    # We use the standard pyodbc connection string wrapped in SQLAlchemy
    connection_url = f"mssql+pyodbc:///?odbc_connect={params}"
    engine = create_engine(connection_url)
    return engine

def load_data():
    engine = get_db_engine()
    # Using your new table name: 'crm_cases'
    try:
        query = "SELECT * FROM crm_cases"
        df = pd.read_sql(query, engine)
        return df
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame() # Return empty if error

def save_new_case(data):
    engine = get_db_engine()
    
    # Construct SQL Insert
    columns = ", ".join(data.keys())
    placeholders = ", ".join([f":{key}" for key in data.keys()])
    sql = text(f"INSERT INTO crm_cases ({columns}) VALUES ({placeholders})")
    
    try:
        with engine.connect() as conn:
            conn.execute(sql, data)
            conn.commit()
        return True
    except Exception as e:
        st.error(f"Database Error: {e}")
        return False

# --- ONEDRIVE FUNCTIONS ---
# (Kept identical to previous working version)
def get_graph_access_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return result.get("access_token")

def get_drive_item_from_link(token, url):
    base64_value = base64.b64encode(url.encode("utf-8")).decode("utf-8")
    encoded_url = "u!" + base64_value.replace("/", "_").replace("+", "-").rstrip("=")
    api_url = f"https://graph.microsoft.com/v1.0/shares/{encoded_url}/driveItem"
    headers = {'Authorization': 'Bearer ' + token}
    response = requests.get(api_url, headers=headers)
    if response.status_code == 200:
        data = response.json()
        return data["parentReference"]["driveId"], data["id"]
    return None, None

def create_folder_and_upload(file_obj, case_id):
    token = get_graph_access_token()
    if not token: return
    drive_id, master_folder_id = get_drive_item_from_link(token, SHARE_LINK)
    if not drive_id: return

    headers = {'Authorization': 'Bearer ' + token}
    create_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{master_folder_id}/children"
    body = {"name": str(case_id), "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    
    res = requests.post(create_url, headers=headers, json=body)
    if res.status_code == 201:
        folder_id = res.json()["id"]
    elif res.status_code == 409:
        # If folder exists, we must find its ID manually (omitted for brevity, assume new cases)
        st.warning("Folder already exists. Uploading to root of shared folder for safety.")
        folder_id = master_folder_id 
    else:
        return

    upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}:/{file_obj.name}:/content"
    file_obj.seek(0)
    requests.put(upload_url, headers=headers, data=file_obj)
    st.success(f"Uploaded {file_obj.name}")

# --- UI ---
st.set_page_config(page_title="Case Manager", layout="wide")
st.title("🏦 UpShift CRM")

page = st.sidebar.radio("Navigate", ["Dashboard", "Add New Case"])

if page == "Dashboard":
    st.header("Active Cases")
    
    # LOAD DATA
    df = load_data()
    
    if not df.empty:
        # DOWNLOAD BUTTON (Feature from your script)
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Cases')
        
        st.download_button(
            label="📥 Download All Data to Excel",
            data=buffer.getvalue(),
            file_name="crm_cases.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # TABLES
        st.subheader("🔹 Sales")
        if 'responsible_entity' in df.columns:
            st.dataframe(df[df['responsible_entity'] == 'Sales'], use_container_width=True)
            
            st.subheader("🔸 Underwriter")
            st.dataframe(df[df['responsible_entity'] == 'Underwriter'], use_container_width=True)
            
            st.subheader("▫️ Client")
            st.dataframe(df[df['responsible_entity'] == 'Client'], use_container_width=True)
        else:
            st.warning("Column 'responsible_entity' not found in table 'crm_cases'. Showing full table:")
            st.dataframe(df)

elif page == "Add New Case":
    st.header("Add New Entity")
    with st.form("new_case"):
        # Adjust these inputs to match your actual 'crm_cases' columns
        col1, col2 = st.columns(2)
        with col1:
            unique_id = st.number_input("Case ID", min_value=1)
            responsible = st.selectbox("Responsible", ["Sales", "Underwriter", "Client"])
            company = st.text_input("Company Name")
        with col2:
            email = st.text_input("Email")
            phone = st.text_input("Phone")
            amount = st.number_input("Sum", 0.0)
        
        uploaded_files = st.file_uploader("Docs", accept_multiple_files=True)
        submitted = st.form_submit_button("Save Case")
        
        if submitted:
            # MAP FORM TO DATABASE COLUMNS
            # Ensure these keys match your SQL table columns exactly!
            new_data = {
                "unique_case_number": unique_id,
                "responsible_entity": responsible,
                "company_name": company,
                "email": email,
                "phone": phone,
                "sum_value": amount,
                "date_added": datetime.now()
            }
            
            if save_new_case(new_data):
                st.success("Saved to DB!")
                if uploaded_files:
                    for f in uploaded_files:
                        create_folder_and_upload(f, unique_id)
