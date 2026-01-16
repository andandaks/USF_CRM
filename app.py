import streamlit as st
import pandas as pd
import random
import requests
import msal
import urllib
from datetime import datetime
from sqlalchemy import create_engine, text

# ==========================================
# --- CONFIGURATION ---
# ==========================================

# Проверяем, настроены ли секреты
if "database" not in st.secrets or "microsoft" not in st.secrets:
    st.error("❌ Secrets are not configured. Please add .streamlit/secrets.toml")
    st.stop()

# Database Config
DB_SERVER = st.secrets["database"]["server"]
DB_NAME = st.secrets["database"]["name"]
DB_USER = st.secrets["database"]["user"]
DB_PASSWORD = st.secrets["database"]["password"]
DB_DRIVER = "ODBC Driver 18 for SQL Server"

# Microsoft Graph Config
CLIENT_ID = st.secrets["microsoft"]["client_id"]
TENANT_ID = st.secrets["microsoft"]["tenant_id"]
ONEDRIVE_ROOT = '/Moco'

# ==========================================
# --- BACKEND FUNCTIONS ---
# ==========================================

@st.cache_resource
def get_db_engine():
    """Secure connection to Azure SQL"""
    params = urllib.parse.quote_plus(
        f"DRIVER={{{DB_DRIVER}}};SERVER={{tcp:{DB_SERVER},1433}};"
        f"DATABASE={{{DB_NAME}}};UID={{{DB_USER}}};PWD={{{DB_PASSWORD}}};"
        "Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
    )
    return create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

def init_db():
    """Create table if it doesn't exist"""
    engine = get_db_engine()
    query = """
    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='upshift_crm' and xtype='U')
    CREATE TABLE upshift_crm (
        deal_id BIGINT PRIMARY KEY,
        created_at NVARCHAR(50),
        client_name NVARCHAR(255),
        phone NVARCHAR(50),
        email NVARCHAR(100),
        website NVARCHAR(100),
        manager NVARCHAR(100),
        product_type NVARCHAR(50),
        sum_requested FLOAT,
        assigned_to NVARCHAR(50),
        comments NVARCHAR(MAX)
    )
    """
    try:
        with engine.begin() as conn:
            conn.execute(text(query))
    except Exception as e:
        st.error(f"DB Init Error: {e}")

# ==========================================
# --- CLOUD FUNCTIONS (OneDrive) ---
# ==========================================

def get_access_token():
    """Get Microsoft Graph Token"""
    if "onedrive_token" in st.session_state:
        return st.session_state["onedrive_token"]

    app = msal.PublicClientApplication(
        CLIENT_ID, 
        authority=f'https://login.microsoftonline.com/{TENANT_ID}'
    )
    
    # Try cache first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(['Files.ReadWrite.All', 'User.Read'], account=accounts[0])
        if result:
            st.session_state["onedrive_token"] = result['access_token']
            return result['access_token']
    return None

def trigger_login():
    """Interactive Login"""
    app = msal.PublicClientApplication(
        CLIENT_ID, 
        authority=f'https://login.microsoftonline.com/{TENANT_ID}'
    )
    result = app.acquire_token_interactive(scopes=['Files.ReadWrite.All', 'User.Read'])
    if "access_token" in result:
        st.session_state["onedrive_token"] = result['access_token']
        st.rerun()

def upload_to_onedrive(deal_id, files):
    """Create folder and upload files"""
    token = get_access_token()
    if not token or not files: return

    headers = {'Authorization': 'Bearer ' + token}
    
    # 1. Create Folder
    folder_url = f'https://graph.microsoft.com/v1.0/me/drive/root:{ONEDRIVE_ROOT}:/children'
    body = {
        "name": str(deal_id), 
        "folder": {}, 
        "@microsoft.graph.conflictBehavior": "rename"
    }
    requests.post(folder_url, headers={'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json'}, json=body)

    # 2. Upload Files
    for f in files:
        upload_url = f'https://graph.microsoft.com/v1.0/me/drive/root:{ONEDRIVE_ROOT}/{deal_id}/{f.name}:/content'
        requests.put(
            upload_url, 
            headers={'Authorization': 'Bearer ' + token, 'Content-Type': 'application/octet-stream'}, 
            data=f.getvalue()
        )

# ==========================================
# --- UI ---
# ==========================================

st.set_page_config(page_title="UpShift Finance CRM", layout="wide")

# Init DB on load
init_db()

# Header
c1, c2 = st.columns([8, 2])
with c1: st.title("🚀 UpShift Finance CRM")
with c2: 
    if not get_access_token():
        if st.button("🔑 Connect Cloud"): trigger_login()
    else:
        st.success("Cloud Connected")

st.markdown("---")

# --- FORM SECTION ---
with st.container():
    st.subheader("➕ New Deal Entry")
    with st.form("new_deal_form", clear_on_submit=True):
        # Row 1
        col1, col2, col3, col4 = st.columns(4)
        with col1: client = st.text_input("Client Name *")
        with col2: phone = st.text_input("Phone")
        with col3: email = st.text_input("Email")
        with col4: website = st.text_input("Website")
        
        # Row 2
        col5, col6, col7, col8 = st.columns(4)
        with col5: manager = st.text_input("Manager Name")
        with col6: product = st.selectbox("Product Type", ["Term Loan", "Credit Line"])
        with col7: req_sum = st.number_input("Sum Requested", min_value=0.0, step=1000.0)
        with col8: assigned = st.selectbox("Assigned To", ["Sales", "Underwriter", "Client"])
        
        # Row 3
        comments = st.text_area("Comments")
        uploaded_files = st.file_uploader("Attach Files", accept_multiple_files=True)

        submitted = st.form_submit_button("Submit Deal", type="primary")

        if submitted:
            if not client:
                st.error("⚠️ Client Name is required!")
            else:
                # 1. Prepare Data
                deal_id = random.randint(100000, 999999)
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                new_data = {
                    "deal_id": deal_id,
                    "created_at": timestamp,
                    "client_name": client,
                    "phone": phone,
                    "email": email,
                    "website": website,
                    "manager": manager,
                    "product_type": product,
                    "sum_requested": req_sum,
                    "assigned_to": assigned,
                    "comments": comments
                }

                # 2. Save to SQL
                try:
                    df = pd.DataFrame([new_data])
                    df.to_sql('upshift_crm', get_db_engine(), if_exists='append', index=False)
                    st.success(f"✅ Deal #{deal_id} saved to Database!")
                except Exception as e:
                    st.error(f"Database Error: {e}")

                # 3. Save to Cloud
                if get_access_token() and uploaded_files:
                    with st.spinner("Uploading files to OneDrive..."):
                        upload_to_onedrive(deal_id, uploaded_files)
                    st.toast(f"Files uploaded to /Moco/{deal_id}", icon="☁️")
                
                st.rerun()

st.markdown("---")

# --- TABLE SECTION ---
st.subheader("📂 Active Cases")

try:
    df_current = pd.read_sql("SELECT * FROM upshift_crm ORDER BY created_at DESC", get_db_engine())
    
    if df_current.empty:
        st.info("📭 The database is currently empty. Add a new case above.")
    else:
        # Display as a clean interactive table
        st.dataframe(
            df_current,
            column_config={
                "deal_id": st.column_config.NumberColumn("Deal ID", format="%d"),
                "sum_requested": st.column_config.NumberColumn("Sum", format="£%.2f"),
                "created_at": "Time Added",
                "website": st.column_config.LinkColumn("Website")
            },
            use_container_width=True,
            hide_index=True
        )
except Exception as e:
    st.error("Could not load data. Database might be initializing.")
