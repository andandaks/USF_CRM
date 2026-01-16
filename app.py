import streamlit as st
import pandas as pd
import random
import requests
import msal
import io
import os
import urllib
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
import altair as alt
from sqlalchemy import create_engine, text

# ==========================================
# --- CONFIGURATION (FROM SECRETS) ---
# ==========================================

# Проверка наличия секретов перед запуском
if "database" not in st.secrets or "microsoft" not in st.secrets:
    st.error("❌ Файл .streamlit/secrets.toml не найден или не настроен.")
    st.stop()

# 1. AZURE SQL DATABASE DETAILS
DB_SERVER = st.secrets["database"]["server"]
DB_NAME = st.secrets["database"]["name"]
DB_USER = st.secrets["database"]["user"]
DB_PASSWORD = st.secrets["database"]["password"]
DB_DRIVER = st.secrets["database"]["driver"]

# 2. ONEDRIVE CONFIG
CLIENT_ID = st.secrets["microsoft"]["client_id"]
TENANT_ID = st.secrets["microsoft"]["tenant_id"]
ONEDRIVE_ROOT = '/Moco'
ONEDRIVE_PAYMENTS_PATH = f'{ONEDRIVE_ROOT}/Payments'

ENTITIES = ['sales', 'underwriter', 'client']

# ==========================================
# --- DATABASE ENGINE (SQL) ---
# ==========================================
@st.cache_resource
def get_db_engine():
    """Creates a secure connection pool to Azure SQL."""
    params = urllib.parse.quote_plus(
        f"DRIVER={{{DB_DRIVER}}};"
        f"SERVER={{tcp:{DB_SERVER},1433}};"
        f"DATABASE={{{DB_NAME}}};"
        f"UID={{{DB_USER}}};"
        f"PWD={{{DB_PASSWORD}}};"
        "Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
    )
    return create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

def init_db():
    """
    1. Creates table if not exists with STRICT TYPES.
    2. Auto-migrates (adds) new columns if they are missing.
    """
    engine = get_db_engine()
    
    # 1. Basic Create
    create_table_query = """
    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='crm_cases' and xtype='U')
    CREATE TABLE crm_cases (
        [unique case number in system] BIGINT PRIMARY KEY,
        [date added] NVARCHAR(50),
        [responsible entity] NVARCHAR(50),
        [company name] NVARCHAR(255),
        [company number] NVARCHAR(50),
        [manager] NVARCHAR(100),
        [product type] NVARCHAR(50),
        [phone] NVARCHAR(50),
        [email] NVARCHAR(100),
        [site] NVARCHAR(100),
        [sum] FLOAT,
        [has pledge] NVARCHAR(10),
        [returning client] NVARCHAR(10),
        [comment] NVARCHAR(MAX),
        [done] NVARCHAR(10),
        [kyc] NVARCHAR(20),
        [aml] NVARCHAR(20),
        [soft_check] NVARCHAR(20),
        [equifax_score] INT
    )
    """
    
    # 2. Migration Columns (Safe to run every time)
    alter_queries = [
        "IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'crm_cases') AND name = 'kyc') ALTER TABLE crm_cases ADD [kyc] NVARCHAR(20)",
        "IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'crm_cases') AND name = 'aml') ALTER TABLE crm_cases ADD [aml] NVARCHAR(20)",
        "IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'crm_cases') AND name = 'soft_check') ALTER TABLE crm_cases ADD [soft_check] NVARCHAR(20)",
        "IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'crm_cases') AND name = 'equifax_score') ALTER TABLE crm_cases ADD [equifax_score] INT"
    ]

    try:
        with engine.connect() as conn:
            conn.execute(text(create_table_query))
            for q in alter_queries:
                conn.execute(text(q))
            conn.commit()
    except Exception as e:
        st.error(f"Database Connection Error: {e}. Check password/firewall.")

# ==========================================
# --- ONEDRIVE AUTH (STABLE) ---
# ==========================================
def get_access_token():
    if "onedrive_token" in st.session_state:
        return st.session_state["onedrive_token"]

    app = msal.PublicClientApplication(CLIENT_ID, authority=f'https://login.microsoftonline.com/{TENANT_ID}')
    scopes = ['Files.ReadWrite.All', 'User.Read']
    
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result and "access_token" in result:
            st.session_state["onedrive_token"] = result['access_token']
            return result['access_token']
    return None

def trigger_login():
    app = msal.PublicClientApplication(CLIENT_ID, authority=f'https://login.microsoftonline.com/{TENANT_ID}')
    result = app.acquire_token_interactive(scopes=['Files.ReadWrite.All', 'User.Read'])
    if "access_token" in result:
        st.session_state["onedrive_token"] = result['access_token']
        st.rerun()

# ==========================================
# --- CLOUD HELPERS ---
# ==========================================
def create_deal_folder(deal_id):
    token = get_access_token()
    if not token: return
    headers = {'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json'}
    url = f'https://graph.microsoft.com/v1.0/me/drive/root:{ONEDRIVE_ROOT}:/children'
    body = {"name": str(deal_id), "folder": {}, "@microsoft.graph.conflictBehavior": "rename"}
    try: requests.post(url, headers=headers, json=body)
    except: pass

def upload_file_to_folder(deal_id, uploaded_file):
    token = get_access_token()
    if not token: return False
    path = f"{ONEDRIVE_ROOT}/{deal_id}/{uploaded_file.name}"
    url = f'https://graph.microsoft.com/v1.0/me/drive/root:{path}:/content'
    headers = {'Authorization': 'Bearer ' + token, 'Content-Type': 'application/octet-stream'}
    try:
        r = requests.put(url, headers=headers, data=uploaded_file.getvalue())
        return r.status_code in [200, 201]
    except: return False

def list_files_in_deal_folder(deal_id):
    token = get_access_token()
    if not token: return []
    headers = {'Authorization': 'Bearer ' + token}
    url = f'https://graph.microsoft.com/v1.0/me/drive/root:{ONEDRIVE_ROOT}/{deal_id}:/children'
    try:
        r = requests.get(url, headers=headers)
        if r.status_code == 200:
            return [i['name'] for i in r.json().get('value', []) if 'file' in i]
    except: pass
    return []

@st.cache_data(ttl=300)
def load_all_payments_from_cloud():
    token = get_access_token()
    if not token: return pd.DataFrame()
    headers = {'Authorization': 'Bearer ' + token}
    list_url = f'https://graph.microsoft.com/v1.0/me/drive/root:{ONEDRIVE_PAYMENTS_PATH}:/children'
    master_list = []
    try:
        r = requests.get(list_url, headers=headers)
        if r.status_code == 200:
            for f in r.json().get('value', []):
                if f.get('name', '').endswith('.xlsx'):
                    d_url = f.get('@microsoft.graph.downloadUrl')
                    if d_url:
                        try:
                            content = requests.get(d_url).content
                            df = pd.read_excel(io.BytesIO(content))
                            df.columns = df.columns.str.strip().str.title()
                            df['Case ID'] = f['name'].replace('.xlsx', '')
                            master_list.append(df)
                        except: pass
    except: pass
    return pd.concat(master_list, ignore_index=True) if master_list else pd.DataFrame()

# ==========================================
# --- APP UI ---
# ==========================================
st.set_page_config(page_title="UpShift Finance", layout="wide")

with st.sidebar:
    st.title("Authentication")
    token = get_access_token()
    if token:
        st.success("✅ Connected to OneDrive")
    else:
        st.warning("⚠️ Disconnected")
        if st.button("🔌 Connect Microsoft"):
            trigger_login()

# --- SQL DB HELPERS ---
def load_data_sql():
    """Loads all cases from SQL DB."""
    try:
        return pd.read_sql("SELECT * FROM crm_cases", get_db_engine())
    except:
        return pd.DataFrame()

def save_new_case_sql(data):
    """Appends a single new case to SQL DB."""
    df = pd.DataFrame([data])
    df.to_sql('crm_cases', get_db_engine(), if_exists='append', index=False)

def update_table_sql(df):
    """
    Updates the entire table safely.
    DELETE + APPEND logic to preserve schema.
    """
    engine = get_db_engine()
    if df.empty:
        return 
    
    with engine.begin() as conn:
        conn.execute(text("DELETE FROM crm_cases"))
        
    df.to_sql('crm_cases', engine, if_exists='append', index=False)

def page_crm():
    st.title("UpShift CRM 🗄️ (SQL + Cloud)")
    
    # 1. Initialize & Load
    init_db()
    
    # Всегда подгружаем свежие данные при старте
    if 'df' not in st.session_state:
        st.session_state.df = load_data_sql()

    # Styling
    def color_prod(val):
        c = {"Term Loan": "#D6EAF8", "Credit Line": "#A9CCE3", "Invoice Factoring": "#E8F8F5", "Other": "#F2F3F4"}
        return f'background-color: {c.get(val, "")}; color: black' if val in c else ''
    def highlight_null(row):
        styles = [''] * len(row)
        if row.isna().any():
            try: styles[row.index.get_loc(row.isna().idxmax())] = 'background-color: #fff9c4; font-weight: bold'
            except: pass
        return styles

    # 2. Add Case Form
    with st.expander("➕ Add New Case", expanded=True):
        with st.form("add_case"):
            st.write("Enter details for new inquiry:")
            
            # Row 1
            c1, c2, c3, c4 = st.columns(4)
            with c1: client = st.text_input("Client Name *")
            with c2: co_num = st.text_input("Company Number")
            with c3: mgr = st.text_input("Manager")
            with c4: assigned = st.selectbox("Assigned To", ENTITIES)
            
            # Row 2
            c5, c6, c7, c8 = st.columns(4)
            with c5: phone = st.text_input("Phone")
            with c6: email = st.text_input("Email")
            with c7: site = st.text_input("Website")
            with c8: prod = st.selectbox("Product", ["Term Loan", "Credit Line", "Invoice Factoring", "Other"])

            # Row 3 (Compliance)
            st.markdown("**Compliance & Scores**")
            c9, c10, c11, c12 = st.columns(4)
            with c9: kyc = st.selectbox("KYC Status", ["N/A", "Passed", "No"])
            with c10: aml = st.selectbox("AML Status", ["N/A", "Passed", "No"])
            with c11: soft = st.selectbox("Soft Check", ["N/A", "Passed", "No"])
            with c12: equifax = st.number_input("Equifax Score", 0, 999, 0)
            
            # Row 4
            c13, c14, c15 = st.columns(3)
            with c13: loan = st.number_input("Sum (£)", step=1000.0)
            with c14: pledge = st.toggle("Pledge?")
            with c15: ret = st.toggle("Returning?")
            
            comm = st.text_area("Comments")
            files = st.file_uploader("📂 Initial Files", accept_multiple_files=True)

            if st.form_submit_button("Create Case", type="primary"):
                if not client:
                    st.error("Client Name Required")
                else:
                    nid = random.randint(100000, 999999)
                    data = {
                        "unique case number in system": nid,
                        "date added": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "responsible entity": assigned, 
                        "company name": client, "company number": co_num, "manager": mgr,
                        "product type": prod, "phone": phone, "email": email, "site": site,
                        "sum": loan, "has pledge": "Yes" if pledge else "No",
                        "returning client": "Yes" if ret else "No", "comment": comm, "done": "No",
                        "kyc": kyc, "aml": aml, "soft_check": soft, "equifax_score": equifax
                    }
                    
                    # SQL Save
                    save_new_case_sql(data)
                    
                    # Update Local State
                    st.session_state.df = load_data_sql()
                    
                    # Cloud Operations
                    if get_access_token():
                        create_deal_folder(nid)
                        if files:
                            for f in files: upload_file_to_folder(nid, f)
                            st.toast("Files Uploaded!", icon="📂")
                    
                    st.success("Case Created & Saved to DB!")
                    st.rerun()

    st.markdown("---")
    
    # 3. Main Table
    col_l, col_r = st.columns([2, 8])
    with col_l: view_mode = st.toggle("Show Formatting", value=True)
    with col_r: 
        if st.button("💾 Save Edits to SQL", type="primary"):
            update_table_sql(st.session_state.df)
            st.success("Database Updated Successfully!")
            st.session_state.df = load_data_sql() 

    conf = {
        "responsible entity": st.column_config.SelectboxColumn("Assigned To", options=ENTITIES, required=True),
        "unique case number in system": st.column_config.NumberColumn("ID", format="%d", disabled=True),
        "sum": st.column_config.NumberColumn("Sum", format="£%.2f"),
        "done": st.column_config.SelectboxColumn("Done?", options=["Yes", "No"], required=True),
        "product type": st.column_config.SelectboxColumn("Product", options=["Term Loan", "Credit Line", "Invoice Factoring", "Other"], required=True),
        "comment": st.column_config.TextColumn("Comment", width="large"),
        "kyc": st.column_config.SelectboxColumn("KYC", options=["N/A", "Passed", "No"]),
        "aml": st.column_config.SelectboxColumn("AML", options=["N/A", "Passed", "No"]),
        "soft_check": st.column_config.SelectboxColumn("Soft Check", options=["N/A", "Passed", "No"]),
        "equifax_score": st.column_config.NumberColumn("Equifax", format="%d")
    }

    if 'done' in st.session_state.df.columns:
        mask = st.session_state.df['done'] != "Yes"
    else:
        mask = [True] * len(st.session_state.df)

    for ent in ENTITIES:
        st.subheader(f"{ent.capitalize()} Queue")
        
        if 'responsible entity' in st.session_state.df.columns:
            curr = st.session_state.df.loc[(st.session_state.df['responsible entity'] == ent) & mask]
            
            if view_mode:
                st.dataframe(curr.style.apply(highlight_null, axis=1).map(color_prod, subset=['product type']), use_container_width=True, hide_index=True)
            else:
                edited = st.data_editor(curr, key=f"ed_{ent}", column_config=conf, use_container_width=True, hide_index=True, num_rows="dynamic")
                if not edited.equals(curr):
                    st.session_state.df.update(edited)
                    st.warning("Unsaved changes. Click 'Save Edits' above.")

    # 4. File Manager
    st.markdown("---")
    st.header("📂 File Manager")
    
    if get_access_token():
        ids = sorted(st.session_state.df['unique case number in system'].unique(), reverse=True) if not st.session_state.df.empty else []
        c_sel, c_up = st.columns([1, 2])
        
        with c_sel:
            sid = st.selectbox("Select Case", ids)
            if sid:
                st.caption(f"Files in Moco/{sid}:")
                for f in list_files_in_deal_folder(sid): st.text(f"📄 {f}")
        
        with c_up:
            if sid:
                u_files = st.file_uploader("Add Files", accept_multiple_files=True, key=f"up_{sid}")
                if u_files and st.button("Upload"):
                    for f in u_files: upload_file_to_folder(sid, f)
                    st.success("Uploaded!")
                    st.rerun()
    else:
        st.info("Connect to OneDrive in the sidebar to manage files.")

def page_archive():
    st.title("✅ Archive")
    if st.button("Refresh Archive"):
        st.session_state.df = load_data_sql()
        
    if 'df' not in st.session_state: st.session_state.df = load_data_sql()
    if st.session_state.df.empty: return
    
    done = st.session_state.df[st.session_state.df['done'] == "Yes"]
    c1, c2 = st.columns(2)
    c1.metric("Total", len(done))
    c2.metric("Sum", f"£{done['sum'].sum():,.2f}")
    st.dataframe(done, use_container_width=True, hide_index=True)

def page_loans():
    st.title("Loan Management (Cloud)")
    if not get_access_token():
        st.warning("Please connect to OneDrive in the sidebar.")
        return

    with st.spinner("Loading Payments..."):
        df = load_all_payments_from_cloud()
    
    if df.empty:
        st.info("No payments found in Moco/Payments.")
        return

    df['Date'] = pd.to_datetime(df['Date'])
    c1, c2, c3 = st.columns(3)
    c1.metric("Loaned", f"£{df[df['Sum']<0]['Sum'].sum():,.0f}")
    c2.metric("Repaid", f"£{df[df['Sum']>0]['Sum'].sum():,.0f}")
    c3.metric("Net", f"£{df['Sum'].sum():,.0f}")
    
    st.altair_chart(alt.Chart(df).mark_bar().encode(
        x='Date:T', y='Sum:Q', color='Case ID:N', tooltip=['Date', 'Case ID', 'Sum']
    ).interactive(), use_container_width=True)

def page_calculator():
    st.title("🧮 Calculator")
    c1, c2, c3, c4 = st.columns(4)
    amt = c1.number_input("Amount", 100000)
    rate = c2.number_input("Rate %", 12.0)
    mths = c3.number_input("Months", 12)
    start = c4.date_input("Start", date.today())
    method = st.radio("Method", ["Annuity", "Differentiated", "Interest Only"], horizontal=True)

    if st.button("Generate"):
        data = []
        bal = amt
        r_mo = rate/100/12
        curr = start
        
        if method == "Annuity":
            pmt = amt * r_mo * ((1+r_mo)**mths)/(((1+r_mo)**mths)-1) if r_mo>0 else amt/mths
            for i in range(1, int(mths)+1):
                inte = bal * r_mo
                princ = pmt - inte
                bal -= princ
                curr += relativedelta(months=1)
                data.append({"Period":i, "Date":curr, "Payment":pmt, "Principal":princ, "Interest":inte, "Balance":max(0, bal)})
        
        elif method == "Differentiated":
            princ_fix = amt/mths
            for i in range(1, int(mths)+1):
                inte = bal * r_mo
                pmt = princ_fix + inte
                bal -= princ_fix
                curr += relativedelta(months=1)
                data.append({"Period":i, "Date":curr, "Payment":pmt, "Principal":princ_fix, "Interest":inte, "Balance":max(0, bal)})
        
        else:
            for i in range(1, int(mths)+1):
                inte = bal * r_mo
                princ = bal if i==mths else 0
                pmt = princ + inte
                bal -= princ
                curr += relativedelta(months=1)
                data.append({"Period":i, "Date":curr, "Payment":pmt, "Principal":princ, "Interest":inte, "Balance":max(0, bal)})

        df = pd.DataFrame(data)
        st.dataframe(df, use_container_width=True)
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button("Download CSV", csv, "schedule.csv", "text/csv")

st.sidebar.title("Navigation")
pg = st.sidebar.radio("Go to:", ["CRM Dashboard", "Archive", "Loan Management", "Loan Calculator"])

if pg == "CRM Dashboard": page_crm()
elif pg == "Archive": page_archive()
elif pg == "Loan Management": page_loans()
elif pg == "Loan Calculator": page_calculator()
