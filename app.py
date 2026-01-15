import streamlit as st
import pandas as pd
import random
import requests
import msal
import io
import os
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
import altair as alt

# ==========================================
# --- CONFIGURATION (SECURE) ---
# ==========================================
try:
    CLIENT_ID = st.secrets["azure"]["client_id"]
    TENANT_ID = st.secrets["azure"]["tenant_id"]
    AUTHORITY_URL = f"https://login.microsoftonline.com/{TENANT_ID}"
    
    ONEDRIVE_ROOT = st.secrets["onedrive"]["root_folder"]
    ONEDRIVE_DB_PATH = f'{ONEDRIVE_ROOT}/case_book_db.xlsx' 
    ONEDRIVE_PAYMENTS_PATH = f'{ONEDRIVE_ROOT}/Payments'
except Exception as e:
    st.error(f"Missing secrets: {e}. Please add them in Streamlit Cloud settings.")
    st.stop()

ENTITIES = ['sales', 'underwriter', 'client']
SCOPES = ['Files.ReadWrite.All', 'User.Read']

# ==========================================
# --- AUTHENTICATION LOGIC (DEVICE CODE FLOW) ---
# ==========================================
def get_access_token():
    if "onedrive_token" in st.session_state:
        return st.session_state["onedrive_token"]
    return None

def login_with_device_code():
    """
    Initiates Device Code Flow for headless environments (Streamlit Cloud).
    """
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY_URL)
    
    # 1. Initiate flow
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        st.error("Failed to create device flow. Check Azure App registration.")
        return
    
    # 2. Show Code to User
    st.warning(f"⚠️ **Action Required**")
    st.markdown(f"""
    1. Go to: [{flow['verification_uri']}]({flow['verification_uri']})  
    2. Enter code: **{flow['user_code']}** """)
    
    # 3. Wait for user to complete login (Blocking)
    with st.spinner("Waiting for you to sign in..."):
        result = app.acquire_token_by_device_flow(flow)
    
    # 4. Handle Result
    if "access_token" in result:
        st.session_state["onedrive_token"] = result['access_token']
        st.success("Login Successful!")
        st.rerun()
    else:
        st.error(f"Login failed: {result.get('error_description')}")

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

# --- EXCEL OPERATIONS ---
def load_excel_from_onedrive(path):
    token = get_access_token()
    if not token: return pd.DataFrame()
    headers = {'Authorization': 'Bearer ' + token}
    url = f'https://graph.microsoft.com/v1.0/me/drive/root:{path}:/content'
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return pd.read_excel(io.BytesIO(response.content), dtype=object)
    return pd.DataFrame()

def save_excel_to_onedrive(df, path):
    token = get_access_token()
    if not token: return
    headers = {'Authorization': 'Bearer ' + token, 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    data = output.getvalue()
    
    url = f'https://graph.microsoft.com/v1.0/me/drive/root:{path}:/content'
    res = requests.put(url, headers=headers, data=data)
    if res.status_code in [200, 201]:
        st.toast("Saved to Cloud!", icon="☁️")
    else:
        st.error(f"Save failed: {res.text}")

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
        st.success("✅ Connected")
    else:
        st.warning("⚠️ Disconnected")
        if st.button("🔌 Connect Microsoft"):
            login_with_device_code()

def page_crm():
    st.title("UpShift CRM ☁️ (Excel on OneDrive)")

    # 1. LOAD DATA
    if 'df' not in st.session_state:
        if get_access_token():
            with st.spinner("Downloading Database..."):
                st.session_state.df = load_excel_from_onedrive(ONEDRIVE_DB_PATH)
                
                # --- NORMALIZE COLUMNS ---
                if not st.session_state.df.empty:
                    st.session_state.df.columns = st.session_state.df.columns.str.strip().str.lower()
                
                required_cols = [
                    "unique case number in system", "date added", "responsible entity", 
                    "company name", "company number", "manager", "product type", 
                    "phone", "email", "site", "sum", "has pledge", 
                    "returning client", "comment", "done"
                ]
                
                if st.session_state.df.empty:
                    st.session_state.df = pd.DataFrame(columns=required_cols)
                
                for c in required_cols:
                    if c not in st.session_state.df.columns:
                        st.session_state.df[c] = None
                
                if 'done' in st.session_state.df.columns:
                     st.session_state.df['done'] = st.session_state.df['done'].astype(str).str.title()
                if 'sum' in st.session_state.df.columns:
                     st.session_state.df['sum'] = pd.to_numeric(st.session_state.df['sum'], errors='coerce').fillna(0)
        else:
            st.info("Please connect to OneDrive in the sidebar.")
            st.session_state.df = pd.DataFrame()

    def save_data():
        with st.spinner("Syncing..."):
            save_excel_to_onedrive(st.session_state.df, ONEDRIVE_DB_PATH)

    # Styling
    def color_prod(val):
        c = {"term loan": "#D6EAF8", "credit line": "#A9CCE3", "invoice factoring": "#E8F8F5", "other": "#F2F3F4", 
             "Term Loan": "#D6EAF8", "Credit Line": "#A9CCE3", "Invoice Factoring": "#E8F8F5", "Other": "#F2F3F4"}
        return f'background-color: {c.get(val, "")}; color: black' if val in c else ''
        
    def highlight_null(row):
        styles = [''] * len(row)
        if row.isna().any():
            try: styles[row.index.get_loc(row.isna().idxmax())] = 'background-color: #fff9c4; font-weight: bold'
            except: pass
        return styles

    # 2. Add Case
    with st.expander("➕ Add New Case", expanded=True):
        with st.form("add_case"):
            c1, c2, c3 = st.columns(3)
            with c1:
                client = st.text_input("Client Name *")
                co_num = st.text_input("Company Number")
                mgr = st.text_input("Manager")
                prod = st.selectbox("Product", ["Term Loan", "Credit Line", "Invoice Factoring", "Other"])
            with c2:
                phone = st.text_input("Phone")
                email = st.text_input("Email")
                site = st.text_input("Website")
                loan = st.number_input("Sum (£)", step=1000.0)
            with c3:
                pledge = st.toggle("Pledge?")
                ret = st.toggle("Returning?")
                comm = st.text_area("Comments")
            
            files = st.file_uploader("📂 Initial Files", accept_multiple_files=True)

            if st.form_submit_button("Create Case", type="primary"):
                if not client:
                    st.error("Client Name Required")
                elif not get_access_token():
                    st.error("Connect to OneDrive first!")
                else:
                    nid = random.randint(100000, 999999)
                    data = {
                        "unique case number in system": nid,
                        "date added": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "responsible entity": "sales",
                        "company name": client, "company number": co_num, "manager": mgr,
                        "product type": prod, "phone": phone, "email": email, "site": site,
                        "sum": loan, "has pledge": "Yes" if pledge else "No",
                        "returning client": "Yes" if ret else "No", "comment": comm, "done": "No"
                    }
                    new_row = pd.DataFrame([data]).astype(object)
                    st.session_state.df = pd.concat([st.session_state.df, new_row], ignore_index=True)
                    
                    save_data()
                    create_deal_folder(nid)
                    if files:
                        for f in files: upload_file_to_folder(nid, f)
                        st.toast("Files Uploaded!")
                    st.rerun()

    st.markdown("---")
    
    # 3. Table
    col_l, col_r = st.columns([2, 8])
    with col_l: view_mode = st.toggle("Show Formatting", value=True)
    with col_r: 
        if st.button("💾 Force Sync"): save_data()

    conf = {
        "responsible entity": st.column_config.SelectboxColumn("Entity", options=ENTITIES, required=True),
        "unique case number in system": st.column_config.NumberColumn("ID", format="%d", disabled=True),
        "sum": st.column_config.NumberColumn("Sum", format="£%.2f"),
        "done": st.column_config.SelectboxColumn("Done?", options=["Yes", "No"], required=True),
        "product type": st.column_config.SelectboxColumn("Product", options=["Term Loan", "Credit Line", "Invoice Factoring", "Other"], required=True),
        "date added": st.column_config.TextColumn("Date", disabled=True),
        "comment": st.column_config.TextColumn("Comment", width="large")
    }

    if not st.session_state.df.empty and 'done' in st.session_state.df.columns:
        mask = st.session_state.df['done'] != "Yes"
    else:
        mask = [True] * len(st.session_state.df)

    for ent in ENTITIES:
        st.subheader(f"{ent.capitalize()} Queue")
        
        if 'responsible entity' in st.session_state.df.columns:
            curr = st.session_state.df.loc[(st.session_state.df['responsible entity'] == ent) & mask]
        else:
            curr = pd.DataFrame()

        if view_mode:
            if 'product type' in curr.columns:
                st.dataframe(curr.style.apply(highlight_null, axis=1).map(color_prod, subset=['product type']), width=None, hide_index=True)
            else:
                st.dataframe(curr.style.apply(highlight_null, axis=1), width=None, hide_index=True)
        else:
            edited = st.data_editor(curr, key=f"ed_{ent}", column_config=conf, width=None, hide_index=True, num_rows="dynamic")
            if not edited.equals(curr):
                st.session_state.df.update(edited)
                st.rerun()

    # 4. File Manager
    st.markdown("---")
    st.header("📂 File Manager")
    
    if get_access_token() and not st.session_state.df.empty:
        if 'unique case number in system' in st.session_state.df.columns:
            ids = sorted(st.session_state.df['unique case number in system'].unique(), reverse=True)
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

def page_archive():
    st.title("✅ Archive")
    if 'df' not in st.session_state or st.session_state.df.empty: 
        st.info("No data.")
        return
    
    if 'done' in st.session_state.df.columns:
        done = st.session_state.df[st.session_state.df['done'] == "Yes"]
        c1, c2 = st.columns(2)
        c1.metric("Total", len(done))
        c2.metric("Sum", f"£{pd.to_numeric(done['sum'], errors='coerce').sum():,.2f}")
        st.dataframe(done, width=None, hide_index=True)

def page_loans():
    st.title("Loan Management (Cloud)")
    if not get_access_token():
        st.warning("Please connect to OneDrive.")
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
    ).interactive(), width=None)

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
        st.dataframe(df, width=None)
        
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button("Download CSV", csv, "schedule.csv", "text/csv")

st.sidebar.title("Navigation")
pg = st.sidebar.radio("Go to:", ["CRM Dashboard", "Archive", "Loan Management", "Loan Calculator"])

if pg == "CRM Dashboard": page_crm()
elif pg == "Archive": page_archive()
elif pg == "Loan Management": page_loans()
elif pg == "Loan Calculator": page_calculator()
