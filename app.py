import streamlit as st
import pandas as pd
import pyodbc
import urllib
import msal
import requests
import base64
from datetime import datetime

# --- CONFIGURATION ---
try:
    # Database
    DB_SERVER = st.secrets["database"]["server"]
    DB_NAME = st.secrets["database"]["name"]
    DB_USER = st.secrets["database"]["user"]
    DB_PASSWORD = st.secrets["database"]["password"]
    DB_DRIVER = 'ODBC Driver 17 for SQL Server'
    
    # Azure
    CLIENT_ID = st.secrets["azure"]["client_id"]
    CLIENT_SECRET = st.secrets["azure"]["client_secret"]
    TENANT_ID = st.secrets["azure"]["tenant_id"]
    AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
    SHARE_LINK = st.secrets["azure"]["sharepoint_link"] # <--- NEW
except Exception:
    st.error("Missing Secrets! Please check .streamlit/secrets.toml")
    st.stop()

# --- DATABASE FUNCTIONS ---
def get_db_connection():
    params = urllib.parse.quote_plus(
        f"DRIVER={{{DB_DRIVER}}};SERVER={DB_SERVER},1433;DATABASE={DB_NAME};"
        f"UID={DB_USER};PWD={DB_PASSWORD};Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
    )
    return pyodbc.connect(f"MSSQL+pyodbc:///?odbc_connect={params}")

def load_data():
    conn = get_db_connection()
    df = pd.read_sql("SELECT * FROM Cases", conn)
    conn.close()
    return df

def save_new_case(data):
    conn = get_db_connection()
    cursor = conn.cursor()
    placeholders = ",".join(["?"] * len(data))
    columns = ",".join(data.keys())
    values = list(data.values())
    try:
        cursor.execute(f"INSERT INTO Cases ({columns}) VALUES ({placeholders})", values)
        conn.commit()
        return True
    except Exception as e:
        st.error(f"Database Error: {e}")
        return False
    finally:
        conn.close()

# --- NEW ONEDRIVE FUNCTIONS (LINK RESOLVER) ---

def get_graph_access_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return result.get("access_token")

def get_drive_item_from_link(token, url):
    """
    Converts a SharePoint sharing link into a usable Drive ID and Item ID
    """
    # 1. Base64 encode the URL according to Graph API specs
    base64_value = base64.b64encode(url.encode("utf-8")).decode("utf-8")
    encoded_url = "u!" + base64_value.replace("/", "_").replace("+", "-").rstrip("=")
    
    # 2. Call the Shares API
    api_url = f"https://graph.microsoft.com/v1.0/shares/{encoded_url}/driveItem"
    headers = {'Authorization': 'Bearer ' + token}
    
    response = requests.get(api_url, headers=headers)
    if response.status_code == 200:
        data = response.json()
        return data["parentReference"]["driveId"], data["id"]
    else:
        st.error(f"Could not resolve SharePoint link. Error: {response.text}")
        return None, None

def create_folder_and_upload(file_obj, case_id):
    token = get_graph_access_token()
    if not token: return

    # 1. Resolve the Master Folder from the link provided in secrets
    drive_id, master_folder_id = get_drive_item_from_link(token, SHARE_LINK)
    
    if not drive_id or not master_folder_id:
        return # Error already handled in helper function

    headers = {'Authorization': 'Bearer ' + token}

    # 2. Create the specific CASE folder INSIDE the Master Folder
    # POST /drives/{drive-id}/items/{master-folder-id}/children
    create_folder_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{master_folder_id}/children"
    folder_body = {
        "name": str(case_id),
        "folder": {},
        "@microsoft.graph.conflictBehavior": "fail" # Use "fail" so we just get the ID if it exists
    }
    
    response = requests.post(create_folder_url, headers=headers, json=folder_body)
    
    if response.status_code == 201:
        case_folder_id = response.json()["id"]
    elif response.status_code == 409:
        # Folder exists, we need to fetch its ID to upload into it
        # GET /drives/{drive-id}/items/{master-folder-id}/children/{case_id}
        # Note: Graph doesn't let you get child by name easily in one call sometimes, 
        # but 409 usually returns the ID of the existing item in the error body? No.
        # We must query for it.
        query_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{master_folder_id}/children"
        query_res = requests.get(query_url, headers=headers)
        items = query_res.json().get('value', [])
        case_folder_id = next((item['id'] for item in items if item['name'] == str(case_id)), None)
    else:
        st.error(f"Error creating case folder: {response.text}")
        return

    if not case_folder_id:
        st.error("Could not locate Case Folder ID.")
        return

    # 3. Upload File into the new Case Folder
    # PUT /drives/{drive-id}/items/{case-folder-id}:/{filename}:/content
    upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{case_folder_id}:/{file_obj.name}:/content"
    
    file_obj.seek(0)
    upload_res = requests.put(upload_url, headers=headers, data=file_obj)
    
    if upload_res.status_code in [200, 201]:
        st.success(f"✅ Uploaded {file_obj.name} to Case Folder {case_id}")
    else:
        st.error(f"❌ Failed to upload {file_obj.name}")

# --- UI ---
st.set_page_config(page_title="Case Manager", layout="wide")
st.title("🏦 UpShift Case Management")

page = st.sidebar.radio("Navigate", ["Dashboard", "Add New Case"])

if page == "Dashboard":
    try:
        df = load_data()
        
        # --- SALES ---
        st.subheader("🔹 Responsible: Sales")
        edited_sales = st.data_editor(df[df['responsible_entity'] == 'Sales'], key="s", use_container_width=True, disabled=["unique_case_number"])
        
        # --- UNDERWRITER ---
        st.subheader("🔸 Responsible: Underwriter")
        edited_uw = st.data_editor(df[df['responsible_entity'] == 'Underwriter'], key="u", use_container_width=True, disabled=["unique_case_number"])
        
        # --- CLIENT ---
        st.subheader("▫️ Responsible: Client")
        edited_client = st.data_editor(df[df['responsible_entity'] == 'Client'], key="c", use_container_width=True, disabled=["unique_case_number"])

        if st.button("Save All Changes"):
            all_edited = pd.concat([edited_sales, edited_uw, edited_client])
            conn = get_db_connection()
            cursor = conn.cursor()
            
            # Simple Update Loop
            progress = st.progress(0)
            total = len(all_edited)
            for i, (index, row) in enumerate(all_edited.iterrows()):
                cursor.execute("""
                    UPDATE Cases SET 
                    responsible_entity=?, company_name=?, manager=?, sum_value=?, 
                    comment=?, done=?, kyc=?, aml=?, soft_check=?, equifax_score=?
                    WHERE unique_case_number=?
                """, (
                    row['responsible_entity'], row['company_name'], row['manager'], row['sum_value'],
                    row['comment'], row['done'], row['kyc'], row['aml'], row['soft_check'], row['equifax_score'],
                    row['unique_case_number']
                ))
                progress.progress((i + 1) / total)
            conn.commit()
            conn.close()
            st.success("Database Updated")
            st.rerun()

    except Exception as e:
        st.error(f"Data Load Error: {e}")

elif page == "Add New Case":
    st.header("Add New Entity")
    with st.form("new", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            unique_id = st.number_input("Unique Case Number", min_value=1, step=1)
            date_added = st.date_input("Date", datetime.now())
            responsible = st.selectbox("Responsible", ["Sales", "Underwriter", "Client"])
            company_name = st.text_input("Company Name")
            company_num = st.text_input("Company Number")
            manager = st.text_input("Manager")
        with col2:
            prod_type = st.selectbox("Product", ["Term Loan", "Credit Line"])
            phone = st.text_input("Phone")
            email = st.text_input("Email")
            site = st.text_input("Site")
            sum_val = st.number_input("Sum", min_value=0.0)
            pledge = st.selectbox("Pledge", ["Yes", "No"])
        with col3:
            ret_client = st.selectbox("Returning", ["Yes", "No"])
            comment = st.text_area("Comment")
            done = st.selectbox("Done", ["Yes", "No"])
            kyc = st.selectbox("KYC", ["N/A", "Passed", "Failed"])
            aml = st.selectbox("AML", ["N/A", "Passed", "Failed"])
            soft = st.selectbox("Soft Check", ["N/A", "Passed"])
            equifax = st.number_input("Equifax", 0, 999)

        uploaded_files = st.file_uploader("Upload Files", accept_multiple_files=True)
        submitted = st.form_submit_button("Create Case")

        if submitted:
            new_data = {
                "unique_case_number": unique_id, "date_added": date_added, "responsible_entity": responsible,
                "company_name": company_name, "company_number": company_num, "manager": manager,
                "product_type": prod_type, "phone": phone, "email": email, "site": site,
                "sum_value": sum_val, "has_pledge": pledge, "returning_client": ret_client,
                "comment": comment, "done": done, "kyc": kyc, "aml": aml,
                "soft_check": soft, "equifax_score": equifax
            }
            
            if save_new_case(new_data):
                st.success(f"Case {unique_id} created in DB.")
                if uploaded_files:
                    with st.spinner("Uploading to OneDrive..."):
                        for f in uploaded_files:
                            create_folder_and_upload(f, case_id=unique_id)
            else:
                st.error("Case ID likely exists already.")
