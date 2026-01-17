import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime

# --- PAGE CONFIG ---
st.set_page_config(page_title="UpShift Finance CRM", layout="wide")
st.title("📊 UpShift Finance CRM Dashboard")

# --- DATABASE CONNECTION (pymssql) ---
@st.cache_resource
def init_connection():
    db_server = st.secrets["DB_SERVER"]
    db_database = st.secrets["DB_NAME"]
    db_username = st.secrets["DB_USER"]
    db_password = st.secrets["DB_PASSWORD"]
    
    # We use pymssql here to avoid driver/EULA issues on Linux
    connection_url = f"mssql+pymssql://{db_username}:{db_password}@{db_server}/{db_database}"
    return create_engine(connection_url)

try:
    engine = init_connection()
    st.success("Connected to Azure SQL (via pymssql)", icon="✅")
except Exception as e:
    st.error(f"Connection failed: {e}")
    st.stop()

# --- SECTION 1: ADD NEW CASE ---
with st.expander("➕ Add New Case", expanded=False):
    with st.form("add_entity_form"):
        st.write("### New Case Details")
        
        # Row 1: Core Info
        c1, c2, c3, c4 = st.columns(4)
        unique_id = c1.text_input("Unique Case Number (Optional)", help="Leave empty if auto-generated")
        date_added = c2.date_input("Date Added", value=datetime.today())
        manager = c3.text_input("Manager")
        resp_entity = c4.text_input("Responsible Entity")
        
        # Row 2: Client Info
        c5, c6, c7, c8 = st.columns(4)
        comp_name = c5.text_input("Company Name")
        comp_num = c6.text_input("Company Number")
        phone = c7.text_input("Phone")
        email = c8.text_input("Email")
        
        # Row 3: Product & Financials
        c9, c10, c11, c12 = st.columns(4)
        prod_type = c9.selectbox("Product Type", ["Loan", "Credit", "Leasing", "Other"]) # Adjust options as needed
        site = c10.text_input("Site/Website")
        amount_sum = c11.number_input("Sum", min_value=0.0, step=100.0)
        equifax = c12.number_input("Equifax Score", min_value=0, step=1)
        
        # Row 4: Status Flags
        st.write("#### Status & Checks")
        f1, f2, f3, f4, f5, f6, f7 = st.columns(7)
        has_pledge = f1.checkbox("Has Pledge")
        returning = f2.checkbox("Returning Client")
        is_done = f3.checkbox("Done")
        is_kyc = f4.checkbox("KYC")
        is_aml = f5.checkbox("AML")
        soft_check = f6.checkbox("Soft Check")
        
        # Row 5: Comments
        comment = st.text_area("Comment")
        
        submitted = st.form_submit_button("Submit New Case", type="primary")
        
        if submitted:
            try:
                with engine.connect() as conn:
                    # We map the inputs to the columns. 
                    # Note: We convert checkboxes (True/False) to 1/0 for SQL BIT fields automatically usually, 
                    # but being explicit helps.
                    
                    query = text("""
                        INSERT INTO crm_cases (
                            [unique case number in system], [date added], [responsible entity], 
                            [company name], [company number], [manager], [product type], 
                            [phone], [email], [site], [sum], [has pledge], [returning client], 
                            [comment], [done], [kyc], [aml], [soft_check], [equifax_score]
                        ) 
                        VALUES (
                            :uid, :date, :resp, :c_name, :c_num, :mgr, :prod, 
                            :ph, :em, :site, :sm, :plg, :ret, :cmt, :dn, :kyc, :aml, :sft, :eq
                        )
                    """)
                    
                    # Handle the optional ID
                    uid_val = unique_id if unique_id else None 
                    
                    conn.execute(query, {
                        "uid": uid_val,
                        "date": date_added,
                        "resp": resp_entity,
                        "c_name": comp_name,
                        "c_num": comp_num,
                        "mgr": manager,
                        "prod": prod_type,
                        "ph": phone,
                        "em": email,
                        "site": site,
                        "sm": amount_sum,
                        "plg": has_pledge,
                        "ret": returning,
                        "cmt": comment,
                        "dn": is_done,
                        "kyc": is_kyc,
                        "aml": is_aml,
                        "sft": soft_check,
                        "eq": equifax
                    })
                    conn.commit()
                st.success(f"Case for '{comp_name}' added successfully!")
                st.rerun()
            except Exception as e:
                st.error(f"Error adding data: {e}")

# --- SECTION 2: VIEW DATA ---
st.divider()
st.subheader("Existing Cases")

if st.button("Refresh Data"):
    st.rerun()

try:
    # We use brackets [] around column names because they contain spaces
    df = pd.read_sql("SELECT * FROM crm_cases", engine)
    st.dataframe(df, use_container_width=True)
except Exception as e:
    st.error(f"Error reading database: {e}")
