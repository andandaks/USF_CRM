import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime, date
import random

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
    
    connection_url = f"mssql+pymssql://{db_username}:{db_password}@{db_server}/{db_database}"
    return create_engine(connection_url)

try:
    engine = init_connection()
    st.toast("Connected to Azure SQL", icon="✅")
except Exception as e:
    st.error(f"Connection failed: {e}")
    st.stop()

# --- HELPER: DATA CLEANING ---
def clean_num(val):
    """Converts empty strings to None (NULL) for SQL"""
    if not val or val == '':
        return None
    try:
        return int(float(val))
    except ValueError:
        return None

# --- TABS LAYOUT ---
tab_view, tab_add, tab_edit = st.tabs(["📂 View Data", "➕ Add New", "✏️ Edit Existing"])

# ==========================================
# TAB 1: VIEW DATA
# ==========================================
with tab_view:
    st.header("All Cases")
    
    # Refresh button
    if st.button("Refresh Table", key="refresh_view"):
        st.rerun()

    try:
        # Load data
        df = pd.read_sql("SELECT * FROM crm_cases", engine)
        st.dataframe(df, use_container_width=True)
    except Exception as e:
        st.error(f"Error reading database: {e}")

# ==========================================
# TAB 2: ADD NEW CASE
# ==========================================
with tab_add:
    st.header("Add New Case")
    with st.form("add_entity_form"):
        # Row 1
        c1, c2, c3, c4 = st.columns(4)
        unique_id = c1.text_input("Unique Case Number (Auto-generated if empty)", key="add_uid")
        date_added = c2.date_input("Date Added", value=date.today(), key="add_date")
        manager = c3.text_input("Manager", key="add_mgr")
        resp_entity = c4.text_input("Responsible Entity", key="add_resp")
        
        # Row 2
        c5, c6, c7, c8 = st.columns(4)
        comp_name = c5.text_input("Company Name", key="add_name")
        comp_num = c6.text_input("Company Number", key="add_num")
        phone = c7.text_input("Phone", key="add_ph")
        email = c8.text_input("Email", key="add_em")
        
        # Row 3
        c9, c10, c11, c12 = st.columns(4)
        prod_type = c9.selectbox("Product Type", ["Loan", "Credit", "Leasing", "Other"], key="add_prod")
        site = c10.text_input("Site/Website", key="add_site")
        amount_sum = c11.number_input("Sum", min_value=0.0, step=100.0, key="add_sum")
        equifax = c12.number_input("Equifax Score", min_value=0, step=1, key="add_eq")
        
        # Row 4
        st.write("#### Status Flags")
        f1, f2, f3, f4, f5, f6 = st.columns(6)
        has_pledge = f1.checkbox("Has Pledge", key="add_plg")
        returning = f2.checkbox("Returning Client", key="add_ret")
        is_done = f3.checkbox("Done", key="add_done")
        is_kyc = f4.checkbox("KYC", key="add_kyc")
        is_aml = f5.checkbox("AML", key="add_aml")
        soft_check = f6.checkbox("Soft Check", key="add_soft")
        
        # Row 5
        comment = st.text_area("Comment", key="add_cmt")
        
        if st.form_submit_button("Submit New Case", type="primary"):
            try:
                with engine.connect() as conn:
                    # Logic: Generate ID if empty
                    if not unique_id:
                        final_uid = random.randint(1000000000, 9999999999)
                    else:
                        final_uid = int(unique_id)

                    query = text("""
                        INSERT INTO crm_cases (
                            [unique case number in system], [date added], [responsible entity], 
                            [company name], [company number], [manager], [product type], 
                            [phone], [email], [site], [sum], [has pledge], [returning client], 
                            [comment], [done], [kyc], [aml], [soft_check], [equifax_score]
                        ) VALUES (
                            :uid, :date, :resp, :c_name, :c_num, :mgr, :prod, 
                            :ph, :em, :site, :sm, :plg, :ret, :cmt, :dn, :kyc, :aml, :sft, :eq
                        )
                    """)
                    
                    conn.execute(query, {
                        "uid": final_uid, "date": date_added, "resp": resp_entity,
                        "c_name": comp_name, "c_num": clean_num(comp_num), "mgr": manager,
                        "prod": prod_type, "ph": clean_num(phone), "em": email,
                        "site": site, "sm": amount_sum, "plg": has_pledge,
                        "ret": returning, "cmt": comment, "dn": is_done,
                        "kyc": is_kyc, "aml": is_aml, "sft": soft_check, "eq": equifax
                    })
                    conn.commit()
                st.success(f"Case #{final_uid} added!")
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")

# ==========================================
# TAB 3: EDIT EXISTING
# ==========================================
with tab_edit:
    st.header("Update Case Details")
    
    # 1. Select Box to choose case
    # Get list of cases for the dropdown
    try:
        cases_df = pd.read_sql("SELECT [unique case number in system], [company name] FROM crm_cases", engine)
        if cases_df.empty:
            st.warning("No cases found to edit.")
        else:
            # Create a list of options like: "102348 - Google"
            cases_df['label'] = cases_df['unique case number in system'].astype(str) + " - " + cases_df['company name']
            selected_label = st.selectbox("Select Case to Edit", cases_df['label'])
            
            # Extract the ID from the selection
            selected_id = int(selected_label.split(" - ")[0])
            
            # 2. Fetch current data for this ID
            current_data = pd.read_sql(f"SELECT * FROM crm_cases WHERE [unique case number in system] = {selected_id}", engine).iloc[0]

            # 3. Edit Form (Pre-filled)
            with st.form("edit_form"):
                st.info(f"Editing Case: {selected_id}")
                
                # We reuse the same layout, but set 'value' to current_data[...]
                e1, e2, e3 = st.columns(3)
                # Note: We convert SQL dates/decimals to Python types if needed
                new_date = e1.date_input("Date Added", value=pd.to_datetime(current_data['date added']))
                new_mgr = e2.text_input("Manager", value=current_data['manager'])
                new_resp = e3.text_input("Responsible Entity", value=current_data['responsible entity'])
                
                e4, e5, e6, e7 = st.columns(4)
                new_name = e4.text_input("Company Name", value=current_data['company name'])
                new_cnum = e5.text_input("Company Number", value=str(current_data['company number'] or ""))
                new_ph = e6.text_input("Phone", value=str(current_data['phone'] or ""))
                new_email = e7.text_input("Email", value=current_data['email'])

                e8, e9, e10, e11 = st.columns(4)
                # Helper to handle dropdown default index
                prod_opts = ["Loan", "Credit", "Leasing", "Other"]
                curr_prod = current_data['product type']
                prod_idx = prod_opts.index(curr_prod) if curr_prod in prod_opts else 0
                
                new_prod = e8.selectbox("Product Type", prod_opts, index=prod_idx)
                new_site = e9.text_input("Site", value=current_data['site'])
                new_sum = e10.number_input("Sum", value=float(current_data['sum'] or 0))
                new_eq = e11.number_input("Equifax", value=int(current_data['equifax_score'] or 0))

                # Checkboxes need boolean values
                st.write("#### Status")
                c_plg, c_ret, c_done, c_kyc, c_aml, c_sft = st.columns(6)
                new_plg = c_plg.checkbox("Has Pledge", value=bool(current_data['has pledge']))
                new_ret = c_ret.checkbox("Returning", value=bool(current_data['returning client']))
                new_done = c_done.checkbox("Done", value=bool(current_data['done']))
                new_kyc = c_kyc.checkbox("KYC", value=bool(current_data['kyc']))
                new_aml = c_aml.checkbox("AML", value=bool(current_data['aml']))
                new_sft = c_sft.checkbox("Soft Check", value=bool(current_data['soft_check']))
                
                new_cmt = st.text_area("Comment", value=current_data['comment'])

                # 4. Update Logic
                if st.form_submit_button("Update Case", type="primary"):
                    with engine.connect() as conn:
                        update_query = text("""
                            UPDATE crm_cases 
                            SET [date added]=:dt, [responsible entity]=:resp, [company name]=:name,
                                [company number]=:cnum, [manager]=:mgr, [product type]=:prod,
                                [phone]=:ph, [email]=:em, [site]=:site, [sum]=:sm,
                                [has pledge]=:plg, [returning client]=:ret, [comment]=:cmt,
                                [done]=:dn, [kyc]=:kyc, [aml]=:aml, [soft_check]=:sft, 
                                [equifax_score]=:eq
                            WHERE [unique case number in system] = :uid
                        """)
                        
                        conn.execute(update_query, {
                            "uid": selected_id, # Can't change ID
                            "dt": new_date, "resp": new_resp, "name": new_name,
                            "cnum": clean_num(new_cnum), "mgr": new_mgr, "prod": new_prod,
                            "ph": clean_num(new_ph), "em": new_email, "site": new_site,
                            "sm": new_sum, "plg": new_plg, "ret": new_ret, "cmt": new_cmt,
                            "dn": new_done, "kyc": new_kyc, "aml": new_aml, "sft": new_sft, 
                            "eq": new_eq
                        })
                        conn.commit()
                    st.success("Case updated successfully!")
                    st.rerun()

    except Exception as e:
        st.error(f"Error loading edit form: {e}")
