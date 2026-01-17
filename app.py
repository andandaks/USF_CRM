import streamlit as st
import pandas as pd
import urllib
from sqlalchemy import create_engine, text

# --- PAGE CONFIG ---
st.set_page_config(page_title="UpShift Finance CRM", layout="wide")
st.title("📊 UpShift Finance CRM Dashboard")

# --- DATABASE CONNECTION (Cached) ---
@st.cache_resource
def init_connection():
    # Use secrets from Streamlit Cloud
    db_server = st.secrets["DB_SERVER"]
    db_database = st.secrets["DB_NAME"]
    db_username = st.secrets["DB_USER"]
    db_password = st.secrets["DB_PASSWORD"]
    
    # Connection String
    params = urllib.parse.quote_plus(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};" 
        f"SERVER={{tcp:{db_server},1433}};"
        f"DATABASE={{{db_database}}};"
        f"UID={{{db_username}}};"
        f"PWD={{{db_password}}};"
        "Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
    )
    return create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

try:
    engine = init_connection()
    st.success("Connected to Azure SQL", icon="✅")
except Exception as e:
    st.error(f"Connection failed: {e}")
    st.stop()

# --- SECTION 1: ADD NEW ENTITY ---
with st.expander("➕ Add New Case", expanded=False):
    with st.form("add_entity_form"):
        st.write("Enter details for the new case:")
        
        # COLUMNS: Update these to match your actual database columns
        c1, c2 = st.columns(2)
        client_name = c1.text_input("Client Name")
        case_status = c2.selectbox("Status", ["Open", "Pending", "Closed"])
        case_description = st.text_area("Case Description")
        
        submitted = st.form_submit_button("Submit New Case")
        
        if submitted:
            if client_name:
                try:
                    with engine.connect() as conn:
                        # UPDATE THIS QUERY to match your exact table schema
                        query = text("""
                            INSERT INTO crm_cases (ClientName, Status, Description) 
                            VALUES (:name, :status, :desc)
                        """)
                        conn.execute(query, {"name": client_name, "status": case_status, "desc": case_description})
                        conn.commit()
                    st.success("New case added successfully!")
                    st.rerun() # Refresh data immediately
                except Exception as e:
                    st.error(f"Error adding data: {e}")
            else:
                st.warning("Please fill in the Client Name.")

# --- SECTION 2: VIEW DATA ---
st.divider()
st.subheader("Existing Cases")

# Function to load data
def load_data():
    return pd.read_sql("SELECT * FROM crm_cases", engine)

# Reload button
if st.button("Refresh Data"):
    st.rerun()

# Display Data
try:
    df = load_data()
    st.dataframe(df, use_container_width=True)
    
    # Download Button
    st.download_button(
        label="Download as CSV",
        data=df.to_csv(index=False).encode('utf-8'),
        file_name='upshift_crm_data.csv',
        mime='text/csv',
    )
except Exception as e:
    st.error(f"Error reading database: {e}")
