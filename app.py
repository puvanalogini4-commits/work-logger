import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Work Record System")
st.title("🏫 School Work Logger")

# Establish Connection
conn = st.connection("gsheets", type=GSheetsConnection)

# --- FORM SECTION ---
with st.form("entry_form", clear_on_submit=True):
    date = st.date_input("Date")
    school = st.text_input("School Name")
    work = st.text_area("Work Done")
    amend = st.text_input("Amendment")
    dist = st.number_input("Distance from Office (km)", min_value=0.0)
    
    submit = st.form_submit_button("Submit Record")

    if submit:
        # Read current data
        existing_data = conn.read(worksheet="Sheet1")
        
        # Create new row
        new_data = pd.DataFrame([{
            "Date": str(date), "School": school, 
            "Work Done": work, "Amendment": amend, "Distance": dist
        }])
        
        # Combine and Update
        updated_df = pd.concat([existing_data, new_data], ignore_index=True)
        conn.update(worksheet="Sheet1", data=updated_df)
        st.success("✅ Data saved to Google Sheets!")

# --- DOWNLOAD SECTION ---
st.divider()
st.subheader("📊 Month-End Export")

if st.button("Generate Excel Report"):
    data = conn.read(worksheet="Sheet1")
    data['Date'] = pd.to_datetime(data['Date'])
    
    # Filter for the current month
    current_month = datetime.now().month
    monthly_df = data[data['Date'].dt.month == current_month]
    
    # Export to Excel
    file_path = "monthly_report.xlsx"
    monthly_df.to_excel(file_path, index=False)
    
    with open(file_path, "rb") as file:
        st.download_button(
            label="📥 Download Excel File",
            data=file,
            file_name=f"Report_{datetime.now().strftime('%B_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )