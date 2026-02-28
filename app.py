import streamlit as st
import pandas as pd
from datetime import datetime
import os

# 1. The Filename (Since it's in the same GitHub folder)
FILE_NAME = "GuangFaBank Transactions.xlsx"

st.set_page_config(page_title="Cloud Finance Entry", page_icon="☁️")
st.title("📲 Mobile Transaction Entry")

# 2. The Form UI
with st.form("transaction_form", clear_on_submit=True):
    col1, col2 = st.columns(2)
    
    with col1:
        table_choice = st.selectbox("Target Table", ["Table2 (Sheet2)", "Table3 (Sheet3)"])
        date_val = st.date_input("Date", datetime.now())
    
    with col2:
        entity = st.selectbox("Entity", ["MV", "YEONG"])
        amount = st.number_input("Amount", min_value=0.0, format="%.2f")

    remarks = st.text_input("Remarks")
    submit = st.form_submit_button("Save to Cloud Excel")

# 3. Logic for Cloud Saving
if submit:
    try:
        # Determine Sheet and Table
        target_sheet = "Sheet2" if "Table2" in table_choice else "Sheet3"
        formatted_date = date_val.strftime("%d-%m-%y")

        # Load the data using Pandas (Better for Cloud)
        # We read the specific sheet
        df = pd.read_excel(FILE_NAME, sheet_name=target_sheet)

        # Create the new row
        new_data = {
            df.columns[0]: formatted_date,
            df.columns[1]: amount,
            df.columns[2]: entity,
            df.columns[3]: remarks
        }

        # Append the new row
        df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)

        # Save back to the Excel file on the server
        with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=target_sheet, index=False)

        st.success(f"Added to {target_sheet}! (Note: This is saved on the cloud version only)")
        
    except Exception as e:
        st.error(f"Error: {e}")
