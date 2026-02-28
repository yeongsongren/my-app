import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl # Explicitly import to help the cloud find it

# 1. The Filename
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

if submit:
    try:
        target_sheet = "Sheet2" if "Table2" in table_choice else "Sheet3"
        formatted_date = date_val.strftime("%d-%m-%y")

        # Load the existing data
        # We use engine='openpyxl' to be 100% sure it uses the right library
        df = pd.read_excel(FILE_NAME, sheet_name=target_sheet, engine='openpyxl')

        # Create the new row as a Dictionary
        # This matches the column names automatically
        new_row = pd.DataFrame([{
            df.columns[0]: formatted_date,
            df.columns[1]: amount,
            df.columns[2]: entity,
            df.columns[3]: remarks
        }])

        # Combine old data with new row
        updated_df = pd.concat([df, new_row], ignore_index=True)

        # Save back to the file
        with pd.ExcelWriter(FILE_NAME, engine='openpyxl') as writer:
            updated_df.to_excel(writer, sheet_name=target_sheet, index=False)

        st.success(f"Successfully added to {target_sheet}!")
        st.balloons() # Just for fun!
        
    except Exception as e:
        st.error(f"Error: {e}")
