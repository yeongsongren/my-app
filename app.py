import streamlit as st
import xlwings as xw
from datetime import datetime
import os

# Configuration
FILE_PATH = r"C:\Users\yeong\OneDrive - MORE WATER SDN. BHD\MV_Documents\Finance\GuangFaBank Transactions.xlsx"

st.set_page_config(page_title="Finance Entry", page_icon="💰")
st.title("🏦 GuangFaBank Transaction Entry")

# 1. Form UI
with st.form("transaction_form", clear_on_submit=True):
    col1, col2 = st.columns(2)
    
    with col1:
        # Table Selection Dropdown
        table_selection = st.selectbox("Target Table", ["Table2 (Sheet2)", "Table3 (Sheet3)"])
        # Date Picker (Defaults to today)
        date_val = st.date_input("Transaction Date", datetime.now())
    
    with col2:
        # Entity Dropdown
        entity = st.selectbox("Entity", ["MV", "YEONG"])
        # Amount Input
        amount = st.number_input("Amount", min_value=0.0, step=0.01, format="%.2f")

    # Remarks Text Input
    remarks = st.text_input("Remarks (Order # / Notes)")
    
    submit = st.form_submit_button("Add to Excel")

# 2. Logic when button is clicked
if submit:
    try:
        # Map selection to actual names
        target_table = "Table2" if "Table2" in table_selection else "Table3"
        target_sheet = "Sheet2" if "Table2" in table_selection else "Sheet3"
        formatted_date = date_val.strftime("%d-%m-%y")

        # Excel Interaction
        wb = xw.Book(FILE_PATH)
        sheet = wb.sheets[target_sheet]
        table = sheet.api.ListObjects(target_table)
        
        # Add new row
        new_row = table.ListRows.Add()
        new_row.Range(1).Value = formatted_date
        new_row.Range(2).Value = amount
        new_row.Range(3).Value = entity
        new_row.Range(4).Value = remarks
        
        wb.save()
        st.success(f"Successfully added to {target_table}!")
        
    except Exception as e:
        st.error(f"Error: {e}")