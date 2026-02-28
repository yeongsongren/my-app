import streamlit as st
import openpyxl
from datetime import datetime
import requests
import base64
import io

# Config from Streamlit Secrets
TOKEN = st.secrets["GITHUB_TOKEN"]
USER = st.secrets["GITHUB_USER"]
REPO = st.secrets["GITHUB_REPO"]
FILE_NAME = "GuangFaBank Transactions.xlsx"

st.set_page_config(page_title="Safe Finance Sync", page_icon="🔐")
st.title("📲 Safe Transaction Entry")
st.info("This version preserves your Formulas, Formatting, and all Sheets.")

with st.form("transaction_form", clear_on_submit=True):
    col1, col2 = st.columns(2)
    with col1:
        # User selects the table/sheet
        table_choice = st.selectbox("Target Sheet", ["Sheet2", "Sheet3"])
        date_val = st.date_input("Date", datetime.now())
    with col2:
        entity = st.selectbox("Entity", ["MV", "YEONG"])
        amount = st.number_input("Amount", min_value=0.0, format="%.2f")
    remarks = st.text_input("Remarks")
    submit = st.form_submit_button("Save & Sync")

if submit:
    try:
        # 1. LOAD the existing workbook into memory (Preserving everything)
        wb = openpyxl.load_workbook(FILE_NAME)
        
        if table_choice in wb.sheetnames:
            sheet = wb[table_choice]
            
            # Find the first empty row in Column A
            new_row = sheet.max_row + 1
            
            # Insert Data
            sheet.cell(row=new_row, column=1).value = date_val.strftime("%d-%m-%y")
            sheet.cell(row=new_row, column=2).value = amount
            sheet.cell(row=new_row, column=3).value = entity
            sheet.cell(row=new_row, column=4).value = remarks
            
            # Save the workbook to a temporary "buffer"
            buffer = io.BytesIO()
            wb.save(buffer)
            content_to_upload = buffer.getvalue()
            
            # 2. Sync back to GitHub
            url = f"https://api.github.com/repos/{USER}/{REPO}/contents/{FILE_NAME}"
            current_file = requests.get(url, headers={"Authorization": f"token {TOKEN}"}).json()
            sha = current_file['sha']

            # Encode and Push
            encoded_content = base64.b64encode(content_to_upload).decode("utf-8")
            data = {
                "message": f"Added row to {table_choice}",
                "content": encoded_content,
                "sha": sha
            }

            response = requests.put(url, json=data, headers={"Authorization": f"token {TOKEN}"})

            if response.status_code == 200:
                # IMPORTANT: Also save a local copy for the session
                with open(FILE_NAME, "wb") as f:
                    f.write(content_to_upload)
                st.success(f"✅ Successfully updated {table_choice} without losing formatting!")
            else:
                st.error(f"GitHub Sync Error: {response.json()}")
        else:
            st.error(f"Error: {table_choice} not found in Excel file.")

    except Exception as e:
        st.error(f"Error: {e}")
