import streamlit as st
import pandas as pd
from datetime import datetime
import requests
import base64

# Config from Streamlit Secrets
TOKEN = st.secrets["GITHUB_TOKEN"]
USER = st.secrets["GITHUB_USER"]
REPO = st.secrets["GITHUB_REPO"]
FILE_NAME = "GuangFaBank Transactions.xlsx"

st.set_page_config(page_title="Cloud Finance Entry", page_icon="💰")
st.title("📲 Live Finance Sync")

with st.form("transaction_form", clear_on_submit=True):
    col1, col2 = st.columns(2)
    with col1:
        table_choice = st.selectbox("Target Table", ["Table2 (Sheet2)", "Table3 (Sheet3)"])
        date_val = st.date_input("Date", datetime.now())
    with col2:
        entity = st.selectbox("Entity", ["MV", "YEONG"])
        amount = st.number_input("Amount", min_value=0.0, format="%.2f")
    remarks = st.text_input("Remarks")
    submit = st.form_submit_button("Save & Sync to GitHub")

if submit:
    try:
        target_sheet = "Sheet2" if "Table2" in table_choice else "Sheet3"
        
        # 1. Update Local Copy
        df = pd.read_excel(FILE_NAME, sheet_name=target_sheet, engine='openpyxl')
        new_row = pd.DataFrame([{df.columns[0]: date_val.strftime("%d-%m-%y"), df.columns[1]: amount, df.columns[2]: entity, df.columns[3]: remarks}])
        updated_df = pd.concat([df, new_row], ignore_index=True)
        with pd.ExcelWriter(FILE_NAME, engine='openpyxl') as writer:
            updated_df.to_excel(writer, sheet_name=target_sheet, index=False)

        # 2. Sync to GitHub (The "Push")
        url = f"https://api.github.com/repos/{USER}/{REPO}/contents/{FILE_NAME}"
        
        # Get the 'sha' (GitHub needs this to overwrite a file)
        current_file = requests.get(url, headers={"Authorization": f"token {TOKEN}"}).json()
        sha = current_file['sha']

        # Encode the updated Excel file to Base64
        with open(FILE_NAME, "rb") as f:
            content = base64.b64encode(f.read()).decode("utf-8")

        data = {
            "message": f"Added transaction via Streamlit: {remarks}",
            "content": content,
            "sha": sha
        }

        response = requests.put(url, json=data, headers={"Authorization": f"token {TOKEN}"})

        if response.status_code == 200:
            st.success("✅ Saved and Synced to GitHub!")
            st.balloons()
        else:
            st.error(f"Sync failed: {response.json()}")

    except Exception as e:
        st.error(f"Error: {e}")
