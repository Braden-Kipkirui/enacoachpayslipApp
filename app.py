import streamlit as st
import pandas as pd
from utils import generate_and_send_payslip

st.set_page_config(page_title="Payslip Generator", layout="centered")
st.title("📩 ENA Coach Payslip Generator & Email Sender")

uploaded_file = st.file_uploader("Upload Payroll Excel File", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        if 'Month' not in df.columns or 'Email' not in df.columns:
            st.error("Missing 'Month' or 'Email' columns in file.")
        else:
            months = sorted(df['Month'].dropna().unique())
            selected_month = st.selectbox("Select Month to Send Payslips", months)
            filtered_df = df[df['Month'] == selected_month]

            st.success(f"Found {len(filtered_df)} employees for {selected_month}")
            st.dataframe(filtered_df)

            st.subheader("📧 Email Settings")
            sender_email = st.text_input("Sender Email")
            sender_password = st.text_input("App Password", type="password")

            if st.button("📨 Send Payslips"):
                if not sender_email or not sender_password:
                    st.error("Email and password are required.")
                else:
                    for index, row in filtered_df.iterrows():
                        try:
                            generate_and_send_payslip(row, sender_email, sender_password, selected_month)
                            st.success(f"Payslip sent to {row['Email']}")
                        except Exception as e:
                            st.error(f"Error sending to {row['Email']}: {e}")

    except Exception as e:
        st.error(f"Failed to process file: {e}")