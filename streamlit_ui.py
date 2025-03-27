import streamlit as st
import asyncio
import os
import pandas as pd
from email_processor import fetch_email
from process_files import process_latest_file
from email_sender import send_email_async
from config import ATTACHMENT_FOLDER, OUTPUT_FOLDER



# Streamlit App Title
st.title("📩 Automated Statement Processor")



# User Input for Email Credentials
email_user = st.text_input("📧 Enter Email Address", type="default")
email_pass = st.text_input("🔑 Enter Email Password", type="password")



# Button to Check Email and Download Attachment
if st.button("📥 Fetch Latest Email"):
    if not email_user or not email_pass:
        st.warning("⚠️ Please enter both email and password.")
    else:
        with st.spinner("Fetching latest email..."):
            file_path, sender_email = asyncio.run(fetch_email(email_user, email_pass))
        
        if file_path:
            st.success(f"✅ Attachment downloaded: {file_path}")
            st.session_state.file_path = file_path
            st.session_state.sender_email = sender_email
        else:
            st.error("❌ No valid attachment found.")



# User Input for Row Number
if "file_path" in st.session_state:
    start_row = st.number_input("🔢 Enter Row Number to Insert Data", min_value=1, value=6)

    # Process File Button
    if st.button("⚙️ Process File"):
        with st.spinner("Processing file..."):
            processed_path, extracted_data = asyncio.run(process_latest_file(st.session_state.file_path, start_row))
        
        if processed_path:
            st.success(f"✅ Processed file saved: {processed_path}")
            st.session_state.processed_path = processed_path
            st.session_state.extracted_data = extracted_data
        else:
            st.error("❌ Processing failed.")



# Preview Extracted Data
if "extracted_data" in st.session_state:
    st.subheader("📄 Preview 1: Extracted Data")
    st.dataframe(st.session_state.extracted_data)


# Preview Processed File Data
if "processed_path" in st.session_state:
    st.subheader("📝 Preview 2: Final Processed Output")
    try:
        df_processed = pd.read_excel(st.session_state.processed_path, engine="openpyxl")
        st.dataframe(df_processed)
    except Exception as e:
        st.error(f"❌ Error loading processed file: {e}")



# Send Email Button
if "processed_path" in st.session_state and "sender_email" in st.session_state:
    if st.button("📤 Send Processed File via Email"):
        if not email_user or not email_pass:
            st.warning("⚠️ Please enter your email and password.")
        else:
            with st.spinner("Sending email..."):
                result = asyncio.run(send_email_async(
                    email_user, 
                    email_pass, 
                    st.session_state.sender_email, 
                    st.session_state.processed_path
                ))

            if result:
                st.success("✅ Email sent successfully!")
            else:
                st.error("❌ Email sending failed.")
