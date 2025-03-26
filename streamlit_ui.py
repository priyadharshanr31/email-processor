import streamlit as st
import pandas as pd
import os
from email_processor import check_email_and_download
from process_files import process_latest_file
from email_sender import send_email_with_attachment

# Initialize session state variables
if "file_path" not in st.session_state:
    st.session_state.file_path = None
if "sender_email" not in st.session_state:
    st.session_state.sender_email = None
if "processed_file" not in st.session_state:
    st.session_state.processed_file = None
if "extracted_data" not in st.session_state:
    st.session_state.extracted_data = None

# Streamlit UI
st.title("📧 Automated Email Processor & File Processor")

# Step 1: Email Login
st.header("Step 1: Enter Email Credentials")
email_user = st.text_input("Email Address", type="default")
email_pass = st.text_input("App Password", type="password")
check_email = st.button("Check for Unread Emails")

# Step 2: Process Email & Extract Data
if check_email:
    if email_user and email_pass:
        file_path, sender_email = check_email_and_download(email_user, email_pass)

        if file_path:
            st.session_state.file_path = file_path
            st.session_state.sender_email = sender_email

            st.success(f"✅ Email received from: {sender_email}")
            st.info(f"📂 Downloaded attachment: {file_path}")

            # Extract and show only filtered rows
            _, extracted_data = process_latest_file(file_path, preview_only=True)

            if extracted_data is not None and not extracted_data.empty:
                st.session_state.extracted_data = extracted_data  # Store extracted data
                st.subheader("Extracted Data Preview (Filtered by Status)")
                st.dataframe(extracted_data)  # Show only extracted rows
            else:
                st.warning("⚠️ No matching data found in the received file.")
        else:
            st.error("❌ No valid email found.")
    else:
        st.warning("⚠️ Please enter email credentials.")

# Step 3: Get Start Row Input
st.header("Step 2: Process with Template")
start_row = st.number_input("Enter the starting row to process:", min_value=1, value=10, step=1)
process_button = st.button("Process Files")

# Step 4: Process Files
if process_button:
    if st.session_state.file_path:
        processed_file, _ = process_latest_file(st.session_state.file_path, start_row)
        if processed_file:
            st.session_state.processed_file = processed_file
            st.success(f"✅ Processed file saved at: {processed_file}")
        else:
            st.error("❌ Failed to process files.")
    else:
        st.warning("⚠️ Please check emails and extract data first.")

# Step 5: Final Processed File Preview
if st.session_state.processed_file:
    st.header("Step 3: Final Processed File Preview")
    final_df = pd.read_excel(st.session_state.processed_file)
    st.dataframe(final_df)  # Show the full final output file

# Step 6: Send Email
st.header("Step 4: Send Email with Processed File")
send_email = st.button("Send Processed File")

if send_email:
    if st.session_state.file_path and st.session_state.processed_file:
        email_status = send_email_with_attachment(
            st.session_state.processed_file, 
            st.session_state.sender_email, 
            email_user,  # Pass sender email
            email_pass   # Pass email password
        )
        if email_status:
            st.success("✅ Email sent successfully!")
        else:
            st.error("❌ Failed to send email.")
    else:
        st.warning("⚠️ Please process files before sending.")
