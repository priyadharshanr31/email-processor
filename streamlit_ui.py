import streamlit as st
import asyncio
import os
import pandas as pd
from email_processor import fetch_email
from process_files import process_latest_file
from email_sender import send_email_async
from config import ATTACHMENT_FOLDER, OUTPUT_FOLDER

# Customizing the Page Layout
st.set_page_config(page_title="Automated Statement Processor", page_icon="📩", layout="wide")

# Header with Style
st.markdown("""
    <style>
        .big-title {font-size: 32px; font-weight: bold; color: #4A90E2; text-align: center;}
        .sub-title {font-size: 24px; font-weight: bold; color: #333; text-align: center;}
        .section-divider {border-top: 2px solid #4A90E2; margin-top: 10px; margin-bottom: 10px;}
    </style>
""", unsafe_allow_html=True)

st.markdown("<p class='big-title'>📩 Automated Statement Processor</p>", unsafe_allow_html=True)

# Layout for Email Credentials
col1, col2 = st.columns([1, 1])
with col1:
    email_user = st.text_input("📧 Enter Email Address", placeholder="your@email.com")
with col2:
    email_pass = st.text_input("🔑 Enter Email Password", type="password")

# Fetch Latest Email
st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
if st.button("📥 Fetch Latest Email", use_container_width=True):
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

# Process File Section
if "file_path" in st.session_state:
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<p class='sub-title'>⚙️ Process File</p>", unsafe_allow_html=True)
    start_row = st.number_input("🔢 Enter Row Number to Insert Data", min_value=1, value=6)
    
    if st.button("🚀 Process File", use_container_width=True):
        with st.spinner("Processing file..."):
            processed_path, extracted_data = asyncio.run(process_latest_file(st.session_state.file_path, start_row))
        
        if processed_path:
            st.success(f"✅ Processed file saved: {processed_path}")
            st.session_state.processed_path = processed_path
            st.session_state.extracted_data = extracted_data
        else:
            st.error("❌ Processing failed.")

# ✅ Preview Extracted Data (WITHOUT "Extracted_Desc")
if "extracted_data" in st.session_state:
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<p class='sub-title'>📄 Preview 1: Extracted Data</p>", unsafe_allow_html=True)
    extracted_df = pd.DataFrame(st.session_state.extracted_data[1:], columns=st.session_state.extracted_data[0])
    st.dataframe(extracted_df, use_container_width=True)

# Preview Processed File Data
if "processed_path" in st.session_state:
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<p class='sub-title'>📝 Preview 2: Final Processed Output</p>", unsafe_allow_html=True)
    try:
        df_processed = pd.read_excel(st.session_state.processed_path, engine="openpyxl")
        st.dataframe(df_processed, use_container_width=True)
    except Exception as e:
        st.error(f"❌ Error loading processed file: {e}")

# Send Email Button
if "processed_path" in st.session_state and "sender_email" in st.session_state:
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<p class='sub-title'>📤 Send Processed File via Email</p>", unsafe_allow_html=True)
    
    if st.button("📤 Send Processed File", use_container_width=True):
        if not email_user or not email_pass:
            st.warning("⚠️ Please enter your email and password.")
        else:
            with st.spinner("Sending email..."):
                result = asyncio.run(send_email_async(email_user, email_pass, st.session_state.sender_email, st.session_state.processed_path))
            if result:
                st.success("✅ Email sent successfully!")
            else:
                st.error("❌ Email sending failed.")
