📧 Automated Email Processor & File Merger
📌 Project Overview
This project automates the process of receiving, extracting, processing, and emailing statement files. It checks unread emails for attachments, extracts relevant data, merges it into a predefined template, and then sends the processed file back to the sender.

🛠 Features
✅ Automatically fetches unread emails with attachments.
✅ Extracts rows from the received file based on a predefined filter.
✅ Processes the extracted data by inserting it into a template file at a user-defined starting row.
✅ Shows previews of the extracted data and the final processed file.
✅ Sends the processed file via email to the original sender.

📂 Project Structure
bash
Copy
Edit
📁 Project Root
│── 📄 config.py                 # Configuration settings for email and file processing
│── 📄 email_processor.py        # Fetches unread emails and downloads attachments
│── 📄 email_sender.py           # Sends the processed file via email
│── 📄 process_files.py          # Processes and merges files with the template
│── 📄 streamlit_ui.py           # Streamlit UI for user interaction
│── 📁 attachments               # Stores downloaded email attachments
│── 📁 output                    # Stores processed output files
│── 📄 template.xlsx             # Template file where extracted data is inserted
│── 📄 requirements.txt          # Required Python packages
│── 📄 README.md                 # Project documentation
🔧 Setup Instructions
1️⃣ Install Dependencies
Ensure you have Python installed, then run:

sh
Copy
Edit
pip install -r requirements.txt
2️⃣ Configure Email Settings
Modify config.py with your email credentials and preferred settings:

python
Copy
Edit
EMAIL_SERVER = "imap.gmail.com"
EMAIL_USER = "your_email@gmail.com"
EMAIL_PASS = "your_app_password"
EMAIL_FOLDER = "INBOX"
EMAIL_SUBJECT = "Statements"
3️⃣ Run the Streamlit UI
Start the Streamlit interface with:

sh
Copy
Edit
streamlit run streamlit_ui.py
🏗 How It Works
Step 1: Check for Unread Emails
The system connects to the email server and searches for unread emails with a specific subject (e.g., "Statements").

If an email contains an Excel attachment, it downloads it to the attachments/ folder.

Step 2: Extract & Process Data
Reads the downloaded file and filters rows where the 'Status' column contains "Y" (or similar values).

Displays the extracted data preview in the UI.

Step 3: Process with Template
The user selects the row where the extracted data should be inserted.

The filtered rows are inserted at the specified row, and the original template data is retained below.

The processed file is saved in the output/ folder.

Step 4: Send Processed File
The processed file is emailed back to the sender.

A confirmation message appears once the email is sent successfully.

📚 Libraries Used
Library	Purpose
pandas	Data manipulation and processing
streamlit	Interactive UI for user input
openpyxl	Reading and writing Excel files
imaplib	Fetching unread emails
email	Processing email content
smtplib	Sending emails
mimetypes	Handling file attachments
⚠️ Notes
Ensure your email account allows IMAP access for fetching emails.

If using Gmail, enable App Passwords instead of your normal password.

The project assumes the template file follows a specific format.