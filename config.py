import os

# Root directory of the project
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# Email Configuration
EMAIL_SERVER = "imap.gmail.com"  
EMAIL_FOLDER = "INBOX"  
EMAIL_SUBJECT = "Statements"  

# File Paths
ATTACHMENT_FOLDER = os.path.join(BASE_DIR, "attachments")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")
TEMPLATE_FILE = os.path.join(BASE_DIR, "template.xlsx")



# Data Processing Settings
FILTER_VALUES = ["Y", "YES", "Yes", "yes", "y", "Done", "done"]

# Ensure required folders exist
os.makedirs(ATTACHMENT_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
