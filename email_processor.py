import imaplib
import email
import os
import re
from config import EMAIL_FOLDER, EMAIL_SUBJECT, ATTACHMENT_FOLDER

def check_email_and_download(email_user, email_pass):
    """Check email and download the latest statement attachment."""

    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(email_user, email_pass)
        mail.select(EMAIL_FOLDER)



        # Search for unread emails with the subject
        result, data = mail.search(None, f'(UNSEEN SUBJECT "{EMAIL_SUBJECT}")')

        
        
        if result != "OK" or not data[0]:
            return None, None

        latest_email_id = data[0].split()[-1]
        result, msg_data = mail.fetch(latest_email_id, "(RFC822)")

        if result != "OK":
            return None, None
        

        msg = email.message_from_bytes(msg_data[0][1])
        sender_email = email.utils.parseaddr(msg["From"])[1]



        # Process attachments
        for part in msg.walk():
            if part.get_content_disposition() == "attachment":
                filename = part.get_filename()
                if re.match(r"Statement_.*\.xlsx", filename):
                    file_path = os.path.join(ATTACHMENT_FOLDER, filename)
                    with open(file_path, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    return file_path, sender_email



        return None, None
    except Exception:
        return None, None
