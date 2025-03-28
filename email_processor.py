import asyncio
from imapclient import IMAPClient
import email
import os
import re
from config import EMAIL_FOLDER, EMAIL_SUBJECT, ATTACHMENT_FOLDER

async def fetch_email(email_user, email_pass):
    """Fetch unread emails with specified subject and save the attachment."""
    try:
        def sync_fetch():
            with IMAPClient("imap.gmail.com", ssl=True) as mail:
                mail.login(email_user, email_pass)
                mail.select_folder(EMAIL_FOLDER)
                messages = mail.search(["UNSEEN", "SUBJECT", EMAIL_SUBJECT])
                if not messages:
                    return None, None
                
                latest_email_id = messages[-1]
                raw_email = mail.fetch(latest_email_id, ["RFC822"])
                msg = email.message_from_bytes(raw_email[latest_email_id][b"RFC822"])
                sender_email = email.utils.parseaddr(msg["From"])[1]

                for part in msg.walk():
                    if part.get_content_disposition() == "attachment" and re.match(r"Statement_.*\.(xlsm|xlsx)", part.get_filename()):
                        file_path = os.path.join(ATTACHMENT_FOLDER, part.get_filename())
                        with open(file_path, "wb") as f:
                            f.write(part.get_payload(decode=True))
                        return file_path, sender_email
                return None, None
        
        return await asyncio.to_thread(sync_fetch)
    except Exception as e:
        print(f"Error fetching email: {e}")
        return None, None
