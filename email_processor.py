import asyncio
from imapclient import IMAPClient
import email
import os
import re
from config import EMAIL_FOLDER, EMAIL_SUBJECT, ATTACHMENT_FOLDER



async def fetch_email(email_user, email_pass):
    """Fetch unread emails from IMAP server and save the attachment properly."""
    try:
        print("[INFO] Connecting to email server...")



        def sync_fetch():
            with IMAPClient("imap.gmail.com", ssl=True) as mail:
                mail.login(email_user, email_pass)
                mail.select_folder(EMAIL_FOLDER)

                
                messages = mail.search(["UNSEEN", "SUBJECT", EMAIL_SUBJECT])
                if not messages:
                    print("[INFO] No new unread emails found.")
                    return None, None
                


                latest_email_id = messages[-1]
                raw_email = mail.fetch(latest_email_id, ["RFC822"])
                msg = email.message_from_bytes(raw_email[latest_email_id][b"RFC822"])
                sender_email = email.utils.parseaddr(msg["From"])[1]



                print(f"[INFO] Email received from: {msg['From']}")
                print(f"[INFO] Sender email saved: {sender_email}")



                # Process attachments
                for part in msg.walk():
                    if part.get_content_disposition() == "attachment":
                        filename = part.get_filename()
                        print(f"[DEBUG] Found attachment: {filename}")


                        
                        if re.match(r"Statement_.*\.(xlsm|xlsx)", filename):
                            file_path = os.path.join(ATTACHMENT_FOLDER, filename)
                            print(f"[DEBUG] Saving attachment to: {file_path}")


                            # Save the attachment to the folder
                            with open(file_path, "wb") as f:
                                f.write(part.get_payload(decode=True))

                            # Confirm the file was saved
                            if os.path.exists(file_path):
                                print(f"[✅ SUCCESS] File saved: {file_path}")
                                return file_path, sender_email
                            else:
                                print("❌ [ERROR] File save failed!")



                print("❌ [ERROR] No valid attachment found.")
                
                return None, None

        return await asyncio.to_thread(sync_fetch)

    except Exception as e:
        print(f"❌ [ERROR] {e}")
        return None, None
