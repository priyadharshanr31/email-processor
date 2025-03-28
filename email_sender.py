import smtplib
import asyncio
from email.message import EmailMessage
from config import SMTP_SERVER, SMTP_PORT

async def send_email_async(email_user, email_pass, recipient_email, file_path):
    """Asynchronously sends an email with the processed file attached."""
    try:
        msg = EmailMessage()
        msg["Subject"] = "Processed Statement"
        msg["From"] = email_user
        msg["To"] = recipient_email
        msg.set_content("Please find the attached processed statement.")

        with open(file_path, "rb") as f:
            msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename="Processed_Statement.xlsm")
        
        def send_email():
            with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
                server.login(email_user, email_pass)
                server.send_message(msg)
        
        await asyncio.to_thread(send_email)
        return True
    except Exception as e:
        print(f"Email sending failed: {e}")
        return False
