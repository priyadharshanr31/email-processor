import smtplib
import os
import mimetypes
from email.message import EmailMessage

def send_email_with_attachment(file_path, receiver_email, email_user, email_pass):
    """Send the processed file to the original sender via email."""
    try:
        msg = EmailMessage()
        msg["Subject"] = "Processed Statement File"
        msg["From"] = email_user
        msg["To"] = receiver_email
        msg.set_content("Hello,\n\nPlease find the attached processed statement file.\n\nRegards,\nAutomated System")

        # Attach the file
        with open(file_path, "rb") as f:
            file_data = f.read()
            file_type, _ = mimetypes.guess_type(file_path)
            msg.add_attachment(file_data, maintype="application", subtype=file_type, filename=os.path.basename(file_path))

        # Send Email
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(email_user, email_pass)
            server.send_message(msg)

        print(f"✅ [SUCCESS] Email sent to {receiver_email}")
        return True
    except Exception as e:
        print(f"❌ [ERROR] Failed to send email: {e}")
        return False
