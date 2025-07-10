import win32com.client as win32
from pathlib import Path

# ====== CONFIGURE THESE ======
TO_EMAIL = "send your email here" # Change to your email for testing
SUBJECT = "Test Email from Python (win32com)"
BODY = "This is a test email sent from Python using win32com and Outlook."
ATTACHMENT_PATH = r"D:\coldmail\vedant_singh resume.pdf"  # e.g., r"D:\coldmail\vedant_singh resume.pdf" or None
# =============================

def send_test_email():
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = TO_EMAIL
        mail.Subject = SUBJECT
        mail.Body = BODY

        if ATTACHMENT_PATH and Path(ATTACHMENT_PATH).exists():
            mail.Attachments.Add(str(Path(ATTACHMENT_PATH).resolve()))
            print(f"Attached: {ATTACHMENT_PATH}")

        mail.Send()
        print(f"✅ Test email sent to {TO_EMAIL} via Outlook!")
    except Exception as e:
        print(f"❌ Failed to send test email: {e}")

if __name__ == "__main__":
    send_test_email() 