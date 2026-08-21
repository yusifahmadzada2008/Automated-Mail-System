import pandas as pd
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

EXCEL_FILE = "temp.xlsx"
SENDER_EMAIL = "YOUR_EMAIL_HERE"
SENDER_PASSWORD = "YOUR_APP_PASSWORD_HERE"  # app password of gmail
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
DELAY_BETWEEN_EMAILS = 1.5  # delay to prevent block by gmail for too frequently sent mails

df = pd.read_excel(EXCEL_FILE)

try:
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        print("Connection to SMTP server is successful.\n")

        for index, row in df.iterrows():
            recipient_name = str(row.get("Name", "")).strip()
            recipient_mail = str(row.get("Mail", "")).strip()

            # check if file or a valid mail exists
            if not recipient_mail or "@" not in recipient_mail:
                print(f"[{index+1}/{len(df)}] [SKIPPED] Invalid email: '{recipient_mail}'")
                continue

            # create the message
            msg = MIMEMultipart()
            msg["From"] = SENDER_EMAIL
            msg["To"] = recipient_mail
            msg["Subject"] = f"Your Subject Here for {recipient_name}"

            body = f"""Dear {recipient_name},

This is the draft message.

Best regards,
GDGoC Bilkent
"""
            msg.attach(MIMEText(body, "plain"))

            # send mail via checking errors for per recipient
            try:
                server.sendmail(SENDER_EMAIL, recipient_mail, msg.as_string())
                print(f"[{index+1}/{len(df)}] [SENT] {recipient_name} → {recipient_mail}")
            except smtplib.SMTPException as e:
                print(f"[{index+1}/{len(df)}] [FAILED] Could not send to {recipient_mail}: {e}")

            time.sleep(DELAY_BETWEEN_EMAILS)

except smtplib.SMTPAuthenticationError: #wrong log in credentials
    print("\n[AUTH ERROR] Could not log in. Verify your 2FA and Gmail App Password.")
except Exception as e:
    print(f"\n[CRITICAL ERROR] Connection failed: {e}")

print("\nProcess finished.")