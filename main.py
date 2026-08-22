import pandas as pd
import time
import random
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

EXCEL_FILE = "path of the file"
SENDER_EMAIL = "email of sender"
SENDER_PASSWORD = "app password"  # app password of gmail
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# load the excel file
df = pd.read_excel(EXCEL_FILE)

# ensure the status and date columns exist if not create
if "Status" not in df.columns:
    df["Status"] = ""
if "Date" not in df.columns:
    df["Date"] = ""

df["Status"] = df["Status"].astype(object)
df["Date"] = df["Date"].astype(object)

# counter to control the delay time
emails_sent_this_session = 0

try:
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        print("Connection to SMTP server is successful.\n")

        for index, row in df.iterrows():
            # only sent emails without SENT status
            status = str(row.get("Status", "")).strip().upper()
            if status == "SENT":
                continue

            organization_name = str(row.get("Organization", "")).strip()
            org_type = str(row.get("Type", "")).strip()
            location = str(row.get("Head Office Location", "")).strip()
            website = str(row.get("Official Website", "")).strip()
            recipient_mail = str(row.get("Official Email", "")).strip()

            # choose random subject to improve uniquity
            SUBJECT_TEMPLATES = [
                f"Partnership with GDGoC Bilkent | {organization_name}",
                f"GDGoC Bilkent & {organization_name} — Collaboration Opportunity",
                f"Connecting with {organization_name} ({location})",
            ]
            selected_subject = random.choice(SUBJECT_TEMPLATES)


            # check if file or a valid mail exists
            if not recipient_mail or "@" not in recipient_mail or str(recipient_mail).lower() == "nan":
                # mark invalid emails
                print(f"[{index + 1}/{len(df)}] [SKIPPED] Invalid email: '{recipient_mail}'")
                df.at[index, "Status"] = "INVALID"
                df.to_excel(EXCEL_FILE, index=False)
                continue

            # create the message
            body = (
                f"Hi {organization_name} Team,\n\n"
                f"I hope you're having a great week in {location}.\n\n"
                f"I'm reaching out from Google Developer Groups on Campus (GDGoC) at Bilkent University. "
                f"We have been following {organization_name}'s impactful work as a leading {org_type.lower()} "
                f"via {website} and would love to explore a sponsorship partnership for our upcoming hackathons and technical workshops.\n\n"
                f"Best regards,\n"
                f"GDGoC Bilkent"
            )

            msg = MIMEMultipart()
            msg["From"] = SENDER_EMAIL
            msg["To"] = recipient_mail
            msg["Subject"] = selected_subject
            msg.attach(MIMEText(body, "plain"))

            # send mail via checking errors for per recipient
            try:
                server.sendmail(SENDER_EMAIL, recipient_mail, msg.as_string())
                print(f"[{index + 1}/{len(df)}] [SENT] {organization_name} → {recipient_mail}")

                # update Status and Date Columns, and the counter of the session
                df.at[index, "Status"] = "SENT"
                df.at[index, "Date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                emails_sent_this_session += 1
            except smtplib.SMTPException as e:
                print(f"[{index+1}/{len(df)}] [FAILED] Could not send to {recipient_mail}: {e}")

                # Update Status on failure
                df.at[index, "Status"] = "NOT SENT"
                server.rset() # reset the connection

            # save the excel file
            df.to_excel(EXCEL_FILE, index=False)

            sleep_time = random.uniform(8.0, 15.0)
            time.sleep(sleep_time)  # wait for a random second to send the next mail

            # wait 1 minute every 60 seconds
            if emails_sent_this_session > 0 and emails_sent_this_session % 20 == 0:
                print(f"\n--- Batch limit reached ({emails_sent_this_session} sent). Pausing for 60 seconds... ---\n")
                time.sleep(60)
# handle mid force stop properly so it will safely continue in the next session
except KeyboardInterrupt:
    print("\n\n[USER STOPPED] Run terminated manually.")
    print(
        f"Successfully saved all {emails_sent_this_session} emails sent during this session."
    )
    print("You can resume anytime—already SENT rows will be skipped.")
except smtplib.SMTPAuthenticationError: #wrong log in credentials
    print("\n[AUTH ERROR] Could not log in. Verify your 2FA and Gmail App Password.")
except Exception as e:
    print(f"\n[CRITICAL ERROR] Connection failed: {e}")

print("\nProcess finished.")