from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import random
import smtplib
import time
import pandas as pd

EXCEL_FILE = "path of the file"
SENDER_EMAIL = "email of sender"
SENDER_PASSWORD = "app password"  # App password for Gmail or standard password for Bilkent
REPLY_TO_EMAIL = "reply to email"  # Club email where replies should land

# SMTP Server Options (toggle as needed)
SMTP_SERVER = "smtp.gmail.com"
# SMTP_SERVER = "asmtp.bilkent.edu.tr"
SMTP_PORT = 587

DAILY_SESSION_LIMIT = 25  # Recommended cap per day


def safe_save_excel(dataframe: pd.DataFrame, file_path: str):
    # saves DataFrame to an atomic temp file first, preventing file corruption on abrupt Ctrl+C
    temp_path = f"{file_path}.tmp"
    try:
        dataframe.to_excel(temp_path, index=False)
        os.replace(temp_path, file_path)
    except Exception as save_err:
        print(f"[SAVE ERROR] Could not save directly: {save_err}")
        dataframe.to_excel(file_path, index=False)


# load the excel file
df = pd.read_excel(EXCEL_FILE)

# ensure status and date columns exist
if "Status" not in df.columns:
    df["Status"] = ""
if "Date" not in df.columns:
    df["Date"] = ""

df["Status"] = df["Status"].astype(object)
df["Date"] = df["Date"].astype(object)

emails_sent_this_session = 0

try:
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        print("Connection to SMTP server is successful.\n")

        for index, row in df.iterrows():
            if emails_sent_this_session >= DAILY_SESSION_LIMIT:
                print(
                    f"\n[SESSION COMPLETE] Reached daily limit of"
                    f" {DAILY_SESSION_LIMIT} emails."
                )
                break

            status = str(row.get("Status", "")).strip().upper()
            if status == "SENT":
                continue

            organization_name = str(row.get("Organization", "")).strip()
            org_type = str(row.get("Type", "")).strip()
            location = str(row.get("Head Office Location", "")).strip()
            website = str(row.get("Official Website", "")).strip()
            recipient_mail = str(row.get("Official Email", "")).strip()

            # basic email validation
            if (
                not recipient_mail
                or "@" not in recipient_mail
                or str(recipient_mail).lower() == "nan"
            ):
                print(
                    f"[{index + 1}/{len(df)}] [SKIPPED] Invalid email: '{recipient_mail}'"
                )
                df.at[index, "Status"] = "INVALID"
                safe_save_excel(df, EXCEL_FILE)
                continue

            # subject variations
            SUBJECT_TEMPLATES = [
                f"Partnership with GDGoC Bilkent | {organization_name}",
                (
                    f"GDGoC Bilkent & {organization_name} — Collaboration"
                    " Opportunity"
                ),
                f"Connecting with {organization_name} ({location})",
            ]
            selected_subject = random.choice(SUBJECT_TEMPLATES)

            # email body
            body = (
                f"Hi {organization_name} Team,\n\n"
                f"I hope you're having a great week in {location}.\n\n"
                "I'm reaching out from Google Developer Groups on Campus"
                " (GDGoC) at Bilkent University. We have been following"
                f" {organization_name}'s impactful work as a leading"
                f" {org_type.lower()} via {website} and would love to explore a"
                " sponsorship partnership for our upcoming hackathons and"
                " technical workshops.\n\n"
                "Best regards,\n"
                "GDGoC Bilkent"
            )

            # construct message
            msg = MIMEMultipart()
            msg["From"] = SENDER_EMAIL
            msg["To"] = recipient_mail
            msg["Reply-To"] = REPLY_TO_EMAIL
            msg["Subject"] = selected_subject
            msg.attach(MIMEText(body, "plain"))

            # send the message
            try:
                server.sendmail(SENDER_EMAIL, recipient_mail, msg.as_string())
                print(
                    f"[{index + 1}/{len(df)}] [SENT] {organization_name} →"
                    f" {recipient_mail}"
                )
                df.at[index, "Status"] = "SENT"
                df.at[index, "Date"] = datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
                emails_sent_this_session += 1
            except smtplib.SMTPException as e:
                print(
                    f"[{index + 1}/{len(df)}] [FAILED] Could not send to"
                    f" {recipient_mail}: {e}"
                )
                df.at[index, "Status"] = "NOT SENT"
                server.rset()

            # safe save to disk
            safe_save_excel(df, EXCEL_FILE)

            # pacing delay between single emails
            if emails_sent_this_session < DAILY_SESSION_LIMIT:
                sleep_time = random.uniform(25.0, 60.0)
                print(
                    f"Waiting {int(sleep_time)}s... (Progress:"
                    f" {emails_sent_this_session}/{DAILY_SESSION_LIMIT})"
                )
                time.sleep(sleep_time)

except KeyboardInterrupt:
    print("\n\n[USER STOPPED] Execution halted safely.")
    print(
        f"Saved all progress. Sent {emails_sent_this_session} emails in this"
        " session."
    )
    print("Already SENT rows remain preserved for next time.")
except smtplib.SMTPAuthenticationError:
    print("\n[AUTH ERROR] Login failed. Check your email credentials.")
except Exception as e:
    print(f"\n[CRITICAL ERROR] Connection error: {e}")

print("\nProcess finished.")