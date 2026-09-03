from datetime import datetime
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import html
import os
import random
import smtplib
import time
import pandas as pd

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = "sample.xlsx"
SENDER_EMAIL = "email of sender"
SENDER_PASSWORD = "password"  # App password for Gmail or standard password for Bilkent
REPLY_TO_EMAIL = "reply to email"  # Club email where replies should land
SIGNATURE_LOGO = os.path.join(SCRIPT_DIR, "assets", "gdg-logo.png")

# SMTP Server Options (toggle as needed)
# SMTP_SERVER = "smtp.gmail.com"
SMTP_SERVER = "asmtp.bilkent.edu.tr"
SMTP_PORT = 587

DAILY_SESSION_LIMIT = 25  # Recommended cap per day

PLACEHOLDER_HINT = (
    "Available placeholders: {organization_name}, {org_type}, {location},"
    " {website}"
)

# Fixed signature details matching the template
SIGNATURE_DATA = {
    "name": "Bayim Abbaszade",
    "role": "Partnership Team Lead",
    "phone": "+ 90 501 554 43 21",
    "email": "gdgoc.bilkent@gmail.com",
    "chapter": "Bilkent University",
}


def apply_placeholders(template: str, row: pd.Series) -> str:
    values = {
        "organization_name": str(row.get("Organization", "")).strip(),
        "org_type": str(row.get("Type", "")).strip(),
        "location": str(row.get("Head Office Location", "")).strip(),
        "website": str(row.get("Official Website", "")).strip(),
    }
    result = template
    for key, value in values.items():
        result = result.replace("{" + key + "}", value)
    return result


def is_valid_email(email: str) -> bool:
    return bool(email and "@" in email and str(email).lower() != "nan")


def prompt_send_mode() -> str:
    print("Send mode:")
    print("  1) Bulk send from Excel file")
    print("  2) Send to individual email")
    while True:
        choice = input("Choose mode [1/2]: ").strip()
        if choice in ("1", "2"):
            return choice
        print("Please enter 1 or 2.")


def prompt_individual_email() -> str:
    while True:
        email = input("Enter recipient email: ").strip()
        if is_valid_email(email):
            return email
        print("Invalid email. Please try again.")


def prompt_email_content(*, use_placeholders: bool = True) -> tuple[str, str]:
    if use_placeholders:
        print(PLACEHOLDER_HINT + "\n")

    subject = input("Enter email subject: ").strip()
    while not subject:
        subject = input("Subject cannot be empty. Enter email subject: ").strip()

    print("\nEnter or paste your email body below.")
    print("Type 'DONE' on a new line and press Enter when finished:\n")

    body_lines = []
    while True:
        try:
            line = input()
        except EOFError:
            break
        if line.strip().upper() == "DONE":
            break
        body_lines.append(line)

    body = "\n".join(body_lines).strip()
    while not body:
        print("\n[!] Body cannot be empty. Enter email body (type 'DONE' on a new line when finished):")
        body_lines = []
        while True:
            try:
                line = input()
            except EOFError:
                break
            if line.strip().upper() == "DONE":
                break
            body_lines.append(line)
        body = "\n".join(body_lines).strip()

    return subject, body


def confirm_dispatch(subject: str, target_summary: str) -> bool:
    print("\n" + "=" * 50)
    print("FINAL DISPATCH CONFIRMATION")
    print("=" * 50)
    print(f"Target: {target_summary}")
    print(f"Subject: {subject}")
    print("Signature: Attached automatically at bottom")
    print("=" * 50)

    while True:
        choice = input("Send out emails now? (y/n): ").strip().lower()
        if choice in ("yes", "y"):
            return True
        if choice in ("no", "n"):
            print("\n[ABORTED] Operation canceled. No emails were sent.")
            return False
        print("Invalid choice. Please enter 'y' or 'n'.")


def build_plain_signature() -> str:
    return (
        f"\n\n--\n"
        f"{SIGNATURE_DATA['name']}\n"
        f"{SIGNATURE_DATA['role']}\n"
        f"{SIGNATURE_DATA['phone']}\n"
        f"{SIGNATURE_DATA['email']}\n\n"
        f"Google Developer Group\n"
        f"{SIGNATURE_DATA['chapter']}"
    )


def build_html_signature() -> str:
    return f"""
<table cellpadding="0" cellspacing="0" style="margin-top:28px;font-family:Arial,sans-serif;color:#202124;">
  <tr>
    <td style="padding:0;">
      <div style="font-size:15px;font-weight:700;color:#202124;">{html.escape(SIGNATURE_DATA["name"])}</div>
      <div style="font-size:13px;color:#5f6368;margin-top:3px;">{html.escape(SIGNATURE_DATA["role"])}</div>
      <div style="font-size:13px;margin-top:6px;">
        <span style="color:#1a73e8;font-weight:600;">{html.escape(SIGNATURE_DATA["phone"])}</span>
      </div>
      <div style="font-size:13px;margin-top:2px;">
        <a href="mailto:{html.escape(SIGNATURE_DATA['email'])}" style="color:#1a73e8;text-decoration:none;font-weight:600;">
          {html.escape(SIGNATURE_DATA["email"])}
        </a>
      </div>
      <table cellpadding="0" cellspacing="0" style="margin-top:16px;">
        <tr>
          <td style="padding:4px;background-color:#ffffff;border-radius:4px;display:inline-block;vertical-align:middle;">
            <img src="cid:gdg_logo" alt="Google Developer Group" width="260"
                 style="display:block;max-width:260px;height:auto;" />
          </td>
        </tr>
        <tr>
          <td style="padding-top:4px;font-size:13px;font-weight:600;color:#1a73e8;">
            {html.escape(SIGNATURE_DATA["chapter"])}
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
""".strip()


from email.utils import formataddr, formatdate, make_msgid


def build_email_message(recipient_mail: str, subject: str, body: str) -> MIMEMultipart:
    plain_body = f"{body}{build_plain_signature()}"
    html_body = (
        f"<div style=\"font-family:Arial,sans-serif;font-size:14px;color:#202124;"
        f"white-space:pre-wrap;line-height:1.5;\">{html.escape(body)}</div>"
        f"{build_html_signature()}"
    )

    msg = MIMEMultipart("related")

    msg["From"] = formataddr(("Bayim Abbaszade", SENDER_EMAIL))
    msg["To"] = recipient_mail
    msg["Reply-To"] = REPLY_TO_EMAIL
    msg["Subject"] = subject

    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain="ug.bilkent.edu.tr")

    alternative = MIMEMultipart("alternative")
    alternative.attach(MIMEText(plain_body, "plain", "utf-8"))
    alternative.attach(MIMEText(html_body, "html", "utf-8"))
    msg.attach(alternative)

    with open(SIGNATURE_LOGO, "rb") as logo_file:
        logo = MIMEImage(logo_file.read(), _subtype="png")
        logo.add_header("Content-ID", "<gdg_logo>")
        logo.add_header("Content-Disposition", "inline", filename="gdg-logo.png")
        msg.attach(logo)

    return msg

def safe_save_excel(dataframe: pd.DataFrame, file_path: str):
    root, ext = os.path.splitext(file_path)
    temp_path = f"{root}.tmp{ext or '.xlsx'}"
    try:
        dataframe.to_excel(temp_path, index=False)
        os.replace(temp_path, file_path)
    except Exception as save_err:
        print(f"[SAVE ERROR] Could not save directly: {save_err}")
        if os.path.exists(temp_path):
            os.remove(temp_path)
        dataframe.to_excel(file_path, index=False)


def ensure_signature_logo() -> None:
    if not os.path.isfile(SIGNATURE_LOGO):
        raise FileNotFoundError(
            f"Signature logo not found at '{SIGNATURE_LOGO}'. "
            "Ensure assets/gdg-logo.png exists."
        )


def send_email(server: smtplib.SMTP, recipient_mail: str, msg: MIMEMultipart) -> None:
    server.sendmail(SENDER_EMAIL, recipient_mail, msg.as_string())


def run_individual() -> None:
    ensure_signature_logo()

    recipient_mail = prompt_individual_email()
    subject, body = prompt_email_content(use_placeholders=False)

    if not confirm_dispatch(subject, f"Individual -> {recipient_mail}"):
        return

    print("\nSending email...\n")

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            print("Connection to SMTP server is successful.\n")

            msg = build_email_message(recipient_mail, subject, body)
            send_email(server, recipient_mail, msg)
            print(f"[SENT] → {recipient_mail}")
    except smtplib.SMTPAuthenticationError:
        print("\n[AUTH ERROR] Login failed. Check your email credentials.")
    except smtplib.SMTPException as e:
        print(f"\n[FAILED] Could not send to {recipient_mail}: {e}")
    except Exception as e:
        print(f"\n[CRITICAL ERROR] Connection error: {e}")

    print("\nProcess finished.")


def run_bulk() -> None:
    df = pd.read_excel(EXCEL_FILE)

    if "Status" not in df.columns:
        df["Status"] = ""
    if "Date" not in df.columns:
        df["Date"] = ""

    df["Status"] = df["Status"].astype(object)
    df["Date"] = df["Date"].astype(object)

    pending_count = len(df[df["Status"].astype(str).str.strip().str.upper() != "SENT"])
    ensure_signature_logo()

    subject_template, body_template = prompt_email_content()

    target_info = (
        f"Bulk Send ({min(pending_count, DAILY_SESSION_LIMIT)} queue limit "
        f"out of {pending_count} pending entries in '{EXCEL_FILE}')"
    )
    if not confirm_dispatch(subject_template, target_info):
        return

    print("\nStarting email dispatch...\n")
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
                recipient_mail = str(row.get("Official Email", "")).strip()

                if not is_valid_email(recipient_mail):
                    print(
                        f"[{index + 1}/{len(df)}] [SKIPPED] Invalid email:"
                        f" '{recipient_mail}'"
                    )
                    df.at[index, "Status"] = "INVALID"
                    safe_save_excel(df, EXCEL_FILE)
                    continue

                selected_subject = apply_placeholders(subject_template, row)
                body = apply_placeholders(body_template, row)
                msg = build_email_message(recipient_mail, selected_subject, body)

                try:
                    send_email(server, recipient_mail, msg)
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

                safe_save_excel(df, EXCEL_FILE)

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


def run() -> None:
    if prompt_send_mode() == "2":
        run_individual()
    else:
        run_bulk()


if __name__ == "__main__":
    run()