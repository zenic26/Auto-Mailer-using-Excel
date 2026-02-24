import pandas as pd
import smtplib
from email.message import EmailMessage
import time
import re
import socket

YOUR_EMAIL = "your@gmail.com"
YOUR_NAME = "yout name"

EXCEL_FILE = r"C:\Users\Admin\Desktop\automailer iwth excel\emails.xlsx"

EMAIL_DELAY = 2
LOG_FILE = "email_log.txt"

APP_PASSWORD = "password" 



def get_email_body(company_name):
    return f"""Hello {company_name},
    thank you for providing me with your email addresses! this mail is sent to you for testing purpose. I hope you receive it without any issues.
    Best regards,
    {YOUR_NAME}
l

 
"""

def get_email_html_body(company_name):
    return f"""\
<html>
  <body>
    <p>Hello {company_name},
    thank you for providing me with your email addresses! this mail is sent to you for testing purpose. I hope you receive it without any issues.
    Best regards,
    {YOUR_NAME}</p>
  </body>
</html>
"""

def get_email_subject(company_name):
    return f"Collaboration Opportunity with {company_name}"


def is_valid_email(email):
    regex = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    return re.match(regex, email) is not None

def send_emails():
    print("=" * 50)
    print("AUTO EMAIL SENDER")
    print("=" * 50)

    if not APP_PASSWORD:
        print("ERROR: Gmail App Password not found in environment variables!")
        print("Set it using your OS environment variable: GMAIL_APP_PASSWORD")
        return
    
    print(f"\nReading Excel file: {EXCEL_FILE}")
    try:
        data = pd.read_excel(EXCEL_FILE)
        if not {'Company', 'Email'}.issubset(data.columns):
            print("ERROR: Excel file must have 'Company' and 'Email' columns.")
            return
        print(f"   Found {len(data)} recipients")
    except FileNotFoundError:
        print(f"ERROR: File '{EXCEL_FILE}' not found!")
        return
    except Exception as e:
        print(f"ERROR reading Excel file: {e}")
        return

    print("\nConnecting to Gmail SMTP server...")
    try:
        server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        server.login(YOUR_EMAIL, APP_PASSWORD)
        print("   Connected successfully!")
    except Exception as e:
        print(f"ERROR connecting to server: {e}")
        return
    print("\nSending emails...")
    success_count = 0
    fail_count = 0

    for index, row in data.iterrows():
        company = str(row["Company"]).strip()
        receiver_email = str(row["Email"]).strip()

        if not is_valid_email(receiver_email):
            print(f"   Invalid email format: {receiver_email}")
            fail_count += 1
            with open(LOG_FILE, "a") as log_file:
                log_file.write(f"{company} ({receiver_email}) - FAILED: Invalid email format\n")
            continue

        try:
            msg = EmailMessage()
            msg["From"] = YOUR_EMAIL
            msg["To"] = receiver_email
            msg["Subject"] = get_email_subject(company)
            msg.set_content(get_email_body(company))
            msg.add_alternative(get_email_html_body(company), subtype='html')

            server.send_message(msg)
            print(f"   Email sent to {company} ({receiver_email})")
            success_count += 1

            with open(LOG_FILE, "a") as log_file:
                log_file.write(f"{company} ({receiver_email}) - SUCCESS\n")

            if index < len(data) - 1:
                time.sleep(EMAIL_DELAY)

        except (smtplib.SMTPException, socket.timeout) as e:
            print(f"   Failed to send to {company} ({receiver_email}): {e}")
            fail_count += 1
            with open(LOG_FILE, "a") as log_file:
                log_file.write(f"{company} ({receiver_email}) - FAILED: {e}\n")

    server.quit()
    print("\nDisconnected from server")

    print("\n" + "=" * 50)
    print("SUMMARY")
    print("=" * 50)
    print(f"   Successful: {success_count}")
    print(f"   Failed: {fail_count}")
    print(f"   Detailed log saved in: {LOG_FILE}")

if __name__ == "__main__":
    send_emails()