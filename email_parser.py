import re
import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import logging
import yaml
import time
from email.policy import default as default_policy

def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)


def load_credentials(filepath):
    try:
        with open(filepath, 'r') as file:
            credentials = yaml.safe_load(file)
            user = credentials['user']
            password = credentials['password']
            return user, password
    except Exception as e:
        logging.error("Failed to load credentials: {}".format(e))
        raise

def connect_to_gmail_imap(user, password):
    imap_url = 'imap.gmail.com'
    # imap_url = "outlook.office365.com"
    try:
        mail = imaplib.IMAP4_SSL(imap_url)
        mail.login(user, password)
        mail.select('inbox')  # Connect to the inbox.
        return mail
    except Exception as e:
        logging.error("Connection failed: {}".format(e))
        raise


def get_email_count(mail):
    status, messages = mail.select("INBOX")
    messages = int(messages[0])
    return messages

def fetch_latest_N_emails(mail, N):
    messages = get_email_count(mail)
    for i in range(messages, messages-N, -1):
        res, msg_data = mail.fetch(str(i), "(RFC822)")
        _, time_data = mail.fetch(str(i), "(INTERNALDATE)")
        arrival_time = imaplib.Internaldate2tuple(time_data[0])
        for response in msg_data:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1], policy=default_policy)
                subject = decode_mime_header(msg.get("Subject"))
                sender = decode_mime_header(msg.get("From"))
                print_email_summary(subject, sender, arrival_time)
                process_email_parts(msg, subject)
                print("\n" + "="*50 + "\n")

def fetch_emails_by_date_and_subject(mail, date_str, keyword):
    messages = get_email_count(mail)
    for i in range(messages, 0, -1):
        _, time_data = mail.fetch(str(i), "(INTERNALDATE)")
        arrival_date = imaplib.Internaldate2tuple(time_data[0])
        if arrival_date is not None and isinstance(arrival_date, tuple) and len(arrival_date) == 9:
            email_date_str = time.strftime("%Y-%m-%d", arrival_date)
            if email_date_str == date_str:
                res, msg_data = mail.fetch(str(i), "(RFC822)")
                for response in msg_data:
                    if isinstance(response, tuple):
                        msg = email.message_from_bytes(response[1], policy=default_policy)
                        subject = decode_mime_header(msg.get("Subject"))
                        sender = decode_mime_header(msg.get("From"))
                        # Check if keyword is in subject (case-insensitive)
                        if keyword and keyword.lower() not in subject.lower():
                            continue
                        print_email_summary(subject, sender, arrival_date)
                        process_email_parts(msg, subject)
                        print("\n" + "="*50 + "\n")

def decode_mime_header(header):
    if not header:
        return ""
    decoded, encoding = decode_header(header)[0]
    if isinstance(decoded, bytes):
        return decoded.decode(encoding or "utf-8", errors="replace")
    return decoded

def print_email_summary(subject, sender, arrival_time):
    print("Subject:", subject)
    print("From:", sender)
    print("Arrival Time:", time.strftime("%a, %d %b %Y %H:%M:%S", arrival_time))

def process_email_parts(msg, subject):
    if msg.is_multipart():
        for part in msg.walk():
            handle_email_part(part, subject)
    else:
        handle_email_part(msg, subject)

def handle_email_part(part, subject):
    content_type = part.get_content_type()
    content_disposition = str(part.get("Content-Disposition"))
    try:
        body = part.get_payload(decode=True)
        if body:
            body = body.decode(part.get_content_charset() or "utf-8", errors="replace")
    except Exception:
        body = None
    if content_type == "text/plain" and "attachment" not in content_disposition:
        if body:
            print(body)
    elif "attachment" in content_disposition:
        filename = part.get_filename()
        if filename:
            folder_name = clean(subject)
            if not os.path.isdir(folder_name):
                os.mkdir(folder_name)
            filepath = os.path.join(folder_name, filename)
            with open(filepath, "wb") as f:
                f.write(part.get_payload(decode=True))

def get_user_preference():
    print("Choose your email extraction preference:")
    print("1. Enter the number of latest emails")
    print("2. Filter by date and subject keywords")
    while True:
        choice = input("Enter your choice (1-2): ")
        if choice in {'1', '2'}:
            return int(choice)
        else:
            print("Invalid choice. Please enter a number between 1 and 2.")
    
def main():
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    try:
        user, password = load_credentials('cred.yaml')
        mail = connect_to_gmail_imap(user, password)
        preference = get_user_preference()
        if preference == 1:
            N = int(input("Enter the number of latest emails to fetch: "))
            fetch_latest_N_emails(mail, N)
        elif preference == 2:
            date_str = input("Enter the date (YYYY-MM-DD): ")
            keyword = input("Enter the subject keyword (or leave empty for all emails): ")
            if keyword.strip() == "":
                keyword = None
            fetch_emails_by_date_and_subject(mail, date_str, keyword)
    except Exception as e:
        logging.error("An error occurred: {}".format(e))
    finally:
        mail.logout()

if __name__ == "__main__":
    main()