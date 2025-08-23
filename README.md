This repository contains a Python script that automates the extraction of emails from a Gmail account using the IMAP protocol. The script allows users to filter emails by date and subject, decode MIME headers, and print email summaries. 

## Features
- Fetch emails from a Gmail account using IMAP.
- Filter emails by date and subject.
- Decode MIME headers for email subjects and senders.
- Print email summaries including subject, sender, and arrival time.

## Requirements
- Python 3.x
Add a cred.yaml file in the same directory with the following content:
```yaml
user: your_email@gmail.com
password: your_password
```
## Usage
1. Everything is set up, run the script:
   ```bash
   python email_parser.py
   ```
2. Enter the date in the format `YYYY-MM-DD` when prompted.
3. Enter the subject keyword to filter emails.

## TODO
- Implement a function to save the extracted emails to a file. Right now, the script only prints the email summaries to the console.
- Need to create a database with tables to store the email data for better management and retrieval.
