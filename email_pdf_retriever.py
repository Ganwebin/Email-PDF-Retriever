#!/usr/bin/env python3

"""
default is using gmail, other email need provide imap and port

if using gmail it will prevent you using original password ,need to open two-step verificatio and create a app password

if using custom search need to using IMAP queries

"""

import email
import imaplib
import os
import re
import getpass
from email.header import decode_header
from datetime import datetime, timedelta  

def clean_filename(filename):
    """Clean and normalize filename for saving to disk"""
    if filename:
        # Decode filename if needed
        if isinstance(filename, bytes):
            filename = filename.decode()
        # Replace invalid chars with underscore
        return re.sub(r'[\\/*?:"<>|]', "_", filename)
    return "unnamed_attachment.pdf"

def download_pdf_attachments(
    email_address,
    password,
    imap_server="imap.gmail.com",
    imap_port=993,
    search_criteria="ALL",
    output_dir="pdf_attachments",
    days_limit=None
):
    
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created directory: {output_dir}")
    
    # Connect to the IMAP server
    try:
        mail = imaplib.IMAP4_SSL(imap_server, imap_port)
        mail.login(email_address, password)
    except Exception as e:
        print(f"Error connecting to mail server: {e}")
        return []
    
    print(f"Successfully connected to {imap_server}")
    
    # Select the mailbox (inbox by default)
    mail.select("INBOX")
    
    # Adjust search criteria for date limit if specified
    if days_limit:
        date_criteria = (datetime.now() - timedelta(days=days_limit)).strftime("%d-%b-%Y")

        search_criteria = f'(SINCE {date_criteria}) {search_criteria}'
    
    # Search for emails
    status, messages = mail.search(None, search_criteria)
    
    if status != 'OK':
        print("No messages found!")
        mail.logout()
        return []
        
    # Parse the messages
    downloaded_files = []
    message_ids = messages[0].split()
    total_messages = len(message_ids)
    
    print(f"Found {total_messages} messages matching criteria")
    
    for i, message_id in enumerate(message_ids):
        try:
            print(f"Processing message {i+1}/{total_messages}...")
            
            # Fetch the message
            status, msg_data = mail.fetch(message_id, "(RFC822)")
            
            if status != 'OK':
                print(f"Error fetching message {message_id}")
                continue
                
            # Parse the raw email message
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)
            
            # Get email subject for better naming
            subject = msg["Subject"]
            if subject:
                # Decode subject if needed
                subject, encoding = decode_header(subject)[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding or 'utf-8', errors='replace')
                subject = clean_filename(subject)
            else:
                subject = "no_subject"
            
            # Process attachments
            if msg.is_multipart():
                for part in msg.walk():
                    content_disposition = str(part.get("Content-Disposition", ""))
                    
                    # Check if it's an attachment and a PDF
                    if "attachment" in content_disposition and part.get_content_type() == "application/pdf":
                        filename = part.get_filename()
                        if filename:
                            # Clean and create a unique filename
                            clean_name = clean_filename(filename)
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            unique_filename = f"{subject}_{timestamp}_{clean_name}"
                            
                            # Save the attachment
                            filepath = os.path.join(output_dir, unique_filename)
                            with open(filepath, "wb") as f:
                                f.write(part.get_payload(decode=True))
                            print(f"Downloaded: {filepath}")
                            downloaded_files.append(filepath)
        except Exception as e:
            print(f"Error processing message {message_id}: {e}")
    
    # Logout when done
    mail.logout()
    
    print(f"Downloaded {len(downloaded_files)} PDF attachments")
    return downloaded_files

def main():
    """Main function to run the script interactively"""
    print("===== Email PDF Attachment Retriever =====")
    
    email_address = input("Enter your email address: ")
    password = getpass.getpass("Enter your email password/app password: ")
    
    # Default to Gmail, but allow custom server
    use_gmail = input("Use Gmail server? (y/n, default: y): ").lower() != 'n'
    
    if use_gmail:
        imap_server = "imap.gmail.com"
        imap_port = 993
    else:
        imap_server = input("Enter IMAP server address: ")
        imap_port = int(input("Enter IMAP port (default: 993): ") or 993)
    
    # Allow user to filter by date
    days_filter = input("Only retrieve emails from the last N days? (leave blank for all): ")
    days_limit = int(days_filter) if days_filter and days_filter.isdigit() else None
    
    # Allow user to specify search criteria
    search_option = input("Search for specific emails? (1: All emails, 2: Unread only, 3: Custom search): ")
    
    if search_option == "2":
        search_criteria = "UNSEEN"
    elif search_option == "3":
        print("Enter custom search criteria (e.g., 'FROM someone@example.com', 'SUBJECT \"report\"')")
        search_criteria = input("Search criteria: ")
    else:
        search_criteria = "ALL"
    
    # Output directory
    output_dir = input("Output directory (default: pdf_attachments): ") or "pdf_attachments"
    
    # Execute download
    print("\nConnecting to email server and searching for PDFs...")
    downloaded_files = download_pdf_attachments(
        email_address=email_address,
        password=password,
        imap_server=imap_server,
        imap_port=imap_port,
        search_criteria=search_criteria,
        output_dir=output_dir,
        days_limit=days_limit
    )
    
    # Report results
    if downloaded_files:
        print(f"\nSuccessfully downloaded {len(downloaded_files)} PDF attachments to '{output_dir}' directory.")
    else:
        print("\nNo PDF attachments were found or downloaded.")

if __name__ == "__main__":
    main()