#!/usr/bin/env python3
"""
Zimbra PDF Retriever via IMAP
Requires: imap-tools (pip install imap-tools)
"""

import os
import re
import sys
from datetime import datetime, timedelta
from imap_tools import MailBox, AND

def clean_filename(filename):
    """Clean and normalize filename for saving to disk"""
    if filename:
        return re.sub(r'[\\/*?:"<>|]', "_", filename)
    return "unnamed_attachment.pdf"

def get_monthly_folder(email_date, base_output_dir):
    """Determine the monthly subfolder (e.g., 2025-05) based on email date"""
    year = email_date.year
    month = email_date.month
    folder_name = f"{year}-{month:02d}"  # Format as YYYY-MM
    monthly_folder = os.path.join(base_output_dir, folder_name)
    if not os.path.exists(monthly_folder):
        os.makedirs(monthly_folder)
    return monthly_folder

def download_pdf_attachments(
    email,
    password,
    server="imap.tiongnam.com.my",
    output_dir="pdf_attachments",
    days_limit=None,
    folder_name="INBOX",
    only_unread=False,
    search_term=None,
    debug=False
):
    """Download PDF attachments via IMAP into monthly folders"""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created directory: {output_dir}")

    downloaded_files = []

    try:
        # Connect to IMAP server
        if debug:
            print(f"Connecting to {server}:993 with email {email}")
        mailbox = MailBox(server, port=993)

        # Login
        mailbox.login(email, password, initial_folder=folder_name)
        if debug:
            print(f"Successfully logged in to folder: {folder_name}")

        # Build search criteria
        criteria = []
        if only_unread:
            criteria.append(AND(seen=False))
        if search_term:
            criteria.append(AND(subject=search_term))
        if days_limit:
            date = datetime.now() - timedelta(days=days_limit)
            criteria.append(AND(date_gte=date))

        # Combine criteria (excluding attachment filter due to version issue)
        if criteria:
            search_criteria = AND(*criteria)
        else:
            search_criteria = None

        if debug:
            print(f"Search criteria: {search_criteria}")

        # Fetch messages (all messages if no criteria)
        messages = mailbox.fetch(search_criteria) if search_criteria else mailbox.fetch()
        total_messages = sum(1 for _ in messages)  # Count messages
        messages = mailbox.fetch(search_criteria) if search_criteria else mailbox.fetch()  # Reset iterator
        print(f"Found {total_messages} messages")

        # Process each message
        for i, msg in enumerate(messages, 1):
            subject = clean_filename(msg.subject)
            print(f"Processing message {i}/{total_messages}: {subject}")

            # Determine the monthly subfolder based on email date
            email_date = msg.date
            monthly_folder = get_monthly_folder(email_date, output_dir)
            if debug:
                print(f"  Saving to monthly folder: {monthly_folder}")

            # Check attachments manually
            if msg.attachments:
                for att in msg.attachments:
                    if att.filename.lower().endswith(".pdf"):
                        filename = clean_filename(att.filename)
                        print(f"  Downloading: {filename}")

                        # Ensure unique filename within the monthly folder
                        base_name, ext = os.path.splitext(filename)
                        final_path = os.path.join(monthly_folder, filename)
                        counter = 1
                        while os.path.exists(final_path):
                            new_filename = f"{base_name}_{counter}{ext}"
                            final_path = os.path.join(monthly_folder, new_filename)
                            counter += 1

                        # Save attachment
                        with open(final_path, "wb") as f:
                            f.write(att.payload)
                        downloaded_files.append(final_path)
                        print(f"  Saved: {final_path}")

        mailbox.logout()
        if debug:
            print("Logged out of IMAP server")

    except Exception as e:
        print(f"Error: {e}")
        print("Ensure the IMAP server, email, and password are correct.")

    print(f"Downloaded {len(downloaded_files)} PDF attachments")
    return downloaded_files

def main():
    """Interactive main function"""
    print("===== Zimbra IMAP PDF Attachment Retriever =====")

    debug_mode = "--debug" in sys.argv
    if debug_mode:
        print("Debug mode enabled")

    email = input("Enter your email address: ")
    password = input("Enter your app-specific password: ")
    days_filter = input("Only retrieve emails from the last N days? (leave blank for all): ")
    days_limit = int(days_filter) if days_filter and days_filter.isdigit() else None
    folder = input("Which folder to search? (default: INBOX): ") or "INBOX"
    filter_option = input("Filter options: (1: All emails, 2: Unread only, 3: Search by term): ")

    only_unread = filter_option == "2"
    search_term = input("Enter search term: ") if filter_option == "3" else None
    output_dir = input("Output directory (default: pdf_attachments): ") or "pdf_attachments"

    downloaded_files = download_pdf_attachments(
        email=email,
        password=password,
        output_dir=output_dir,
        days_limit=days_limit,
        folder_name=folder,
        only_unread=only_unread,
        search_term=search_term,
        debug=debug_mode
    )

    if downloaded_files:
        print(f"\nSuccessfully downloaded {len(downloaded_files)} PDF attachments to '{output_dir}' directory.")
    else:
        print("\nNo PDF attachments were found or downloaded.")

if __name__ == "__main__":
    main()