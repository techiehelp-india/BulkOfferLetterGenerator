#!/usr/bin/env python3
# Bulk Email Sender for Offer Letters
# Reads students.xlsx, attaches matching PDF, sends via Gmail SMTP with app password
# Usage: Edit sender_email/app_password, run `python bulk_email_sender.py`

import pandas as pd
import os
import getpass
from email_sender import EmailSender

def sanitize_filename(name):
    """Sanitize name for PDF filename matching generate_letters.py"""
    import re
    name = re.sub(r'[^\w\s-]', '', name).strip().replace(' ', '_')
    return name

def main():
    # Config
    excel_file = 'students.xlsx'
    offer_letters_dir = 'offer_letters'
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    subject = "Offer Letter from TechieHelp"
    
    # Input credentials securely
    sender_email = input("Enter sender Gmail: ").strip()
    sender_password = getpass.getpass("Enter Gmail App Password: ")
    
    if not sender_email or not sender_password:
        print("❌ Sender credentials required!")
        return
    
    # Initialize sender
    sender = EmailSender(smtp_server, smtp_port, sender_email, sender_password)
    
    # Read Excel
    try:
        df = pd.read_excel(excel_file)
        print(f"📊 Loaded {len(df)} students from {excel_file}")
    except Exception as e:
        print(f"❌ Error reading Excel: {e}")
        return
    
    # Validate columns
    required_cols = ['Name', 'Email']
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        print(f"❌ Missing columns: {missing}")
        return
    
    # Stats
    total = 0
    sent = 0
    failed = []
    
    for idx, row in df.iterrows():
        total += 1
        name = row['Name']
        email = row['Email']
        
        # Optional fields with defaults
        college = row.get('College Name', 'N/A')
        student_id = row.get('TechieHelp Student ID', 'N/A')
        domain = row.get('Domain', 'N/A')
        duration = row.get('Duration', 'N/A')
        issue_date = row.get('Date of Issued', 'N/A')
        
        print(f"\n📧 Sending to {name} ({email})...")
        
        # Find PDF
        safe_name = sanitize_filename(str(name))
        pdf_filename = f"offer_letter_{safe_name}.pdf"
        pdf_path = os.path.join(offer_letters_dir, pdf_filename)
        
        if not os.path.exists(pdf_path):
            error = f"PDF not found: {pdf_path}"
            print(f"❌ {error}")
            failed.append({'name': name, 'email': email, 'error': error})
            continue
        
        # Personalized body
        body = f"""Dear {name},

Congratulations on your selection!

Here are your offer details:
• College Name: {college}
• TechieHelp Student ID: {student_id}
• Domain: {domain}
• Duration: {duration}
• Date of Issued: {issue_date}

Please find your Offer Letter PDF attached.

We look forward to your contributions!

Best regards,
TechieHelp Team"""
        
        # Send
        try:
            success, msg = sender.send_email(email, name, subject, body, pdf_path)
            if success:
                print(f"✅ {msg}")
                sent += 1
            else:
                print(f"❌ Failed: {msg}")
                failed.append({'name': name, 'email': email, 'error': msg})
        except Exception as e:
            error = f"Unexpected error: {str(e)}"
            print(f"❌ {error}")
            failed.append({'name': name, 'email': email, 'error': error})
    
    # Summary
    print("\n" + "="*50)
    print("📈 SUMMARY")
    print("="*50)
    print(f"Total students: {total}")
    print(f"Emails sent: {sent}")
    print(f"Emails failed: {len(failed)}")
    
    if failed:
        print("\n❌ Failed emails:")
        for f in failed:
            print(f"  - {f['name']} ({f['email']}): {f['error']}")

if __name__ == '__main__':
    print("Bulk Offer Letter Email Sender")
    print("Requirements: pip install pandas openpyxl")
    print("Assumes: students.xlsx and offer_letters/*.pdf exist")
    print("-" * 50)
    main()

