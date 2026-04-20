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
    subject = "🎉 Congratulations! You’re Selected for TechieHelp Internship 2026"
    
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
        
        display_name = name if name and str(name).strip() else "Candidate"
        
        # Personalized HTML body
        body = f"""<html>
<head>
<style>
    body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333333; }}
    a {{ color: #0066cc; text-decoration: none; }}
    a:hover {{ text-decoration: underline; }}
</style>
</head>
<body>
    <p>Dear {display_name},</p>
    
    <p>You have been selected for the TECHIEHELP – Industry-Oriented Internship Program 2026.<br>
    We’re excited to have you onboard as you take the next step toward skill development and career growth.</p>
    
    <p>📝 <strong>Important:</strong> Please ensure that you select your internship department correctly (e.g., Android Developer) during further processes.</p>
    
    <p>🏅 <strong>Claim Your LinkedIn Internship Badge:</strong><br>
    <a href="https://twibbo.nz/techiehelpinternsbadges">https://twibbo.nz/techiehelpinternsbadges</a></p>
    
    <p>📌 <strong>Learn more about your internship program:</strong><br>
    <a href="https://www.techiehelp.in/careers/training-internships">https://www.techiehelp.in/careers/training-internships</a></p>
    
    <p>⚠️ Please check your Inbox & Spam folder regularly for important updates and communication.</p>
    
    <p>Best Regards,<br>
    Team TechieHelp<br>
    Empowering Students. Building Futures.</p>
    
    <p>🌐 <a href="https://www.techiehelp.in">https://www.techiehelp.in</a><br>
    📧 <a href="mailto:support@techiehelp.in">support@techiehelp.in</a></p>
    
    <p>Amit Kumar<br>
    Founder & CEO – TechieHelp<br>
    +91 76738 25079<br>
    Jodhpur, Rajasthan – India</p>
</body>
</html>"""
        
        # Send
        try:
            success, msg = sender.send_email(email, name, subject, body, pdf_path, is_html=True)
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

