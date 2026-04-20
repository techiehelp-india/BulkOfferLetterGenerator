"""
Email Sender Module
-------------------
This module handles sending generated offer letters via email.

Author: Automation Team
"""

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os


class EmailSender:
    """
    A class to send offer letters via email with attachments.
    
    Attributes:
        smtp_server (str): SMTP server address
        smtp_port (int): SMTP server port
        sender_email (str): Sender's email address
        sender_password (str): Sender's email password or app password
    """
    
    def __init__(self, smtp_server='smtp.gmail.com', smtp_port=587, 
                 sender_email=None, sender_password=None):
        """
        Initialize the EmailSender.
        
        Args:
            smtp_server (str): SMTP server address
            smtp_port (int): SMTP server port
            sender_email (str): Sender's email address
            sender_password (str): Sender's email password
        """
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.sender_email = sender_email
        self.sender_password = sender_password
    
    def create_email_message(self, recipient_email, recipient_name, 
                            subject, body, attachment_path=None, is_html=False):
        """
        Create an email message with optional attachment.
        
        Args:
            recipient_email (str): Recipient's email address
            recipient_name (str): Recipient's name
            subject (str): Email subject
            body (str): Email body
            attachment_path (str): Path to attachment file
            is_html (bool): Whether the email body is HTML
            
        Returns:
            MIMEMultipart: Email message object
        """
        # Create message
        message = MIMEMultipart()
        message['From'] = self.sender_email
        message['To'] = recipient_email
        message['Subject'] = subject
        
        # Attach body
        content_type = 'html' if is_html else 'plain'
        message.attach(MIMEText(body, content_type))
        
        # Attach file if provided
        if attachment_path:
            paths = attachment_path if isinstance(attachment_path, list) else [attachment_path]
            for path in paths:
                if os.path.exists(path):
                    with open(path, 'rb') as attachment:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    filename = os.path.basename(path)
                    part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
                    message.attach(part)
        
        return message
    
    def send_email(self, recipient_email, recipient_name, subject, body, 
                   attachment_path=None, is_html=False):
        """
        Send an email to a recipient.
        
        Args:
            recipient_email (str): Recipient's email address
            recipient_name (str): Recipient's name
            subject (str): Email subject
            body (str): Email body
            attachment_path (str): Path to attachment file
            is_html (bool): Whether the email body is HTML
            
        Returns:
            tuple: (success, message)
        """
        try:
            # Create message
            message = self.create_email_message(
                recipient_email, recipient_name, subject, body, attachment_path, is_html
            )
            
            # Connect to SMTP server
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()  # Secure the connection
            
            # Login
            server.login(self.sender_email, self.sender_password)
            
            # Send email
            server.sendmail(self.sender_email, recipient_email, message.as_string())
            
            # Close connection
            server.quit()
            
            return True, f"Email sent successfully to {recipient_email}"
            
        except smtplib.SMTPAuthenticationError:
            return False, "Authentication failed. Please check your email and password."
        except smtplib.SMTPException as e:
            return False, f"SMTP Error: {str(e)}"
        except Exception as e:
            return False, f"Error sending email: {str(e)}"
    
    def send_offer_letter(self, recipient_email, recipient_name, 
                         attachment_path, company_name="TechieHelp"):
        """
        Send an internship offer letter to a student.
        
        Args:
            recipient_email (str): Recipient's email address
            recipient_name (str): Recipient's name
            attachment_path (str): Path to the offer letter
            company_name (str): Company name for the email
            
        Returns:
            tuple: (success, message)
        """
        subject = "🎉 Congratulations! You’re Selected for TechieHelp Internship 2026"
        
        display_name = recipient_name if recipient_name and recipient_name.strip() else "Candidate"
        
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
    📧 support@techiehelp.in</p>
    
    <p>Amit Kumar<br>
    Founder & CEO – TechieHelp<br>
    +91 76738 25079<br>
    Jodhpur, Rajasthan – India</p>
</body>
</html>"""
        
        return self.send_email(recipient_email, recipient_name, subject, body, attachment_path, is_html=True)
    
    def send_batch_emails(self, recipients, attachment_folder=None):
        """
        Send emails to multiple recipients.
        
        Args:
            recipients (list): List of dictionaries with 'email', 'name', and 'attachment' or 'pdf_path' keys
            attachment_folder (str): Folder containing attachments (optional if pdf_path used)
            
        Returns:
            dict: Results with success and error counts
        """
        results = {
            'success': [],
            'errors': [],
            'total': len(recipients),
            'sent': 0
        }
        
        for recipient in recipients:
            try:
                email = recipient.get('email')
                name = recipient.get('name')
                pdf_path = recipient.get('pdf_path') or recipient.get('attachment')
                
                # Get full path if only filename and folder provided
                if attachment_folder and pdf_path and not os.path.isabs(pdf_path):
                    pdf_path = os.path.join(attachment_folder, pdf_path)
                
                # Send email
                success, message = self.send_offer_letter(email, name, pdf_path)
                
                if success:
                    results['success'].append({'email': email, 'name': name})
                    results['sent'] += 1
                else:
                    results['errors'].append({'email': email, 'name': name, 'error': message})
                    
            except Exception as e:
                error_msg = f"Error processing {recipient.get('name', 'Unknown')}: {str(e)}"
                results['errors'].append({'email': email, 'name': name, 'error': error_msg})
        
        return results


def send_offer_letter_email(sender_email, sender_password, recipient_email, 
                           recipient_name, attachment_path, 
                           smtp_server='smtp.gmail.com', smtp_port=587):
    """
    Convenience function to send an offer letter email.
    """
    sender = EmailSender(smtp_server, smtp_port, sender_email, sender_password)
    return sender.send_offer_letter(recipient_email, recipient_name, attachment_path)


# Configuration for common email providers
EMAIL_CONFIG = {
    'gmail': {
        'smtp_server': 'smtp.gmail.com',
        'smtp_port': 587,
        'use_tls': True
    },
    'outlook': {
        'smtp_server': 'smtp.office365.com',
        'smtp_port': 587,
        'use_tls': True
    },
    'yahoo': {
        'smtp_server': 'smtp.mail.yahoo.com',
        'smtp_port': 587,
        'use_tls': True
    }
}


# Main execution for testing
if __name__ == '__main__':
    print("=" * 60)
    print("Email Sender Module - Test Information")
    print("=" * 60)
    
    print("\nThis module can be used to send generated offer letters via email.")
    print("\nExample usage:")
    print("""
    from email_sender import EmailSender
    
    # Initialize sender
    sender = EmailSender(
        smtp_server='smtp.gmail.com',
        smtp_port=587,
        sender_email='your_email@gmail.com',
        sender_password='your_app_password'
    )
    
    # Send offer letter
    success, message = sender.send_offer_letter(
        recipient_email='student@email.com',
        recipient_name='John Smith',
        attachment_path='offer_letters/offer_letter_John_Smith.pdf'
    )
    
    print(message)
    """)
    
    print("\n⚠️  NOTE: For Gmail, you need to use an App Password.")
    print("   Enable 2-Factor Authentication and create an App Password.")

