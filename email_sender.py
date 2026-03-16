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
                            subject, body, attachment_path=None):
        """
        Create an email message with optional attachment.
        
        Args:
            recipient_email (str): Recipient's email address
            recipient_name (str): Recipient's name
            subject (str): Email subject
            body (str): Email body
            attachment_path (str): Path to attachment file
            
        Returns:
            MIMEMultipart: Email message object
        """
        # Create message
        message = MIMEMultipart()
        message['From'] = self.sender_email
        message['To'] = recipient_email
        message['Subject'] = subject
        
        # Attach body
        message.attach(MIMEText(body, 'plain'))
        
        # Attach file if provided
        if attachment_path and os.path.exists(attachment_path):
            # Open the file in binary mode
            with open(attachment_path, 'rb') as attachment:
                # Create MIMEBase object
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            
            # Encode the file
            encoders.encode_base64(part)
            
            # Add header
            filename = os.path.basename(attachment_path)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename= {filename}'
            )
            
            # Attach to message
            message.attach(part)
        
        return message
    
    def send_email(self, recipient_email, recipient_name, subject, body, 
                   attachment_path=None):
        """
        Send an email to a recipient.
        
        Args:
            recipient_email (str): Recipient's email address
            recipient_name (str): Recipient's name
            subject (str): Email subject
            body (str): Email body
            attachment_path (str): Path to attachment file
            
        Returns:
            tuple: (success, message)
        """
        try:
            # Create message
            message = self.create_email_message(
                recipient_email, recipient_name, subject, body, attachment_path
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
        subject = "Your Internship Offer Letter"
        
        body = f"""Dear {recipient_name},

Congratulations!

Please find attached your Internship Offer Letter from TechieHelp.

We are excited to have you onboard and look forward to your contributions.

Best regards,
TechieHelp Team"""
        
        return self.send_email(recipient_email, recipient_name, subject, body, attachment_path)
    
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

