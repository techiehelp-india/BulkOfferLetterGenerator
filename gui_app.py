"""
GUI Application for Internship Offer Letter Automation
------------------------------------------------------
A Tkinter-based GUI application to generate internship offer letters
from Excel data using a Word template.

Author: Automation Team
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
from datetime import datetime

# Import project modules
from generate_letters import OfferLetterGenerator, generate_letters_from_excel
from email_sender import EmailSender
from pdf_converter import PDFConverter


class OfferLetterGUI:
    """
    Main GUI Application class for Offer Letter Generation.
    
    This class creates a user-friendly interface for:
    - Selecting Excel file and Word template
    - Generating offer letters
    - Converting to PDF
    - Sending emails
    """
    
    def __init__(self, root):
        """
        Initialize the GUI application.
        
        Args:
            root (tk.Tk): The root Tkinter window
        """
        self.root = root
        self.root.title("Internship Offer Letter Automation Tool")
        self.root.geometry("700x650")
        self.root.resizable(False, False)
        
        # Set window icon (if available)
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass  # Ignore if icon not found
        
        # Initialize variables
        self.excel_file = tk.StringVar()
        self.template_file = tk.StringVar()
        self.output_folder = tk.StringVar(value='offer_letters')
        
        # Email settings
        self.sender_email = tk.StringVar()
        self.sender_password = tk.StringVar()
        self.smtp_server = tk.StringVar(value='smtp.gmail.com')
        self.smtp_port = tk.IntVar(value=587)
        
        # Options
        self.convert_to_pdf = tk.BooleanVar(value=False)
        self.send_emails = tk.BooleanVar(value=False)
        
        # Style configuration
        self.setup_styles()
        
        # Create GUI components
        self.create_widgets()
        
        # Center window on screen
        self.center_window()
    
    def setup_styles(self):
        """Configure ttk styles for the application"""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Configure button styles
        self.style.configure('TButton', font=('Arial', 10, 'bold'), padding=6)
        self.style.configure('Primary.TButton', background='#0078D4', foreground='white')
        self.style.map('Primary.TButton', background=[('active', '#005A9E')])
        
        # Configure label styles
        self.style.configure('TLabel', font=('Arial', 10), background='#f0f0f0')
        self.style.configure('Header.TLabel', font=('Arial', 12, 'bold'))
        
        # Configure frame styles
        self.style.configure('Card.TFrame', background='white', relief='raised', borderwidth=1)
    
    def center_window(self):
        """Center the window on the screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_widgets(self):
        """Create all GUI widgets"""
        # Main container
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(
            main_frame, 
            text="📧 Internship Offer Letter Generator",
            font=('Arial', 18, 'bold'),
            foreground='#0078D4'
        )
        title_label.pack(pady=(0, 20))
        
        # File Selection Section
        self.create_file_selection_section(main_frame)
        
        # Options Section
        self.create_options_section(main_frame)
        
        # Email Settings Section
        self.create_email_section(main_frame)
        
        # Generate Button
        self.create_generate_button(main_frame)
        
        # Status Display
        self.create_status_display(main_frame)
        
        # Footer
        self.create_footer(main_frame)
    
    def create_file_selection_section(self, parent):
        """Create the file selection section"""
        # Section frame
        section_frame = ttk.LabelFrame(parent, text="📁 File Selection", padding=10)
        section_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Excel File
        excel_frame = ttk.Frame(section_frame)
        excel_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(excel_frame, text="Excel File:", width=12).pack(side=tk.LEFT)
        ttk.Entry(excel_frame, textvariable=self.excel_file, width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            excel_frame, 
            text="Browse", 
            command=lambda: self.browse_file('excel')
        ).pack(side=tk.LEFT)
        
        # Template File
        template_frame = ttk.Frame(section_frame)
        template_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(template_frame, text="Template:", width=12).pack(side=tk.LEFT)
        ttk.Entry(template_frame, textvariable=self.template_file, width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            template_frame, 
            text="Browse", 
            command=lambda: self.browse_file('template')
        ).pack(side=tk.LEFT)
        
        # Output Folder
        output_frame = ttk.Frame(section_frame)
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_frame, text="Output Folder:", width=12).pack(side=tk.LEFT)
        ttk.Entry(output_frame, textvariable=self.output_folder, width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            output_frame, 
            text="Browse", 
            command=self.browse_output_folder
        ).pack(side=tk.LEFT)
    
    def create_options_section(self, parent):
        """Create the options section"""
        section_frame = ttk.LabelFrame(parent, text="⚙️ Options", padding=10)
        section_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Checkboxes
        checkbox_frame = ttk.Frame(section_frame)
        checkbox_frame.pack(fill=tk.X)
        
        ttk.Checkbutton(
            checkbox_frame,
            text="Convert to PDF",
            variable=self.convert_to_pdf,
            command=self.toggle_email_options
        ).pack(side=tk.LEFT, padx=10)
        
        ttk.Checkbutton(
            checkbox_frame,
            text="Send via Email",
            variable=self.send_emails,
            command=self.toggle_email_options
        ).pack(side=tk.LEFT, padx=10)
    
    def create_email_section(self, parent):
        """Create the email settings section"""
        self.email_frame = ttk.LabelFrame(parent, text="📧 Email Settings", padding=10)
        self.email_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Sender Email
        email_frame = ttk.Frame(self.email_frame)
        email_frame.pack(fill=tk.X, pady=3)
        
        ttk.Label(email_frame, text="Your Email:", width=12).pack(side=tk.LEFT)
        ttk.Entry(email_frame, textvariable=self.sender_email, width=35).pack(side=tk.LEFT, padx=5)
        
        # Password
        password_frame = ttk.Frame(self.email_frame)
        password_frame.pack(fill=tk.X, pady=3)
        
        ttk.Label(password_frame, text="Password:", width=12).pack(side=tk.LEFT)
        ttk.Entry(
            password_frame, 
            textvariable=self.sender_password, 
            width=35, 
            show="*"
        ).pack(side=tk.LEFT, padx=5)
        
        # SMTP Server
        smtp_frame = ttk.Frame(self.email_frame)
        smtp_frame.pack(fill=tk.X, pady=3)
        
        ttk.Label(smtp_frame, text="SMTP Server:", width=12).pack(side=tk.LEFT)
        ttk.Entry(smtp_frame, textvariable=self.smtp_server, width=20).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(smtp_frame, text="Port:", width=6).pack(side=tk.LEFT)
        ttk.Entry(smtp_frame, textvariable=self.smtp_port, width=6).pack(side=tk.LEFT, padx=5)
        
        # Help text
        help_label = ttk.Label(
            self.email_frame,
            text="⚠️ For Gmail, use App Password (not your regular password)",
            font=('Arial', 8),
            foreground='gray'
        )
        help_label.pack(pady=(5, 0))
        
        # Initially disable email frame
        self.toggle_email_options()
    
    def create_generate_button(self, parent):
        """Create the main generate button"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.generate_button = ttk.Button(
            button_frame,
            text="🚀 Generate Offer Letters",
            style='Primary.TButton',
            command=self.start_generation
        )
        self.generate_button.pack(fill=tk.X, pady=5)
    
    def create_status_display(self, parent):
        """Create the status display area"""
        # Section frame
        section_frame = ttk.LabelFrame(parent, text="📊 Status", padding=10)
        section_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Progress bar
        self.progress = ttk.Progressbar(
            section_frame, 
            mode='determinate', 
            length=100
        )
        self.progress.pack(fill=tk.X, pady=(0, 10))
        
        # Status text area
        self.status_text = scrolledtext.ScrolledText(
            section_frame,
            height=10,
            width=70,
            font=('Courier', 9),
            wrap=tk.WORD,
            state='disabled'
        )
        self.status_text.pack(fill=tk.BOTH, expand=True)
        
        # Configure tags for colored text
        self.status_text.tag_config('success', foreground='green')
        self.status_text.tag_config('error', foreground='red')
        self.status_text.tag_config('info', foreground='blue')
        self.status_text.tag_config('warning', foreground='orange')
    
    def create_footer(self, parent):
        """Create the footer"""
        footer_frame = ttk.Frame(parent)
        footer_frame.pack(fill=tk.X)
        
        version_label = ttk.Label(
            footer_frame,
            text="Version 1.0 | Internship Offer Letter Automation Tool",
            font=('Arial', 8),
            foreground='gray'
        )
        version_label.pack()
    
    def browse_file(self, file_type):
        """Browse for Excel or template file"""
        if file_type == 'excel':
            file_path = filedialog.askopenfilename(
                title="Select Excel File",
                filetypes=[
                    ("Excel files", "*.xlsx *.xls"),
                    ("All files", "*.*")
                ]
            )
            if file_path:
                self.excel_file.set(file_path)
                self.log_message(f"Selected Excel file: {os.path.basename(file_path)}", 'info')
        
        elif file_type == 'template':
            file_path = filedialog.askopenfilename(
                title="Select Word Template",
                filetypes=[
                    ("Word documents", "*.docx"),
                    ("All files", "*.*")
                ]
            )
            if file_path:
                self.template_file.set(file_path)
                self.log_message(f"Selected template: {os.path.basename(file_path)}", 'info')
    
    def browse_output_folder(self):
        """Browse for output folder"""
        folder_path = filedialog.askdirectory(
            title="Select Output Folder"
        )
        if folder_path:
            self.output_folder.set(folder_path)
            self.log_message(f"Selected output folder: {folder_path}", 'info')
    
    def toggle_email_options(self):
        """Toggle email section based on checkbox"""
        state = 'normal' if self.send_emails.get() else 'disabled'
        
        for child in self.email_frame.winfo_children():
            if isinstance(child, (ttk.Frame, ttk.Entry, ttk.Label)):
                try:
                    child.configure(state=state)
                except:
                    pass
    
    def log_message(self, message, tag='info'):
        """Add a message to the status display"""
        self.status_text.config(state='normal')
        
        # Add timestamp
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        
        self.status_text.insert(tk.END, formatted_message, tag)
        self.status_text.see(tk.END)
        self.status_text.config(state='disabled')
    
    def validate_inputs(self):
        """Validate all input fields"""
        # Check Excel file
        if not self.excel_file.get():
            messagebox.showerror("Error", "Please select an Excel file!")
            return False
        
        if not os.path.exists(self.excel_file.get()):
            messagebox.showerror("Error", "Excel file not found!")
            return False
        
        # Check template file
        if not self.template_file.get():
            messagebox.showerror("Error", "Please select a Word template!")
            return False
        
        if not os.path.exists(self.template_file.get()):
            messagebox.showerror("Error", "Template file not found!")
            return False
        
        # Check email settings if enabled
        if self.send_emails.get():
            if not self.sender_email.get():
                messagebox.showerror("Error", "Please enter your email address!")
                return False
            if not self.sender_password.get():
                messagebox.showerror("Error", "Please enter your password!")
                return False
        
        return True
    
    def start_generation(self):
        """Start the offer letter generation process"""
        if not self.validate_inputs():
            return
        
        # Disable generate button during processing
        self.generate_button.config(state='disabled')
        
        # Reset progress
        self.progress['value'] = 0
        
        # Start generation in separate thread
        thread = threading.Thread(target=self.generate_letters_thread)
        thread.daemon = True
        thread.start()
    
    def generate_letters_thread(self):
        """Run generation in separate thread"""
        try:
            # Clear status
            self.root.after(0, lambda: self.status_text.config(state='normal'))
            self.root.after(0, lambda: self.status_text.delete('1.0', tk.END))
            self.root.after(0, lambda: self.status_text.config(state='disabled'))
            
            self.log_message("Starting offer letter generation...", 'info')
            
            # Create generator
            generator = OfferLetterGenerator(
                self.excel_file.get(),
                self.template_file.get(),
                self.output_folder.get()
            )
            
            # Generate letters
            results = generator.generate_all_letters()
            
            # Update progress
            self.root.after(0, lambda: self.progress.config(value=50))
            
            # Log results
            self.log_message(f"Generated {results['generated']} out of {results['total']} letters", 'success')
            
            # Convert to PDF if enabled
            if self.convert_to_pdf.get() and results['generated'] > 0:
                self.log_message("\nConverting to PDF...", 'info')
                converter = PDFConverter(self.output_folder.get())
                pdf_results = converter.convert_all_in_folder()
                self.log_message(f"Converted {pdf_results['converted']} files to PDF", 'success')
            
            # Send emails if enabled
            if self.send_emails.get() and results['success']:
                self.log_message("\nSending emails...", 'info')
                
                sender = EmailSender(
                    smtp_server=self.smtp_server.get(),
                    smtp_port=self.smtp_port.get(),
                    sender_email=self.sender_email.get(),
                    sender_password=self.sender_password.get()
                )
                
                # Prepare recipients
                recipients = []
                for success_info in results['success']:
                    filename = os.path.basename(success_info['file'])
                    recipients.append({
                        'email': success_info['email'],
                        'name': success_info['name'],
                        'attachment': filename
                    })
                
                # Send emails
                email_results = sender.send_batch_emails(
                    recipients,
                    self.output_folder.get()
                )
                
                self.log_message(f"Sent {email_results['sent']} emails successfully", 'success')
            
            # Log errors if any
            if results['errors']:
                self.log_message("\nErrors:", 'warning')
                for error in results['errors']:
                    self.log_message(f"  - {error}", 'error')
            
            # Complete progress
            self.root.after(0, lambda: self.progress.config(value=100))
            self.log_message("\n✅ Generation complete!", 'success')
            
            # Show completion message
            self.root.after(0, lambda: messagebox.showinfo(
                "Success",
                f"Generated {results['generated']} offer letters successfully!\n\n"
                f"Files saved in: {self.output_folder.get()}"
            ))
            
        except Exception as e:
            self.log_message(f"Error: {str(e)}", 'error')
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        
        finally:
            # Re-enable generate button
            self.root.after(0, lambda: self.generate_button.config(state='normal'))


def main():
    """Main entry point for the application"""
    # Create root window
    root = tk.Tk()
    
    # Create application
    app = OfferLetterGUI(root)
    
    # Run main loop
    root.mainloop()


if __name__ == '__main__':
    main()

