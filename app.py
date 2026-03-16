"""
Internship Offer Letter Generator - Streamlit Application
---------------------------------------------------------
A modern web application that generates internship offer letters
in bulk from an Excel file using a Word template.

Author: Automation Team
Enhanced with email sending feature.
"""

import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import os
import zipfile
import io
from datetime import datetime
import re
from pdf_converter import PDFConverter
from email_sender import EmailSender


# ============================================================================
# CONFIGURATION
# ============================================================================

# Page configuration
st.set_page_config(
    page_title="Internship Offer Letter Generator",
    page_icon="📧",
    layout="centered",
    initial_sidebar_state="expanded"
)

# Required columns in Excel (Email required for sending)
REQUIRED_COLUMNS = ['Name', 'Email', 'Domain', 'Duration', 'Start Date', 'College Name', 'TechieHelp Student Id']  # End Date optional

# Output folder
OUTPUT_FOLDER = 'offer_letters'


# ============================================================================
# FUNCTIONS
# ============================================================================

def sanitize_filename(name):
    """
    Sanitize a name to be used as a filename.
    """
    name = re.sub(r'[^\w\s-]', '', name)
    name = name.strip().replace(' ', '_')
    return name


def validate_excel(df):
    """
    Validate Excel columns.
    """
    if df.empty:
        return False, "The Excel file is empty. Please add student data."
    
    missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    
    if missing_columns:
        return False, f"Missing required columns: {', '.join(missing_columns)}. Note: 'Email' is required for sending feature."
    
    return True, ""


def clean_data(df):
    """
    Clean the data.
    """
    df = df.dropna(how='all')
    df = df.dropna(subset=['Name', 'Email'])
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].str.strip()
    return df


def generate_single_letter(student_data, template_path, output_folder):
    """
    Generate single PDF without st.error calls.
    """
    try:
        doc = DocxTemplate(template_path)
        
        context = {
            'name': student_data['Name'],
            'domain': student_data['Domain'],
            'duration': student_data['Duration'],
            'start_date': student_data['Start Date'],
            'college_name': student_data.get('College Name', ''),
            'student_id': student_data.get('TechieHelp Student Id', ''),
            'end_date': student_data.get('End Date', ''),
            'current_date': datetime.now().strftime("%d %B %Y")
        }
        
        doc.render(context)
        
        safe_name = sanitize_filename(student_data['Name'])
        docx_filename = f"{safe_name}_Offer_Letter.docx"
        docx_path = os.path.join(output_folder, docx_filename)
        
        doc.save(docx_path)
        
        converter = PDFConverter(output_folder)
        pdf_filename = f"offer_letter_{safe_name}.pdf"
        pdf_path = os.path.join(output_folder, pdf_filename)
        pdf_result = converter.convert_single(docx_path, pdf_path)
        
        try:
            os.remove(docx_path)
        except:
            pass
            
        return pdf_path if pdf_result else None
        
    except:
        return None



def create_zip_file(output_folder):
    """
    Create ZIP of PDFs.
    """
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(output_folder):
            for file in files:
                if file.endswith('.pdf'):
                    file_path = os.path.join(root, file)
                    arcname = file
                    zipf.write(file_path, arcname)
    
    zip_buffer.seek(0)
    return zip_buffer.getvalue()


def clear_output():
    """
    Clear output folder and session state.
    """
    for file in os.listdir(OUTPUT_FOLDER):
        file_path = os.path.join(OUTPUT_FOLDER, file)
        if os.path.isfile(file_path):
            os.remove(file_path)
    st.session_state.email_data = []
    st.session_state.generated_count = 0
    st.session_state.secrets_valid = False
    st.success("✅ Output folder and data cleared!")



# ============================================================================
# STREAMLIT UI
# ============================================================================

def main():
    st.markdown("""
        <style>
        .main-header {
            font-size: 36px;
            font-weight: bold;
            color: #1f77b4;
            text-align: center;
            margin-bottom: 30px;
        }
        .success-message {
            padding: 15px;
            background-color: #d4edda;
            border-radius: 5px;
            color: #155724;
        }
        </style>
    """, unsafe_allow_html=True)
    
    # Title
    st.markdown('<p class="main-header">📧 Internship Offer Letter Generator</p>', unsafe_allow_html=True)
    
    # Output folder
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)
    
    # Sidebar Email Config
    with st.sidebar:
        st.header("🔒 Secure Email Config")
        st.info("👉 Edit `.streamlit/secrets.toml` with Gmail + App Password")
        st.markdown("[Generate Gmail App Password](https://support.google.com/accounts/answer/185833)")
        
        # Safe secrets validation
        try:
            test_secrets = st.secrets.get("gmail", {})
            if all(k in test_secrets for k in ["sender_email", "sender_password"]):
                st.success("✅ Gmail Config Valid - Ready to Send!")
                st.session_state.secrets_valid = True
            else:
                raise KeyError("Missing sender_email/password")
        except Exception as e:
            st.error(f"❌ Invalid Gmail Config: {str(e)[:100]}...")
            st.info("📝 Fix: Update secrets.toml with REAL values & restart app")
            st.session_state.secrets_valid = False

    
    # Session state
    if 'email_data' not in st.session_state:
        st.session_state.email_data = []
    if 'generated_count' not in st.session_state:
        st.session_state.generated_count = 0
    if 'secrets_valid' not in st.session_state:
        st.session_state.secrets_valid = False

    
    # Upload
    col1, col2 = st.columns([2, 1])
    uploaded_file = col1.file_uploader("📁 Upload Excel (.xlsx)", type=['xlsx', 'xls'])
    template_ok = os.path.exists('offer_template.docx')
    col2.success("✅ Template OK") if template_ok else col2.warning("⚠️ Missing offer_template.docx")
    
    # Buttons - SEND EMAIL RIGHT OF GENERATE
    col1, col2, col3 = st.columns(3)
    generate_button = col1.button("🚀 Generate Offer Letters", type="primary", disabled=not uploaded_file)
    send_button = col2.button("📧 Send Email", disabled=(not st.session_state.email_data or not st.session_state.secrets_valid))
    col3.button("🔄 Clear Cache", on_click=lambda: st.cache_data.clear())
    col3.button("🗑️ Clear All", on_click=clear_output)

    
    if generate_button:
        with st.status("Generating offer letters...", expanded=True) as status:
            try:
                status.update(label="Validating data...")
                df = pd.read_excel(uploaded_file)
                valid, msg = validate_excel(df)
                if not valid:
                    st.error(msg)
                    st.rerun()
                
                status.update(label="Cleaning data...")
                df = clean_data(df)
                if df.empty:
                    st.error("No data")
                    st.rerun()
                
            status.update(label="Generating letters...")

            try:
                status.update(label="Clearing output & generating...")
                # Clear & generate
                for f in os.listdir(OUTPUT_FOLDER):
                    os.remove(os.path.join(OUTPUT_FOLDER, f))
                
                progress_bar = st.progress(0)
                status_list = []
                data = []
                count = 0
                
                for i, row in df.iterrows():
                    status.update(label=f"Processing {row['Name']} ({i+1}/{len(df)})...")
                    progress_bar.progress((i+1)/len(df))
                    
                    try:
                        pdf = generate_single_letter(row.to_dict(), 'offer_template.docx', OUTPUT_FOLDER)
                        if pdf and os.path.exists(pdf):
                            count += 1
                            data.append({'name': row['Name'], 'email': row['Email'], 'pdf_path': pdf})
                        else:
                            status_list.append(f"⚠️ Failed PDF: {row['Name']}")
                    except Exception as row_e:
                        status_list.append(f"❌ Error {row['Name']}: {str(row_e)[:50]}")
                
                st.session_state.email_data = data
                st.session_state.generated_count = count
                
                status.update(label="Complete!")
                
                if count > 0:
                    st.success(f"✅ Generated {count}/{len(df)} PDFs")
                    st.download_button("📥 Download ZIP", create_zip_file(OUTPUT_FOLDER), "offer_letters.zip", "application/zip")
                else:
                    st.warning("No PDFs generated. Check template/path/errors.")
                
                if status_list:
                    st.info("Details:")
                    for s in status_list[:10]:  # First 10
                        st.write(s)
            except Exception as e:
                status.update(label="Error!")
                st.error(f"Generation failed: {str(e)}")

    
    if send_button:
        with st.spinner('Sending emails...'):
            try:
                gmail = st.secrets["gmail"]
                sender = EmailSender(gmail["smtp_server"], gmail["smtp_port"], gmail["sender_email"], gmail["sender_password"])
                results = sender.send_batch_emails(st.session_state.email_data)
                
                st.metric("Sent", results['sent'], results['total'])
                st.success(f"✅ {results['sent']}/{results['total']} sent!")
                
                if results['errors']:
                    with st.expander(f"❌ {len(results['errors'])} errors"):
                        for e in results['errors']:
                            st.error(f"{e['name']} ({e['email']}): {e['error']}")
                
                if results['sent'] > 0:  # Conditional balloons
                    st.balloons()
                    
            except Exception as e:
                st.error(f"Send Error: {str(e)}")
                st.info("💡 Clear cache & check secrets.toml")

    
    if st.session_state.generated_count:
        st.info(f"📊 {st.session_state.generated_count} ready to send")
    
    # Footer
    st.markdown("---")
    st.markdown("© 2024 Bulk Offer Letter Generator | Send Email Feature Added")


if __name__ == "__main__":
    main()

