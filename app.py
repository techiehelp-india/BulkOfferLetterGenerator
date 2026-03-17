"""
Internship Offer Letter Generator - Streamlit Application
---------------------------------------------------------
A modern web app that generates bulk internship offer letters from Excel using Word template.
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

st.set_page_config(page_title="Offer Letter Generator", page_icon=":envelope:", layout="wide")

REQUIRED_COLUMNS = ['Name', 'Email', 'Domain', 'Duration', 'Start Date', 'College Name', 'TechieHelp Student Id']
OUTPUT_FOLDER = 'offer_letters'

if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

def sanitize_filename(name):
    name = re.sub(r'[^\\w\\s-]', '', name)
    return name.strip().replace(' ', '_')

def validate_excel(df):
    if df.empty:
        return False, "Excel file is empty."
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    return len(missing) == 0, ', '.join(missing) if missing else "OK"

def clean_data(df):
    df = df.dropna(how='all').dropna(subset=['Name', 'Email'])
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip()
    return df

def generate_single_letter(student_data, template_path, output_folder):
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
        docx_path = os.path.join(output_folder, f"{safe_name}_Offer_Letter.docx")
        doc.save(docx_path)
        
        converter = PDFConverter(output_folder)
        pdf_path = os.path.join(output_folder, f"offer_letter_{safe_name}.pdf")
        success = converter.convert_single(docx_path, pdf_path)
        
        try:
            os.remove(docx_path)
        except:
            pass
            
        return pdf_path if success else None
    except:
        return None

def create_zip(output_folder):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(output_folder):
            for f in files:
                if f.endswith('.pdf'):
                    zf.write(os.path.join(root, f), f)
    buffer.seek(0)
    return buffer.getvalue()

# Session state
if 'data' not in st.session_state:
    st.session_state.data = []
if 'count' not in st.session_state:
    st.session_state.count = 0

st.title(":envelope_with_arrow: Bulk Offer Letter Generator")

tab1, tab2 = st.tabs(["Generate", "Send Email"])

with tab1:
    uploaded = st.file_uploader("Upload Excel", type=['xlsx'])
    
    col1, col2 = st.columns(2)
    template_status = "✅ Found" if os.path.exists('offer_template.docx') else "❌ Missing"
    col1.metric("Template", template_status)
    col2.metric("Emails Ready", st.session_state.count)
    
    if st.button("🚀 Generate Letters", disabled=not uploaded):
        try:
            df = pd.read_excel(uploaded)
            valid, msg = validate_excel(df)
            if not valid:
                st.error(msg)
            else:
                df_clean = clean_data(df)
                progress = st.progress(0)
                results = []
                failed = []
                
                for i, row in df_clean.iterrows():
                    progress.progress((i+1) / len(df_clean))
                    pdf = generate_single_letter(row.to_dict(), 'offer_template.docx', OUTPUT_FOLDER)
                    if pdf:
                        results.append({'name': row['Name'], 'email': row['Email'], 'pdf': pdf})
                    else:
                        failed.append(row['Name'])
                
                st.session_state.data = results
                st.session_state.count = len(results)
                
                col_a, col_b = st.columns(2)
                with col_a:
                    st.success(f"✅ {len(results)}/{len(df_clean)} success")
                with col_b:
                    if failed:
                        st.warning(f"⚠️ {len(failed)} failed")
                
                st.download_button("📦 Download ZIP", create_zip(OUTPUT_FOLDER), "letters.zip", "application/zip")
        except Exception as e:
            st.error(f"Error: {e}")

with tab2:
    st.info("Configure `.streamlit/secrets.toml`:")
    st.code("""
[gmail]
sender_email = "your@gmail.com"
sender_password = "app_password"
smtp_server = "smtp.gmail.com"
smtp_port = 587
    """)
    
    if st.button("📧 Send All", disabled=not st.session_state.data):
        try:
            gmail = st.secrets["gmail"]
            sender = EmailSender(gmail.get("smtp_server", "smtp.gmail.com"), 
                               int(gmail.get("smtp_port", 587)), 
                               gmail["sender_email"], 
                               gmail["sender_password"])
            results = sender.send_batch_emails(st.session_state.data)
            st.success(f"Sent {results.get('sent', 0)}/{len(st.session_state.data)}")
            if results.get('sent'):
                st.balloons()
        except Exception as e:
            st.error(f"Send failed: {e}")

# App runs directly without main() call
st.caption("✅ Syntax fixed - Bulk offer letter generator ready!")


