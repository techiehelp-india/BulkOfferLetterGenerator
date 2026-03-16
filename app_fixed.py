"""
Internship Offer Letter Generator - Fixed Version
Stable, no syntax errors, DeltaGenerator safe.
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

st.set_page_config(page_title="Offer Letter Generator", page_icon="📧", layout="centered")

REQUIRED_COLUMNS = ['Name', 'Email', 'Domain', 'Duration', 'Start Date', 'College Name', 'TechieHelp Student Id']
OUTPUT_FOLDER = 'offer_letters'

if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

def sanitize_filename(name):
    name = re.sub(r'[^\w\s-]', '', name)
    return name.strip().replace(' ', '_')

def validate_excel(df):
    if df.empty:
        return False, "Empty Excel"
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        return False, f"Missing: {', '.join(missing)}"
    return True, ""

def clean_data(df):
    return df.dropna(how='all').dropna(subset=['Name', 'Email']).apply(lambda x: x.str.strip() if x.dtype == "object" else x)

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
        converter.convert_single(docx_path, pdf_path)
        try:
            os.remove(docx_path)
        except:
            pass
        return pdf_path if os.path.exists(pdf_path) else None
    except:
        return None

def create_zip_file(output_folder):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w') as zf:
        for root, dirs, files in os.walk(output_folder):
            for f in files:
                if f.endswith('.pdf'):
                    zf.write(os.path.join(root, f), f)
    buffer.seek(0)
    return buffer.getvalue()

if 'generated_count' not in st.session_state:
    st.session_state.generated_count = 0
if 'email_data' not in st.session_state:
    st.session_state.email_data = []

st.title("📧 Bulk Offer Letter Generator")

# Sidebar
with st.sidebar:
    st.header("🔒 Email Setup")
    if st.secrets.get("gmail"):
        st.success("✅ Ready")
    else:
        st.error("Update secrets.toml")

# Upload & buttons
uploaded_file = st.file_uploader("Excel", type=['xlsx'])
if st.button("🚀 Generate", disabled=not uploaded_file):
    try:
        df = pd.read_excel(uploaded_file)
        valid, msg = validate_excel(df)
        if not valid:
            st.error(msg)
        else:
            df = clean_data(df)
            if df.empty:
                st.error("No data")
            else:
                bar = st.progress(0)
                data = []
                for i, row in df.iterrows():
                    bar.progress((i+1)/len(df))
                    try:
                        row_dict = row.to_dict()
                        pdf = generate_single_letter(row_dict, 'offer_template.docx', OUTPUT_FOLDER)
                        docx_path = os.path.join(OUTPUT_FOLDER, f"{sanitize_filename(row_dict['Name'])}_Offer_Letter.docx")
                        if pdf and os.path.exists(pdf):
                            data.append({'name': row['Name'], 'email': row['Email'], 'pdf_path': pdf})
                        elif os.path.exists(docx_path):
                            data.append({'name': row['Name'], 'email': row['Email'], 'pdf_path': docx_path})
                            st.info(f"Generated DOCX for {row['Name']} (PDF fallback)")
                        else:
                            st.warning(f"Failed {row['Name']}")
                    except Exception as row_e:
                        st.error(f"Row {i}: {str(row_e)}")
                st.session_state.email_data = data
                st.session_state.generated_count = len(data)
                st.success(f"✅ {len(data)}/{len(df)} letters generated")
                st.download_button("ZIP", create_zip_file(OUTPUT_FOLDER), "letters.zip", "application/zip")
    except Exception as e:
        st.error(f"Error: {e}")

if st.button("📧 Send Emails", disabled=not st.secrets.get("gmail") or not st.session_state.email_data):
    gmail = st.secrets["gmail"]
    sender = EmailSender(gmail.get("smtp_server", "smtp.gmail.com"), gmail.get("smtp_port", 587), gmail["sender_email"], gmail["sender_password"])
    results = sender.send_batch_emails(st.session_state.email_data)
    st.metric("Sent", results["sent"])
    if results["sent"] > 0:
        st.balloons()

st.info(f"Ready: {st.session_state.generated_count} letters")

