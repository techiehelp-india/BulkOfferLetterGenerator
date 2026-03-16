"""
Internship Offer Letter Generator - Fixed Syntax Version
No try/except indentation errors. Uses app_fixed.py base + fixes.
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
        return False, \"Empty Excel\"
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    return len(missing) == 0, ', '.join(missing) if missing else \"OK\"

def clean_data(df):
    df = df.dropna(how='all').dropna(subset=['Name', 'Email'])
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].str.strip()
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
        docx_path = os.path.join(output_folder, f\"{safe_name}_Offer_Letter.docx\")
        doc.save(docx_path)
        converter = PDFConverter(output_folder)
        pdf_path = os.path.join(output_folder, f\"offer_letter_{safe_name}.pdf\")
        success = converter.convert_single(docx_path, pdf_path)
        try:
            os.remove(docx_path)
        except:
            pass
        return pdf_path if success and os.path.exists(pdf_path) else None
    except Exception:
        return None

def create_zip_file(output_folder):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(output_folder):
            for f in files:
                if f.endswith('.pdf'):
                    zf.write(os.path.join(root, f), f)
    buffer.seek(0)
    return buffer.getvalue()

st.session_state.setdefault('email_data', [])
st.session_state.setdefault('generated_count', 0)

st.title("📧 Bulk Offer Letter Generator (Syntax Fixed)")

st.sidebar.header(\"🔒 Email Setup\")
if st.secrets.get(\"gmail\"):
    st.sidebar.success(\"✅ Ready\")
else:
st.sidebar.error("Update secrets.toml")

uploaded_file = st.file_uploader(\"Excel (.xlsx)\", type=['xlsx'])
template_ok = os.path.exists('offer_template.docx')
st.info(f\"Template: {'✅ OK' if template_ok else '⚠️ Missing'}\")

if st.button(\"🚀 Generate\", disabled=not uploaded_file):
    try:
        df = pd.read_excel(uploaded_file)
        is_valid, msg = validate_excel(df)
        if not is_valid:
st.error(f"Invalid: {msg}")
        else:
            df = clean_data(df)
            if df.empty:
                st.error(\"No valid data\")
            else:
                bar = st.progress(0)
                data = []
                status_list = []
                for i, row in df.iterrows():
                    bar.progress((i+1) / len(df))
                    pdf = generate_single_letter(row.to_dict(), 'offer_template.docx', OUTPUT_FOLDER)
                    if pdf:
                        data.append({'name': row['Name'], 'email': row['Email'], 'pdf_path': pdf})
                    else:
                        status_list.append(row['Name'])
                st.session_state.email_data = data
                st.session_state.generated_count = len(data)
                if data:
                    st.success(f\"✅ {len(data)}/{len(df)} PDFs\")
                    st.download_button(\"📥 ZIP\", create_zip_file(OUTPUT_FOLDER), \"letters.zip\", \"application/zip\")
                if status_list:
                    st.warning(f\"⚠️ Failed {len(status_list)}: {', '.join(status_list[:5])}\")
    except Exception as e:
        st.error(f\"Generation error: {str(e)}\")

if st.button(\"📧 Send\", disabled=not st.session_state.email_data or not st.secrets.get(\"gmail\")):
    try:
        gmail = st.secrets[\"gmail\"]
        sender = EmailSender(gmail.get(\"smtp_server\", \"smtp.gmail.com\"), int(gmail.get(\"smtp_port\", 587)), gmail[\"sender_email\"], gmail[\"sender_password\"])
        results = sender.send_batch_emails(st.session_state.email_data)
        st.metric(\"Emails sent\", results.get(\"sent\", 0))
        if results.get(\"sent\", 0) > 0:
            st.balloons()
        if results.get(\"errors\"):
            st.error(f\"Errors: {len(results['errors'])}\")
    except Exception as e:
        st.error(f\"Send error: {str(e)}\")

st.info(f\"Ready letters: {st.session_state.generated_count}\")

