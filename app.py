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

st.set_page_config(page_title="Bulk Generator", page_icon="🏆", layout="wide")

OFFER_REQUIRED_COLUMNS = ['name', 'email', 'domain', 'duration', 'start_date', 'college_name', 'techiehelp_student_id']
CERT_REQUIRED_COLUMNS = ['name', 'email', 'student_id', 'college_name', 'domain', 'start_date', 'end_date']
SHARED_REQUIRED_COLUMNS = list(set(OFFER_REQUIRED_COLUMNS + CERT_REQUIRED_COLUMNS))  # Case-insensitive unique columns
OFFER_OUTPUT_FOLDER = 'offer_letters'
CERT_OUTPUT_FOLDER = 'certificates'

for folder in [OFFER_OUTPUT_FOLDER, CERT_OUTPUT_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

def sanitize_filename(name):
    name = re.sub(r'[^a-zA-Z0-9\s-]', '', name)
    return name.strip().replace(' ', '_')

def validate_excel(df, offer_columns, cert_columns):
    if df.empty:
        return "Excel file is empty."
    # Normalize columns: lower, replace space/underscore with _
    df_normalized = df.copy()
    df_normalized.columns = df_normalized.columns.str.lower().str.replace(r'[ _]', '_', regex=True)
    df_cols_lower = set(df_normalized.columns.str.lower())
    
    offer_req_lower = [col.lower().replace(' ', '_') for col in offer_columns]
    cert_req_lower = [col.lower().replace(' ', '_') for col in cert_columns]
    
    missing_offer = [col for col in offer_req_lower if col not in df_cols_lower]
    missing_cert = [col for col in cert_req_lower if col not in df_cols_lower]
    
    offer_ready = len(missing_offer) == 0
    cert_ready = len(missing_cert) == 0
    
    msg = f"Offer ready: {'✅' if offer_ready else '❌'} (missing: {missing_offer}). Cert ready: {'✅' if cert_ready else '❌'} (missing: {missing_cert})"
    return msg

def generate_single_certificate(student_data, template_path, output_folder):
    try:
        doc = DocxTemplate(template_path)
        context = {
            'name': student_data.get('name', ''),
            'student_id': student_data.get('student_id', ''),
            'college_name': student_data.get('college_name', ''),
            'domain': student_data.get('domain', ''),
            'start_date': student_data.get('start_date', ''),
            'end_date': student_data.get('end_date', ''),
            'current_date': datetime.now().strftime("%d %B %Y")
        }
        doc.render(context)
        
        safe_name = sanitize_filename(student_data['name'])
        docx_path = os.path.join(output_folder, f"certificate_{safe_name}.docx")
        doc.save(docx_path)
        
        converter = PDFConverter(output_folder)
        pdf_path = os.path.join(output_folder, f"certificate_{safe_name}.pdf")
        success = converter.convert_single(docx_path, pdf_path)
        
        try:
            os.remove(docx_path)
        except:
            pass
            
        return pdf_path if success else None
    except:
        return None

def create_zip_cert(output_folder):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(output_folder):
            for f in files:
                if f.endswith('.pdf') and 'certificate' in f:
                    zf.write(os.path.join(root, f), f)
    buffer.seek(0)
    return buffer.getvalue()

def generate_offer_letter(shared_data):
    """Generate offer letters from shared data (filter rows with offer columns)."""
    offer_rows = [row for row in shared_data if all(row.get(col, '') for col in OFFER_REQUIRED_COLUMNS)]
    st.info(f"Found {len(offer_rows)} offer candidates")
    results = []
    progress_placeholder = st.empty()
    progress_bar = progress_placeholder.progress(0)
    status_text = st.empty()
    failed = 0
    
    for i, row in enumerate(offer_rows):
        progress_bar.progress((i + 1) / len(offer_rows))
        status_text.text(f"Generating offer {i+1}/{len(offer_rows)}: {row['name']}")
        pdf_path = generate_single_letter(row, 'offer_template.docx', OFFER_OUTPUT_FOLDER)
        if pdf_path:
            results.append({'name': row['name'], 'email': row['email'], 'pdf_path': pdf_path, 'type': 'offer'})
        else:
            st.warning(f"PDF failed for {row['name']}")
            failed += 1
    
    progress_placeholder.empty()
    status_text.empty()
    return results

def generate_certificate(shared_data):
    """Generate certificates from shared data (filter rows with cert columns)."""
    cert_rows = [row for row in shared_data if all(row.get(col, '') for col in CERT_REQUIRED_COLUMNS)]
    results = []
    progress_placeholder = st.empty()
    progress_bar = progress_placeholder.progress(0)
    status_text = st.empty()
    
    for i, row in enumerate(cert_rows):
        progress_bar.progress((i + 1) / len(cert_rows))
        status_text.text(f"Generating cert {i+1}/{len(cert_rows)}: {row['name']}")
        pdf_path = generate_single_certificate(row, 'certificate_template.docx', CERT_OUTPUT_FOLDER)
        if pdf_path:
            results.append({'name': row['name'], 'email': row['email'], 'pdf_path': pdf_path, 'type': 'cert'})
    
    progress_placeholder.empty()
    status_text.empty()
    return results

def send_offer_email(offer_list):
    """Send offer emails using EmailSender."""

    if not st.secrets.get("gmail"):
        st.error("❌ Configure .streamlit/secrets.toml with Gmail app password")
        st.stop()
    gmail = st.secrets["gmail"]
    sender = EmailSender(gmail.get("smtp_server"), int(gmail.get("smtp_port")), gmail["sender_email"], gmail["sender_password"])
    results = sender.send_batch_emails(offer_list)
    st.success(f"✅ Sent {results.get('sent', 0)}/{len(offer_list)} offer emails")
    if results.get('errors'):
        st.error(f"Failed {len(results['errors'])} emails")
    return results

def send_certificate_email(cert_list):
    """Send certificate emails using EmailSender."""
    if not st.secrets.get("gmail"):
        st.error("Configure .streamlit/secrets.toml")
        return
    gmail = st.secrets["gmail"]
    sender = EmailSender(gmail.get("smtp_server"), int(gmail.get("smtp_port")), gmail["sender_email"], gmail["sender_password"])
    # Override send_offer_letter to cert body
    original_send = sender.send_offer_letter
    def cert_body(recipient_email, recipient_name, attachment_path):
        subject = "Your Completion Certificate"
        body = f"""Dear {recipient_name},

Congratulations on completing your internship!

Please find attached your Certificate of Completion.

Best regards,
TechieHelp Team"""
        return sender.send_email(recipient_email, recipient_name, subject, body, attachment_path)
    sender.send_offer_letter = cert_body
    results = sender.send_batch_emails(cert_list)
    sender.send_offer_letter = original_send
    st.success(f"Sent {results.get('sent', 0)}/{len(cert_list)} certs")
    return results


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
            'name': student_data['name'],
            'domain': student_data['domain'],
            'duration': student_data['duration'],
            'start_date': student_data['start_date'],
            'college_name': student_data.get('college_name', ''),
            'student_id': student_data.get('techiehelp_student_id', student_data.get('student_id', '')),
            'end_date': student_data.get('end_date', ''),
            'current_date': datetime.now().strftime("%d %B %Y")
        }
        doc.render(context)
        
        safe_name = sanitize_filename(student_data['name'])
        docx_path = os.path.join(output_folder, f"{safe_name}_offer_letter.docx")
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
st.session_state.setdefault('shared_data', [])
st.session_state.setdefault('offer_generated', [])
st.session_state.setdefault('cert_generated', [])
st.session_state.setdefault('offer_count', 0)
st.session_state.setdefault('cert_count', 0)

st.title("🏆 Bulk Generator - Shared Data System")

# Shared Excel Upload (SINGLE upload for both tabs)
st.header("📁 Shared Data Upload")
uploaded_shared = st.file_uploader("Upload Shared Excel (contains both offer & cert data)", type=['xlsx'], key="shared_upload")

if uploaded_shared:
    try:
        df = pd.read_excel(uploaded_shared)
        msg = validate_excel(df, OFFER_REQUIRED_COLUMNS, CERT_REQUIRED_COLUMNS)
        st.info(f"📊 Columns analysis: {msg}")
        shared_data = df.fillna('').to_dict('records')  # Handle NaN as ''
        st.session_state.shared_data = [{k.lower().replace(' ', '_').replace('_', '_'): str(v).strip() if pd.notna(v) else '' for k, v in row.items()} for row in shared_data]
        
        st.success(f"✅ Loaded {len(st.session_state.shared_data)} rows")
        st.dataframe(pd.DataFrame(st.session_state.shared_data), use_container_width=True)
        
        if st.button("🗑️ Clear Shared Data"):
            st.session_state.shared_data = []
            st.session_state.offer_generated = []
            st.session_state.cert_generated = []
            st.session_state.offer_count = 0
            st.session_state.cert_count = 0
            st.rerun()
    except Exception as e:
        st.error(f"Upload error: {e}")

col1, col2 = st.columns(2)
col1.metric("Shared Data", len(st.session_state.shared_data))
col2.metric("Generated", st.session_state.offer_count + st.session_state.cert_count)

st.markdown("---")


tab1, tab2, tab3 = st.tabs(["1️⃣ Offer Letters", "2️⃣ Certificates", "3️⃣ Send Emails"])

with tab1:
    st.header("Offer Letters (using shared data)")
    col1, col2 = st.columns(2)
    offer_template_status = "✅ Found" if os.path.exists('offer_template.docx') else "❌ Missing"
    col1.metric("Offer Template", offer_template_status)
    col2.metric("Offer Generated", len(st.session_state.offer_generated))
    
    if st.button("🚀 Generate Offers from Shared Data", disabled=not st.session_state.shared_data):
        with st.spinner("Generating offers..."):
            st.session_state.offer_generated = generate_offer_letter(st.session_state.shared_data)
            st.session_state.offer_count = len(st.session_state.offer_generated)
        
        st.success(f"✅ {st.session_state.offer_count} offer letters generated!")
        st.subheader("Generated Files:")
        for item in st.session_state.offer_generated:
            filename = os.path.basename(item.get('pdf_path') or item.get('pdf', ''))
            st.success(f"📄 {filename} → {item['name']}")
        
        st.download_button("📦 Download Offers ZIP", create_zip(OFFER_OUTPUT_FOLDER), "offers.zip", "application/zip")
    
    if st.session_state.offer_generated:
        st.info(f"Ready to send: {len(st.session_state.offer_generated)} offers")

    if st.button("📧 Send Offers", disabled=not st.session_state.offer_generated):
        send_offer_email(st.session_state.offer_generated)

with tab2:
    st.header("Certificates (using shared data)")
    col1, col2 = st.columns(2)
    cert_template_status = "✅ Found" if os.path.exists('certificate_template.docx') else "❌ Missing"
    col1.metric("Cert Template", cert_template_status)
    col2.metric("Cert Generated", len(st.session_state.cert_generated))
    
    if st.button("🎓 Generate Certs from Shared Data", disabled=not st.session_state.shared_data):
        with st.spinner("Generating certificates..."):
            st.session_state.cert_generated = generate_certificate(st.session_state.shared_data)
            st.session_state.cert_count = len(st.session_state.cert_generated)
        
        st.success(f"✅ {st.session_state.cert_count} certificates generated!")
        st.subheader("Generated Files:")
        for item in st.session_state.cert_generated:
            filename = os.path.basename(item.get('pdf_path') or item.get('pdf', ''))
            st.success(f"📄 {filename} → {item['name']}")
        
        st.download_button("📦 Download Certs ZIP", create_zip(CERT_OUTPUT_FOLDER), "certs.zip", "application/zip")
    
    if st.session_state.cert_generated:
        st.info(f"Ready to send: {len(st.session_state.cert_generated)} certs")
    
    if st.button("📧 Send Certs", disabled=not st.session_state.cert_generated):
        send_certificate_email(st.session_state.cert_generated)

with tab3:
    st.header("📧 Email Summary & Bulk Send")
    col1, col2, col3 = st.columns(3)
    col1.metric("Offers Ready", st.session_state.offer_count)
    col2.metric("Certs Ready", st.session_state.cert_count)
    col3.metric("Total", st.session_state.offer_count + st.session_state.cert_count)
    
    all_data = st.session_state.offer_generated + st.session_state.cert_generated
    if all_data:
        st.info(f"📊 Total ready: {len(all_data)} ({st.session_state.offer_count} offers + {st.session_state.cert_count} certs)")
    
        if st.button("🚀 Send All Emails", disabled=not all_data):
            with st.spinner("Sending all emails..."):
                sent_offers = 0
                sent_certs = 0
                if st.session_state.offer_generated:
                    results_offer = send_offer_email(st.session_state.offer_generated)
                    sent_offers = results_offer.get('sent', 0)
                if st.session_state.cert_generated:
                    results_cert = send_certificate_email(st.session_state.cert_generated)
                    sent_certs = results_cert.get('sent', 0)
                total_sent = sent_offers + sent_certs
                st.success(f"✅ Sent {total_sent}/{len(all_data)} total emails ({sent_offers} offers + {sent_certs} certs)")
                st.balloons()
    else:
        st.warning("❌ No documents ready to send. Generate offers/certs first.")
    
    st.info("💡 Configure `.streamlit/secrets.toml` for emails:")
    st.code("""
[gmail]
sender_email = "your@gmail.com"
sender_password = "your_app_password"
smtp_server = "smtp.gmail.com"
smtp_port = 587
""")

# App runs directly without main() call
st.caption("✅ Syntax fixed - Bulk offer letter generator ready!")


