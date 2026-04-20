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

DEBUG = False

OFFER_REQUIRED_COLUMNS = ['name', 'email', 'domain']
CERT_REQUIRED_COLUMNS = ['name', 'email', 'domain']
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
            'student_id': student_data.get('student_id') or student_data.get('techiehelp_student_id') or "N/A",

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
    if DEBUG:
        st.info(f"📊 Total rows loaded: {len(shared_data) if shared_data else 0}")
        if shared_data:
            st.info(f"📋 Columns: {list(shared_data[0].keys())}")
            st.info("📄 First 2 rows:")
            st.json(shared_data[:2])

    offer_rows = []
    skipped = []
    for row in shared_data:
        missing = [col for col in OFFER_REQUIRED_COLUMNS if not str(row.get(col, '')).strip()]
        if not missing:
            offer_rows.append(row)
        else:
            skipped.append({'name': row.get('name', 'N/A'), 'missing': missing})

    st.info(f"✅ Found {len(offer_rows)} offer candidates (skipped {len(skipped)})")
    if skipped:
        st.warning(f"Skipped rows: {skipped[:5]}...")

    if not offer_rows:
        st.error("❌ No valid rows found. Please check your Excel data.")
        return []

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
    if DEBUG:
        st.info(f"📊 Total rows loaded: {len(shared_data) if shared_data else 0}")
        if shared_data:
            st.info(f"📋 Columns: {list(shared_data[0].keys())}")
            st.info("📄 First 2 rows:")
            st.json(shared_data[:2])
    cert_rows = []
    skipped = []
    for row in shared_data:
        missing = [col for col in CERT_REQUIRED_COLUMNS if not row.get(col, '').strip()]
        if not missing:
            cert_rows.append(row)
        else:
            skipped.append({'name': row.get('name', 'N/A'), 'missing': missing})
    st.info(f"✅ Found {len(cert_rows)} cert candidates (skipped {len(skipped)})")
    if skipped:
        st.warning(f"Skipped rows: {skipped[:5]}...")
    results = []
    progress_placeholder = st.empty()
    progress_bar = progress_placeholder.progress(0)
    status_text = st.empty()
    failed = 0
    
    for i, row in enumerate(cert_rows):
        progress_bar.progress((i + 1) / len(cert_rows))
        status_text.text(f"Generating cert {i+1}/{len(cert_rows)}: {row['name']}")
        pdf_path = generate_single_certificate(row, 'certificate_template.docx', CERT_OUTPUT_FOLDER)
        if pdf_path:
            results.append({'name': row['name'], 'email': row['email'], 'pdf_path': pdf_path, 'type': 'cert'})
        else:
            st.warning(f"PDF failed for {row['name']} cert")
            failed += 1
    
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


def execute_bulk_custom_email(recipients, subject, template_message, attachment_path=None):
    if not st.secrets.get("gmail"):
        st.error("❌ Configure .streamlit/secrets.toml with Gmail app password")
        return None
    gmail = st.secrets["gmail"]
    sender = EmailSender(gmail.get("smtp_server"), int(gmail.get("smtp_port")), gmail["sender_email"], gmail["sender_password"])
    
    results = {'sent': 0, 'errors': []}
    progress_placeholder = st.empty()
    progress_bar = progress_placeholder.progress(0)
    status_text = st.empty()
    
    for i, recipient in enumerate(recipients):
        progress_bar.progress((i + 1) / len(recipients))
        email = recipient['email']
        name = recipient.get('name', 'User')
        status_text.text(f"Sending email {i+1}/{len(recipients)}: {email}")
        
        message = template_message.replace('{name}', name)
        success, msg = sender.send_email(email, name, subject, message, attachment_path)
        if success:
            results['sent'] += 1
        else:
            results['errors'].append({'email': email, 'error': msg})
            
    progress_placeholder.empty()
    status_text.empty()
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
            'duration': student_data.get('duration', ''),
            'start_date': student_data.get('start_date', ''),
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
        pdf_path = os.path.join(output_folder, f"{safe_name}_offer_letter.pdf")
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

st.markdown("""
<style>
    .stApp { background-color: #F9FAFB; }
    h1, h2, h3 { color: #111827; font-weight: 600; margin-bottom: 0.2rem !important; }
    h1 { font-size: 1.75rem !important; }
    h2 { font-size: 1.4rem !important; }
    p { margin-bottom: 0.2rem !important; }
    .block-container { padding-top: 1.5rem !important; padding-bottom: 1.5rem !important; }
    
    .stButton>button, .stDownloadButton>button {
        color: white;
        border-radius: 6px;
        border: none;
        padding: 0.35rem 1rem !important;
        font-size: 0.9rem !important;
        font-weight: 500;
        transition: all 0.2s;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
    }
    .stButton>button { background-color: #4F46E5; }
    .stButton>button:hover { background-color: #4338CA; color: white; }
    .stDownloadButton>button { background-color: #22C55E !important; }
    .stDownloadButton>button:hover { background-color: #16A34A !important; }
    
    [data-testid="stSidebar"] { background-color: white; border-right: 1px solid #e5e7eb; }
    div[data-testid="stMetricValue"] { color: #4F46E5; font-size: 1.5rem !important; }
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.title("🏢 Workspace")
    st.markdown("---")
    menu = st.radio("Navigation", 
        ["📥 Data Entry", "📄 Offer Letters", "🎓 Certificates", "📊 Batch Operations", "📧 Custom Mailing"],
        label_visibility="collapsed"
    )
    st.markdown("---")
    st.metric("Total Roster", len(st.session_state.shared_data))
    if st.button("🗑️ Clear Session Data", use_container_width=True):
        st.session_state.shared_data = []
        st.session_state.offer_generated = []
        st.session_state.cert_generated = []
        st.session_state.offer_count = 0
        st.session_state.cert_count = 0
        st.rerun()

if menu == "📥 Data Entry":
    st.header("Data Initialization")
    st.markdown("Upload your employee or student roster to prime the generation pipelines.")
    
    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded_shared = st.file_uploader("Upload Master Excel Workbook", type=['xlsx', 'xls'], key="shared_upload")
        if uploaded_shared:
            try:
                df = pd.read_excel(uploaded_shared)
                msg = validate_excel(df, OFFER_REQUIRED_COLUMNS, CERT_REQUIRED_COLUMNS)
                if DEBUG:
                    st.info(f"📊 Analysis: {msg}")
                shared_data = df.fillna('').to_dict('records')
                normalized_data = []
                for row in shared_data:
                    new_row = {}
                    for k, v in row.items():
                        k_str = str(k).strip().lower().replace(' ', '_')
                        if k_str == 'help_stud':
                            k_str = 'techiehelp_student_id'
                        new_row[k_str] = str(v).strip() if pd.notna(v) else ''
                    normalized_data.append(new_row)
                st.session_state.shared_data = normalized_data
                st.success(f"✅ Roster synchronized with **{len(st.session_state.shared_data)}** verified records.")
                
                if DEBUG:
                    st.dataframe(pd.DataFrame(st.session_state.shared_data), use_container_width=True)
            except Exception as e:
                st.error(f"Operation failed: {e}")
    with col2:
        st.info("💡 **Formatting Rule:** Ensure your core columns include `Name`, `Email`, and `Domain`.")

elif menu == "📄 Offer Letters":
    st.header("Offer Compilations")
    offer_template_status = "Active" if os.path.exists('offer_template.docx') else "Missing Configuration"
    
    col1, col2, col3 = st.columns(3)
    col1.metric("System Template", offer_template_status)
    col2.metric("Eligible Entries", len([r for r in st.session_state.shared_data if all(str(r.get(c, '')).strip() for c in OFFER_REQUIRED_COLUMNS)]) if st.session_state.shared_data else 0)
    col3.metric("Compiled Archive", st.session_state.offer_count)

    st.markdown("---")
    
    act_col, _ = st.columns([1, 3])
    with act_col:
        if st.button("🚀 Process Workload", disabled=not st.session_state.shared_data):
            with st.spinner("Processing rendering tasks..."):
                st.session_state.offer_generated = generate_offer_letter(st.session_state.shared_data)
                st.session_state.offer_count = len(st.session_state.offer_generated)
        
    if len(st.session_state.offer_generated) > 0:
        st.success(f"Processing complete: {st.session_state.offer_count} artifacts generated.")
        
        post_col1, post_col2, post_col3 = st.columns([1, 1, 2])
        with post_col1:
            st.download_button("📦 Download ZIP Archive", create_zip(OFFER_OUTPUT_FOLDER), "offers.zip", "application/zip")
        with post_col2:
            if st.button("📧 Send Offer Emails"):
                send_offer_email(st.session_state.offer_generated)
        
        with st.expander("View Audit Logs"):
            for item in st.session_state.offer_generated:
                filename = os.path.basename(item.get('pdf_path') or item.get('pdf', ''))
                st.markdown(f"**{item['name']}** → `{filename}`")

elif menu == "🎓 Certificates":
    st.header("Certificate Issuance")
    cert_template_status = "Active" if os.path.exists('certificate_template.docx') else "Missing Configuration"
    
    col1, col2, col3 = st.columns(3)
    col1.metric("System Template", cert_template_status)
    col2.metric("Eligible Entries", len([r for r in st.session_state.shared_data if all(str(r.get(c, '')).strip() for c in CERT_REQUIRED_COLUMNS)]) if st.session_state.shared_data else 0)
    col3.metric("Compiled Archive", st.session_state.cert_count)

    st.markdown("---")
    
    act_col_c, _ = st.columns([1, 3])
    with act_col_c:
        if st.button("🎓 Process Certificates", disabled=not st.session_state.shared_data):
            with st.spinner("Processing rendering tasks..."):
                st.session_state.cert_generated = generate_certificate(st.session_state.shared_data)
                st.session_state.cert_count = len(st.session_state.cert_generated)
        
    if len(st.session_state.cert_generated) > 0:
        st.success(f"Processing complete: {st.session_state.cert_count} certificates generated.")
        
        post_col_c1, post_col_c2, post_col_c3 = st.columns([1, 1, 2])
        with post_col_c1:
            st.download_button("📦 Download ZIP Archive", create_zip(CERT_OUTPUT_FOLDER), "certs.zip", "application/zip")
        with post_col_c2:
            if st.button("📧 Send Certificate Emails"):
                send_certificate_email(st.session_state.cert_generated)
        
        with st.expander("View Audit Logs"):
            for item in st.session_state.cert_generated:
                filename = os.path.basename(item.get('pdf_path') or item.get('pdf', ''))
                st.markdown(f"**{item['name']}** → `{filename}`")

elif menu == "📊 Batch Operations":
    st.header("Asset Distribution")
    st.markdown("Review operational readiness and execute outbound distribution.")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Offers Queued", st.session_state.offer_count)
    col2.metric("Certificates Queued", st.session_state.cert_count)
    col3.metric("Total Load", st.session_state.offer_count + st.session_state.cert_count)
    
    all_data = st.session_state.offer_generated + st.session_state.cert_generated
    
    st.markdown("---")
    if all_data:
        st.info(f"System primed for dispatch. Conveying {len(all_data)} outbound items.")
        batch_col, _ = st.columns([1, 3])
        with batch_col:
            if st.button("🚀 Execute Global Dispatch", disabled=not all_data):
                with st.spinner("Transmitting data network..."):
                    sent_offers, sent_certs = 0, 0
                    if st.session_state.offer_generated:
                        results_offer = send_offer_email(st.session_state.offer_generated)
                        sent_offers = results_offer.get('sent', 0)
                    if st.session_state.cert_generated:
                        results_cert = send_certificate_email(st.session_state.cert_generated)
                        sent_certs = results_cert.get('sent', 0)
                    total_sent = sent_offers + sent_certs
                    st.success(f"✅ Operations complete. Successfully fired {total_sent}/{len(all_data)} dispatches.")
                    st.balloons()
    else:
        st.warning("Action restricted: Provide operational assets before firing dispatch execution.")
        
elif menu == "📧 Custom Mailing":
    st.header("📧 Custom Mailing")
    st.markdown("Deploy clean text, multimedia links, and file payloads across your targeted rosters.")
    
    input_source = st.radio("Email Source Parameter", ["Upload Excel", "Manual Input"], horizontal=True)
    valid_recipients = []
    
    if input_source == "Upload Excel":
        bulk_upload = st.file_uploader("Upload Mailing List ('email' column required)", type=['xlsx', 'xls'], key="custom_mail_excel")
        if bulk_upload:
            try:
                df_bulk = pd.read_excel(bulk_upload)
                email_col = next((c for c in df_bulk.columns if str(c).strip().lower() == 'email'), None)
                name_col = next((c for c in df_bulk.columns if str(c).strip().lower() == 'name'), None)
                if not email_col:
                    st.error("❌ Schema mismatch: The provided Excel must contain an 'email' column.")
                else:
                    for row in df_bulk.to_dict('records'):
                        email_val = str(row.get(email_col, '')).strip()
                        if email_val and pd.notna(row.get(email_col)) and '@' in email_val:
                            valid_recipients.append({'email': email_val, 'name': str(row.get(name_col, 'User')).strip() if name_col else 'User'})
            except Exception as e:
                st.error(f"Error parsing Excel block: {e}")
    else:
        manual_emails = st.text_area("Enter absolute emails (comma-separated)", "")
        if manual_emails:
            for email in manual_emails.split(','):
                email = email.strip()
                if email and '@' in email:
                    valid_recipients.append({'email': email, 'name': 'User'})
                    
    with st.container():
        st.subheader("Payload Composer")
        subject = st.text_input("Title Header *", "")
        message = st.text_area("Message Body (Links permitted)", "Hello {name},\n\n\nRegards,\nTeam")
        st.caption("Parameters: map {name} dynamically inside the body string.")
        
        attachments = st.file_uploader("Upload Payload Assets", type=['pdf', 'docx', 'png', 'jpg', 'jpeg', 'mp4', 'txt'], accept_multiple_files=True)
    
    if len(valid_recipients) > 0:
        st.success(f"✅ Indexed {len(valid_recipients)} target leads.")
        
        with st.expander("👁️ Preview Broadcast Payload", expanded=False):
            st.write(f"**Target Volume:** {len(valid_recipients)}")
            st.write(f"**Subject:** {subject if subject else '⚠ (Missing)'}")
            st.write(f"**Files:** {len(attachments)} attached")
            if attachments:
                for a in attachments:
                    st.caption(f"📎 {a.name}")
            st.markdown("---")
            st.write(message.replace('{name}', 'PreviewUser'))
            
        missing_core_params = not subject or (not message.strip() and not attachments)
        
        st.markdown("---")
        total_emails = len(valid_recipients)
        st.write(f"**Queued Pipeline Vol:** {total_emails}")
        
        if total_emails > 50:
            st.warning("⚠️ High Volume Operations: Deploying upwards of 50 payloads. Confirm SMTP server burst limits accommodate heavy outbound traffic.")
            
        auto_test_mode = total_emails <= 1
        is_test = st.checkbox("☑ Send as test (only first email mapped)", value=auto_test_mode)
        
        target_recipients = [valid_recipients[0]] if is_test else valid_recipients
        btn_disabled = missing_core_params
        
        if len(target_recipients) > 1:
            confirm = st.checkbox(f"🚨 You are about to deploy broadcast payloads to {len(target_recipients)} distinct users. Check to authorize transmission.")
            if not confirm:
                btn_disabled = True
                
        act_col, _ = st.columns([1, 2])
        with act_col:
            if st.button("🚀 Send Emails", disabled=btn_disabled, use_container_width=True):
                attachment_paths = []
                temp_dir = "temp_attachments"
                os.makedirs(temp_dir, exist_ok=True)
                if attachments:
                    for attachment in attachments:
                        path = os.path.join(temp_dir, attachment.name)
                        with open(path, "wb") as f:
                            f.write(attachment.getbuffer())
                        attachment_paths.append(path)
                        
                with st.spinner("Executing dispatch sequence..."):
                    results = execute_bulk_custom_email(target_recipients, subject, message, attachment_paths)
                    
                if results:
                    st.success(f"✅ Success. Transmitted {results['sent']}/{len(target_recipients)} payloads.")
                    if results['errors']:
                        st.error(f"Failed Delivery Exceptions: {len(results['errors'])}.")
                        with st.expander("Inspect Failure Nodes"):
                            for e in results['errors']:
                                st.write(f"- {e['email']}: {e['error']}")
                                
                for path in attachment_paths:
                    try: os.remove(path)
                    except: pass

        if missing_core_params:
            st.error("⚠️ Transmission Locked: Subject Field and Payload Data (Text/Attachments) are required parameters.")

# App runs directly without main() call
st.caption("✅ Syntax fixed - Bulk offer letter generator ready!")


