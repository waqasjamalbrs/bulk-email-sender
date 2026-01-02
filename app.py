import streamlit as st
import pandas as pd
import smtplib
import imaplib
import time
import random
import io
import itertools
from collections import defaultdict
from datetime import datetime
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header

# ==========================================
# 1. SESSION STATE MANAGEMENT
# ==========================================
if 'logs' not in st.session_state:
    st.session_state.logs = []

# ==========================================
# 2. PAGE CONFIGURATION
# ==========================================
st.set_page_config(page_title="OutreachMaster Pro", page_icon="🚀", layout="wide")

# Custom CSS for Professional Look
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { 
        border-radius: 6px; font-weight: 600; height: 48px; width: 100%;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .live-card {
        padding: 15px; border-radius: 8px; background: white; 
        border-left: 5px solid #007bff; box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    .success-tag { color: #28a745; font-weight: bold; }
    .error-tag { color: #dc3545; font-weight: bold; }
    .instruction-box { background-color: #e9ecef; padding: 15px; border-radius: 8px; font-size: 14px; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. HELPER FUNCTIONS
# ==========================================
PROVIDERS = {
    "Hostinger": {"smtp": "smtp.hostinger.com", "port": 465, "imap": "imap.hostinger.com", "i_port": 993},
    "Gmail": {"smtp": "smtp.gmail.com", "port": 465, "imap": "imap.gmail.com", "i_port": 993},
    "Outlook": {"smtp": "smtp.office365.com", "port": 587, "imap": "outlook.office365.com", "i_port": 993},
    "Custom": {"smtp": "", "port": 465, "imap": "", "i_port": 993}
}

def test_connection(conf, user, password):
    try:
        if conf['port'] == 465:
            server = smtplib.SMTP_SSL(conf['smtp'], conf['port'])
        else:
            server = smtplib.SMTP(conf['smtp'], conf['port'])
            server.starttls()
        server.login(user, password)
        server.quit()
        return True, "Connection Successful! Credentials are valid."
    except Exception as e:
        return False, f"Connection Failed: {str(e)}"

def get_technical_domain(email):
    try:
        if "@" in email: return email.split('@')[1].lower().strip()
    except: pass
    return "unknown"

def send_email_smtp(conf, user, password, to_email, to_name, sender_name, subject, body):
    msg = MIMEMultipart()
    msg['From'] = formataddr((str(Header(sender_name, 'utf-8')), user))
    msg['To'] = formataddr((str(Header(to_name, 'utf-8')), to_email))
    msg['Subject'] = Header(subject, 'utf-8')
    msg.attach(MIMEText(body, 'html', 'utf-8'))

    try:
        if conf['port'] == 465:
            server = smtplib.SMTP_SSL(conf['smtp'], conf['port'])
        else:
            server = smtplib.SMTP(conf['smtp'], conf['port'])
            server.starttls()
        server.login(user, password)
        server.sendmail(user, [to_email], msg.as_string())
        server.quit()
        return True, "Sent Successfully"
    except Exception as e:
        return False, str(e)

def save_sent_folder(conf, user, password, raw_msg):
    try:
        mail = imaplib.IMAP4_SSL(conf['imap'], conf['i_port'])
        mail.login(user, password)
        mail.append("INBOX.Sent", '\\Seen', imaplib.Time2Internaldate(time.time()), raw_msg.encode('utf-8'))
        mail.logout()
    except: pass

# ==========================================
# 4. SIDEBAR CONFIGURATION
# ==========================================
with st.sidebar:
    st.title("⚙️ Configuration")
    
    st.subheader("1. SMTP Settings")
    p_choice = st.selectbox("Email Provider", list(PROVIDERS.keys()))
    conf = PROVIDERS[p_choice]
    
    if p_choice == "Custom":
        conf['smtp'] = st.text_input("SMTP Host")
        conf['port'] = st.number_input("SMTP Port", 465)
        
    email_user = st.text_input("Email Address")
    email_pass = st.text_input("Password / App Password", type="password")
    sender_display = st.text_input("Sender Name (Display)", "Joseph Miller")
    
    if st.button("🔌 Test Connection"):
        if email_user and email_pass:
            with st.spinner("Verifying credentials..."):
                success, msg = test_connection(conf, email_user, email_pass)
                if success: st.success(msg)
                else: st.error(msg)
        else:
            st.warning("Please enter Email and Password first.")

    st.divider()
    st.subheader("2. Safety & Limits")
    daily_limit = st.number_input("Stop after sending X emails", 50, 5000, 100)
    delay_min = st.number_input("Min Delay (Seconds)", 5)
    delay_max = st.number_input("Max Delay (Seconds)", 20)
    
    st.divider()
    if st.button("🗑️ Clear History & Reset", type="primary"):
        st.session_state.logs = []
        st.rerun()

# ==========================================
# 5. MAIN INTERFACE
# ==========================================
st.title("🚀 OutreachMaster Pro")
st.markdown("Automated Email Outreach with **Randomized Templates** and **Company Grouping**.")

col1, col2 = st.columns([1, 1.2])

# --- LEFT COLUMN: INPUTS ---
with col1:
    st.subheader("📂 Step 1: Recipients")
    uploaded_file = st.file_uploader("Upload Excel/CSV File", type=['xlsx', 'csv'])
    if uploaded_file:
        st.success("✅ File Loaded Successfully")

    st.divider()
    
    st.subheader("📝 Step 2: Subject Lines")
    st.caption("Paste ALL subject lines below (One per line). System will pick randomly.")
    subjects_input = st.text_area("Global Subject List", height=200, 
                                  placeholder="Collaboration for {Website}\nQuick question regarding {Company}\nFeature inquiry for {Name}",
                                  help="Supported Tags: {Name}, {Company}, {Website}")

# --- RIGHT COLUMN: TEMPLATES ---
with col2:
    st.subheader("📄 Step 3: Body Templates")
    st.info("You can use Manual Tabs OR Bulk Upload (or both). System will merge them.")
    
    tab_manual, tab_bulk = st.tabs(["Manual Upload (Single)", "Bulk Upload (Multiple)"])
    
    collected_templates = []

    # A. MANUAL TABS
    with tab_manual:
        st.caption("Upload individual files for specific slots.")
        sub_tabs = st.tabs(["T1", "T2", "T3", "T4", "T5"])
        for i, t in enumerate(sub_tabs):
            with t:
                f = st.file_uploader(f"Body HTML/TXT {i+1}", type=['html', 'txt'], key=f"manual_{i}")
                if f:
                    c = io.StringIO(f.getvalue().decode("utf-8")).read()
                    collected_templates.append({"id": f"Manual-T{i+1}", "content": c})
                    st.caption(f"✅ Loaded: {len(c)} chars")

    # B. BULK UPLOAD (OPTIONAL)
    with tab_bulk:
        st.caption("Optional: Select multiple files at once.")
        bulk_files = st.file_uploader("Upload multiple HTML/TXT files", type=['html', 'txt'], accept_multiple_files=True)
        if bulk_files:
            for bf in bulk_files:
                # Read content (TXT or HTML)
                c = io.StringIO(bf.getvalue().decode("utf-8")).read()
                collected_templates.append({"id": f"Bulk-{bf.name}", "content": c})
            st.success(f"✅ {len(bulk_files)} Bulk Templates Added!")

# ==========================================
# 6. LIVE OPERATIONS DASHBOARD
# ==========================================
st.divider()
st.subheader("📊 Live Operations")

# Real-time placeholders
status_box = st.empty()
progress_bar = st.progress(0)
log_table = st.empty()

# Display Persistent History
if st.session_state.logs:
    log_table.dataframe(pd.DataFrame(st.session_state.logs), use_container_width=True)

# ==========================================
# 7. EXECUTION LOGIC
# ==========================================
if st.button("🚀 START CAMPAIGN"):
    # --- VALIDATION ---
    if not uploaded_file or not email_user or not email_pass:
        st.error("❌ Critical Error: Missing File or Credentials.")
    elif not subjects_input.strip():
        st.error("❌ Critical Error: Please add at least one Subject Line.")
    elif not collected_templates:
        st.error("❌ Critical Error: No Body Templates found. Please upload via Manual or Bulk tabs.")
    else:
        # Prepare Subject List
        subject_pool = [s.strip() for s in subjects_input.split('\n') if s.strip()]
        
        # --- READ DATA ---
        try:
            if uploaded_file.name.endswith('.csv'): df = pd.read_csv(uploaded_file)
            else: df = pd.read_excel(uploaded_file)
        except:
            st.error("Error reading recipient file.")
            st.stop()

        # --- PRE-PROCESS & GROUPING ---
        all_leads = []
        for _, row in df.iterrows():
            raw_emails = str(row.get('Email', '')).split(',')
            name_val = str(row.get('Name', 'there')).strip()
            
            # Strict Variable Logic
            comp = str(row.get('Company', '')).strip()
            web = str(row.get('Website', '')).strip()
            if comp.lower() == 'nan': comp = ""
            if web.lower() == 'nan': web = ""
            
            # Display Logic (Fallback to 'your company' if empty)
            d_comp = comp if comp else (web if web else "your company")
            d_web = web if web else (comp if comp else "your website")
            
            for em in raw_emails:
                clean_em = em.strip()
                if "@" in clean_em:
                    # Grouping Key
                    g_key = comp if comp else get_technical_domain(clean_em)
                    all_leads.append({
                        "Group": g_key, "D_Comp": d_comp, "D_Web": d_web,
                        "Email": clean_em, "Name": name_val, "Row": row
                    })

        grouped_leads = defaultdict(list)
        for lead in all_leads: grouped_leads[lead['Group']].append(lead)
        
        total_groups = len(grouped_leads)
        sent_counter = 0
        curr_progress = 0
        
        # --- SENDING LOOP ---
        for group_key, contacts in grouped_leads.items():
            if sent_counter >= daily_limit:
                st.warning("⚠️ Daily Limit Reached. Process Stopped."); break
            
            # Pick Random Template for this Company Group
            curr_tpl = random.choice(collected_templates)
            
            # UI Update
            display_name = contacts[0]['D_Comp'] if contacts[0]['D_Comp'] != "your company" else group_key
            status_box.markdown(f"""
                <div class="live-card">
                    <h4>🔄 Processing: {display_name}</h4>
                    <span><b>Emails:</b> {len(contacts)}</span> | 
                    <span><b>Template:</b> {curr_tpl['id']}</span>
                </div>
            """, unsafe_allow_html=True)
            
            for person in contacts:
                # Pick Random Subject
                curr_subj = random.choice(subject_pool)
                
                # Replace Variables
                sub = curr_subj.replace("{Name}", person['Name']).replace("{Company}", person['D_Comp']).replace("{Website}", person['D_Web'])
                bod = curr_tpl['content'].replace("{Name}", person['Name']).replace("{Company}", person['D_Comp']).replace("{Website}", person['D_Web'])
                
                # Replace Extra Excel Columns
                for col in df.columns:
                    val = str(person['Row'].get(col, ''))
                    sub = sub.replace(f"{{{col}}}", val)
                    bod = bod.replace(f"{{{col}}}", val)
                
                # Send Email
                success, response_msg = send_email_smtp(conf, email_user, email_pass, person['Email'], person['Name'], sender_display, sub, bod)
                
                status_txt = "✅ Sent" if success else "❌ Failed"
                if success:
                    save_sent_folder(conf, email_user, email_pass, response_msg)
                    sent_counter += 1
                
                # Update Logs
                log_entry = {
                    "Time": datetime.now().strftime("%H:%M:%S"),
                    "Company": display_name,
                    "Email": person['Email'],
                    "Status": status_txt,
                    "Template": curr_tpl['id'],
                    "Subject": sub,
                    "Error Info": response_msg if not success else ""
                }
                st.session_state.logs.append(log_entry)
                
                # Refresh Table
                log_table.dataframe(pd.DataFrame(st.session_state.logs), use_container_width=True)
                
                # Internal Delay
                time.sleep(2)

            # External Delay & Progress
            curr_progress += 1
            progress_bar.progress(curr_progress / total_groups)
            time.sleep(random.randint(delay_min, delay_max))

        st.success("🎉 Campaign Completed Successfully!")

# ==========================================
# 8. EXPORT SECTION
# ==========================================
st.divider()
if st.session_state.logs:
    final_df = pd.DataFrame(st.session_state.logs)
    col_a, col_b = st.columns([1, 1])
    
    with col_a:
        st.metric("Total Processed", len(final_df))
    with col_b:
        success_count = len(final_df[final_df['Status'].str.contains("Sent")])
        st.metric("Successful Deliveries", success_count)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False)
    
    st.download_button("📥 Download Final Report", buffer, "Campaign_Report.xlsx", type="primary")

# ==========================================
# 9. USER GUIDE (INSTRUCTIONS)
# ==========================================
with st.expander("📘 User Guide & Documentation", expanded=False):
    st.markdown("""
    <div class="instruction-box">
    <h4>Getting Started</h4>
    <ol>
        <li><strong>Credentials:</strong> Enter your email provider details in the sidebar and click 'Test Connection'.</li>
        <li><strong>Upload Recipients:</strong> Upload your Excel/CSV file. Ensure columns <code>Name</code> and <code>Email</code> exist.</li>
        <li><strong>Subject Lines:</strong> Paste multiple subject lines in the 'Global Subject List'. The system will pick randomly for each email.</li>
        <li><strong>Templates:</strong>
            <ul>
                <li><b>Manual:</b> Upload specific files in Tabs T1-T5.</li>
                <li><b>Bulk (Optional):</b> Use the 'Bulk Upload' tab to select 10, 20, or more <code>.html</code>/<code>.txt</code> files at once.</li>
                <li>The system will combine all uploaded templates and pick randomly.</li>
            </ul>
        </li>
        <li><strong>Run:</strong> Click Start. Do not close the browser tab until the process finishes.</li>
    </ol>
    <p><strong>Note:</strong> You can open this tool in multiple browser tabs to run parallel campaigns.</p>
    </div>
    """, unsafe_allow_html=True)
