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
# 1. SESSION STATE
# ==========================================
if 'logs' not in st.session_state:
    st.session_state.logs = []
if 'detected_folder' not in st.session_state:
    st.session_state.detected_folder = None

# ==========================================
# 2. PAGE CONFIGURATION
# ==========================================
st.set_page_config(page_title="OutreachMaster Pro", page_icon="🚀", layout="wide")

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
    .instruction-box { 
        background-color: #e3f2fd; padding: 20px; border-radius: 10px; 
        border: 1px solid #bbdefb; color: #0d47a1;
    }
    .status-success {
        padding: 10px; background-color: #d4edda; color: #155724; 
        border-radius: 5px; border: 1px solid #c3e6cb; margin-bottom: 10px;
    }
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

def get_folder_priority_list(provider_name):
    """Returns folder list optimized for the selected provider"""
    defaults = ['Sent', 'Sent Items', 'INBOX.Sent', 'INBOX/Sent', '[Gmail]/Sent Mail']
    
    if provider_name == "Hostinger":
        # Hostinger prefers INBOX.Sent
        return ['INBOX.Sent', 'Sent', 'Sent Items', 'INBOX/Sent']
    elif provider_name == "Gmail":
        return ['[Gmail]/Sent Mail', 'Sent', 'Sent Items']
    elif provider_name == "Outlook":
        return ['Sent Items', 'Sent', 'INBOX.Sent']
    else:
        return defaults

# --- CONNECTION TESTER & FOLDER DETECTOR ---
def test_connection_and_find_folder(conf, user, password, provider_name):
    results = {"smtp": False, "imap": False, "folder": None, "msg": ""}
    
    # 1. Test SMTP
    try:
        if conf['port'] == 465:
            server = smtplib.SMTP_SSL(conf['smtp'], conf['port'])
        else:
            server = smtplib.SMTP(conf['smtp'], conf['port'])
            server.starttls()
        server.login(user, password)
        server.quit()
        results["smtp"] = True
    except Exception as e:
        results["msg"] = f"SMTP Error: {str(e)}"
        return results

    # 2. Test IMAP & Detect Folder
    try:
        mail = imaplib.IMAP4_SSL(conf['imap'], conf['i_port'])
        mail.login(user, password)
        results["imap"] = True
        
        # Get optimized list based on provider
        possible_folders = get_folder_priority_list(provider_name)
        
        for folder in possible_folders:
            try:
                status, _ = mail.select(folder)
                if status == 'OK':
                    results["folder"] = folder
                    break
            except:
                continue
        
        mail.logout()
    except Exception as e:
        results["msg"] = f"IMAP Error: {str(e)}"
    
    return results

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
        return True, "Sent Successfully", msg.as_string()
    except Exception as e:
        return False, str(e), None

def save_sent_folder(conf, user, password, raw_msg, provider_name, specific_folder=None):
    if not raw_msg: return False
    try:
        mail = imaplib.IMAP4_SSL(conf['imap'], conf['i_port'])
        mail.login(user, password)
        
        saved = False
        
        # 1. Try the specifically detected folder first
        if specific_folder:
            try:
                mail.append(specific_folder, '\\Seen', imaplib.Time2Internaldate(time.time()), raw_msg.encode('utf-8'))
                mail.logout()
                return True
            except:
                pass # If failed, fall back to list logic
        
        # 2. Fallback: Try priority list
        folders_to_try = get_folder_priority_list(provider_name)
        for folder in folders_to_try:
            try:
                mail.append(folder, '\\Seen', imaplib.Time2Internaldate(time.time()), raw_msg.encode('utf-8'))
                saved = True
                break
            except:
                continue
        
        mail.logout()
        return saved
    except: 
        return False

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
    
    # --- TEST CONNECTION ---
    if st.button("🔌 Test Connection"):
        if email_user and email_pass:
            with st.spinner("Authenticating & Finding Sent Folder..."):
                # Pass provider name to optimize search
                res = test_connection_and_find_folder(conf, email_user, email_pass, p_choice)
                
                if res["smtp"]:
                    # Store detected folder in session
                    st.session_state.detected_folder = res["folder"]
                    
                    st.markdown(f"""
                    <div class="status-success">
                        <b>✅ Connected Successfully!</b><br>
                        Sending (SMTP): Working<br>
                        Saving (IMAP): Working
                    </div>
                    """, unsafe_allow_html=True)
                    
                    if res["folder"]:
                        st.info(f"📂 Sent Folder Detected: **{res['folder']}**")
                    else:
                        st.warning("⚠️ Connected, but 'Sent' folder not found. Emails will send but not save.")
                else:
                    st.error(f"❌ Connection Failed. {res['msg']}")
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
# 5. MAIN INTERFACE (TOP)
# ==========================================
st.title("🚀 OutreachMaster Pro")
st.markdown("Automated Email Outreach with **Randomized Templates** and **Strict Grouping**.")
st.divider()

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
    st.info("Upload HTML or TXT files. You can use Bulk Upload for multiple files.")
    
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
        st.caption("Select multiple HTML/TXT files at once.")
        bulk_files = st.file_uploader("Upload multiple files", type=['html', 'txt'], accept_multiple_files=True)
        if bulk_files:
            for bf in bulk_files:
                c = io.StringIO(bf.getvalue().decode("utf-8")).read()
                collected_templates.append({"id": f"Bulk-{bf.name}", "content": c})
            st.success(f"✅ {len(bulk_files)} Bulk Templates Added!")

# ==========================================
# 6. LIVE OPERATIONS DASHBOARD
# ==========================================
st.divider()
st.subheader("📊 Campaign Dashboard (Live Operations)")

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
        
        # Get detected folder from session (set during Test Connection)
        target_folder = st.session_state.get('detected_folder', None)
        
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
            
            # --- STRICT VARIABLE LOGIC ---
            comp_raw = str(row.get('Company', '')).strip()
            web_raw = str(row.get('Website', '')).strip()
            
            if comp_raw.lower() == 'nan': comp_raw = ""
            if web_raw.lower() == 'nan': web_raw = ""
            
            for em in raw_emails:
                clean_em = em.strip()
                if "@" in clean_em:
                    g_key = comp_raw if comp_raw else get_technical_domain(clean_em)
                    
                    all_leads.append({
                        "Group": g_key, 
                        "D_Comp": comp_raw,
                        "D_Web": web_raw,
                        "Email": clean_em, 
                        "Name": name_val, 
                        "Row": row
                    })

        grouped_leads = defaultdict(list)
        for lead in all_leads: grouped_leads[lead['Group']].append(lead)
        
        total_groups = len(grouped_leads)
        sent_counter = 0
        curr_progress = 0
        start_sr_no = len(st.session_state.logs) + 1
        
        # --- SENDING LOOP ---
        for group_key, contacts in grouped_leads.items():
            if sent_counter >= daily_limit:
                st.warning("⚠️ Daily Limit Reached. Process Stopped."); break
            
            # Pick Random Template for this Company Group
            curr_tpl = random.choice(collected_templates)
            
            # UI Status Update
            ui_label = contacts[0]['D_Comp'] if contacts[0]['D_Comp'] else group_key
            
            status_box.markdown(f"""
                <div class="live-card">
                    <h4>🔄 Processing: {ui_label}</h4>
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
                success, response_msg, raw_data = send_email_smtp(conf, email_user, email_pass, person['Email'], person['Name'], sender_display, sub, bod)
                
                status_txt = "✅ Sent" if success else "❌ Failed"
                
                # Save to Sent Folder Logic
                if success and raw_data:
                    # Pass provider choice AND detected folder
                    is_saved = save_sent_folder(conf, email_user, email_pass, raw_data, p_choice, specific_folder=target_folder)
                    
                    if not is_saved:
                        status_txt += " (Not Saved)"
                    sent_counter += 1
                
                # Current Time
                now_time = datetime.now().strftime("%I:%M:%S %p")
                
                # Update Logs
                log_entry = {
                    "Sr. No": start_sr_no,
                    "Time": now_time,
                    "Company": person['D_Comp'], 
                    "Email": person['Email'],
                    "Status": status_txt,
                    "Template": curr_tpl['id'],
                    "Subject": sub,
                    "Error Info": response_msg if not success else ""
                }
                st.session_state.logs.append(log_entry)
                start_sr_no += 1
                
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
# 9. USER GUIDE (BOTTOM)
# ==========================================
st.divider()
with st.expander("📘 User Guide & Instructions (Read First)", expanded=True):
    st.markdown("""
    <div class="instruction-box">
    <h4>🚀 How to use OutreachMaster?</h4>
    <ol>
        <li><strong>Credentials Setup:</strong> Select Provider (e.g. Hostinger), enter Email/Pass, and click <b>'Test Connection'</b>.
            <ul><li>The system will tell you if Login worked AND which <b>Sent Folder</b> (e.g., INBOX.Sent) it detected.</li></ul>
        </li>
        <li><strong>Prepare Recipients:</strong> Upload your Excel/CSV file. 
            <ul>
                <li>Required Column: <code>Email</code></li>
                <li>Optional Columns: <code>Name</code>, <code>Company</code>, <code>Website</code></li>
            </ul>
        </li>
        <li><strong>Using Dynamic Tags:</strong>
            <ul>
                <li>Use <code>{Name}</code>, <code>{Company}</code>, <code>{Website}</code> in Subject/Body.</li>
                <li>Example Subject: <em>"Question for {Company}"</em></li>
            </ul>
        </li>
        <li><strong>Global Subjects:</strong> Paste multiple subject lines. System picks randomly.</li>
        <li><strong>Templates:</strong> Upload HTML/TXT files via Manual or Bulk tabs.</li>
        <li><strong>Launch:</strong> Click 'Start Campaign'.</li>
    </ol>
    </div>
    """, unsafe_allow_html=True)
