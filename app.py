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
# 1. SESSION STATE (Data Save Rakhne k liye)
# ==========================================
if 'logs' not in st.session_state:
    st.session_state.logs = []
if 'campaign_active' not in st.session_state:
    st.session_state.campaign_active = False

# ==========================================
# 2. PAGE CONFIG
# ==========================================
st.set_page_config(page_title="Multi-Campaign Emailer", page_icon="🚀", layout="wide")
st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    .stButton>button { 
        border-radius: 8px; font-weight: bold; height: 45px; width: 100%;
    }
    .live-stat {
        padding: 15px; border-radius: 8px; background: white; 
        border-left: 5px solid #2196F3; box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    .subject-box textarea {
        background-color: #e3f2fd; border: 1px solid #90caf9;
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

def test_smtp_connection(conf, user, password):
    try:
        if conf['port'] == 465:
            server = smtplib.SMTP_SSL(conf['smtp'], conf['port'])
        else:
            server = smtplib.SMTP(conf['smtp'], conf['port'])
            server.starttls()
        server.login(user, password)
        server.quit()
        return True, "Login Successful! ✅"
    except Exception as e:
        return False, str(e)

def get_technical_domain(email):
    try:
        if "@" in email:
            return email.split('@')[1].lower().strip()
    except:
        pass
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
        return True, "OK"
    except Exception as e:
        return False, str(e)

def save_sent(conf, user, password, raw_msg):
    try:
        mail = imaplib.IMAP4_SSL(conf['imap'], conf['i_port'])
        mail.login(user, password)
        mail.append("INBOX.Sent", '\\Seen', imaplib.Time2Internaldate(time.time()), raw_msg.encode('utf-8'))
        mail.logout()
    except:
        pass

# ==========================================
# 4. SIDEBAR
# ==========================================
with st.sidebar:
    st.header("⚙️ Settings")
    
    # Credentials
    p = st.selectbox("Provider", list(PROVIDERS.keys()))
    conf = PROVIDERS[p]
    if p == "Custom":
        conf['smtp'] = st.text_input("SMTP Host")
        conf['port'] = st.number_input("SMTP Port", 465)
        
    e_user = st.text_input("Email Address")
    e_pass = st.text_input("Password", type="password")
    sender_name = st.text_input("Sender Name", "Joseph Miller")
    
    if st.button("🔌 Test Login"):
        if e_user and e_pass:
            ok, msg = test_smtp_connection(conf, e_user, e_pass)
            if ok: st.success(msg)
            else: st.error(msg)
        else:
            st.warning("Fill credentials first.")

    st.divider()
    limit = st.number_input("Daily Limit", 50, 5000, 100)
    d_min = st.number_input("Min Delay", 5)
    d_max = st.number_input("Max Delay", 20)
    
    st.divider()
    if st.button("🗑️ Clear All History", type="primary"):
        st.session_state.logs = []
        st.rerun()

# ==========================================
# 5. MAIN UI
# ==========================================
st.title("🚀 Smart SEO Outreach (Global Subjects)")

col1, col2 = st.columns([1, 1.5])

with col1:
    st.subheader("1. Recipients Data")
    f = st.file_uploader("Upload Excel/CSV", type=['xlsx', 'csv'])
    if f:
        st.success("File Loaded!")

    st.divider()
    st.subheader("2. Global Subject Lines")
    st.caption("Enter ALL subject lines here (One per line). System will pick randomly.")
    global_subjects_raw = st.text_area("Subject Lines", height=200, placeholder="Collaboration for {Website}\nFeature opportunity for {Company}\nQuestion regarding {Name}", key="glob_subs")

with col2:
    st.subheader("3. Email Body Templates (HTML)")
    st.caption("Upload HTML files. Code will rotate these templates per company.")
    
    tabs = st.tabs(["Template 1", "Template 2", "Template 3", "Template 4", "Template 5"])
    body_templates = []
    
    for i, tab in enumerate(tabs):
        with tab:
            b_file = st.file_uploader(f"Upload HTML for Body {i+1}", type=['html', 'txt'], key=f"b{i}")
            if b_file:
                b_content = io.StringIO(b_file.getvalue().decode("utf-8")).read()
                st.caption(f"✅ Loaded {len(b_content)} chars")
                body_templates.append({"id": f"Body-T{i+1}", "content": b_content})

# ==========================================
# 6. REAL-TIME DASHBOARD
# ==========================================
st.divider()
st.subheader("📊 Live Campaign Status")

status_placeholder = st.empty()
progress_bar = st.progress(0)
table_placeholder = st.empty()

# Show history if exists
if st.session_state.logs:
    hist_df = pd.DataFrame(st.session_state.logs)
    table_placeholder.dataframe(hist_df, use_container_width=True)

# ==========================================
# 7. EXECUTION LOGIC
# ==========================================
if st.button("🚀 START CAMPAIGN NOW"):
    # Validation
    if not f or not e_user or not e_pass:
        st.error("Missing File or Credentials.")
    elif not global_subjects_raw.strip():
        st.error("Please enter at least one Subject Line.")
    elif not body_templates:
        st.error("Please upload at least one HTML Body Template.")
    else:
        # Prepare Subjects List
        subject_list = [s.strip() for s in global_subjects_raw.split('\n') if s.strip()]
        
        # --- PREPARE DATA ---
        try:
            if f.name.endswith('.csv'): df = pd.read_csv(f)
            else: df = pd.read_excel(f)
        except:
            st.error("Error reading file.")
            st.stop()
            
        all_contacts = []
        for _, row in df.iterrows():
            raw_emails = str(row.get('Email', '')).split(',')
            name_val = str(row.get('Name', 'there')).strip()
            
            # Company/Website Logic (Strict)
            comp_col = str(row.get('Company', '')).strip()
            web_col = str(row.get('Website', '')).strip()
            
            if comp_col.lower() == 'nan': comp_col = ""
            if web_col.lower() == 'nan': web_col = ""
            
            d_comp = comp_col if comp_col else (web_col if web_col else "your company")
            d_web = web_col if web_col else (comp_col if comp_col else "your website")
            
            for em in raw_emails:
                clean_em = em.strip()
                if "@" in clean_em:
                    # Grouping Key
                    grp = comp_col if comp_col else get_technical_domain(clean_em)
                    all_contacts.append({
                        "Group": grp, "D_Comp": d_comp, "D_Web": d_web,
                        "Email": clean_em, "Name": name_val, "Row": row
                    })
        
        grouped = defaultdict(list)
        for c in all_contacts: grouped[c['Group']].append(c)
        
        total_groups = len(grouped)
        tpl_cycle = itertools.cycle(body_templates) # Rotate Bodies
        curr_idx = 0
        campaign_sent = 0
        
        # --- SENDING LOOP ---
        for grp_key, contacts in grouped.items():
            if campaign_sent >= limit:
                st.warning("Daily limit reached!"); break
                
            curr_body = next(tpl_cycle) # Next HTML Template
            
            # UI Update
            display_name = contacts[0]['D_Comp'] if contacts[0]['D_Comp'] != "your company" else grp_key
            status_placeholder.markdown(f"""
                <div class="live-stat">
                    <h4>🔄 Processing: {display_name}</h4>
                    <p>Emails: {len(contacts)} | Body: {curr_body['id']}</p>
                </div>
            """, unsafe_allow_html=True)
            
            for p in contacts:
                # Pick Random Subject from GLOBAL list
                curr_subj = random.choice(subject_list)
                
                # Variables Replacement
                sub = curr_subj.replace("{Name}", p['Name']).replace("{Company}", p['D_Comp']).replace("{Website}", p['D_Web'])
                bod = curr_body['content'].replace("{Name}", p['Name']).replace("{Company}", p['D_Comp']).replace("{Website}", p['D_Web'])
                
                # Extra Columns
                for k in df.columns:
                    val = str(p['Row'].get(k, ''))
                    sub = sub.replace(f"{{{k}}}", val)
                    bod = bod.replace(f"{{{k}}}", val)
                
                # Send
                ok, msg = send_email_smtp(conf, e_user, e_pass, p['Email'], p['Name'], sender_name, sub, bod)
                
                status = "✅ Sent" if ok else "❌ Failed"
                if ok: 
                    save_sent(conf, e_user, e_pass, msg)
                    campaign_sent += 1
                
                # --- UPDATE REAL-TIME LOGS ---
                new_log = {
                    "Time": datetime.now().strftime("%H:%M:%S"),
                    "Company": display_name,
                    "Email": p['Email'],
                    "Status": status,
                    "Template": curr_body['id'],
                    "Subject Used": sub,
                    "Error": msg if not ok else ""
                }
                
                st.session_state.logs.append(new_log)
                
                # Refresh Table
                current_df = pd.DataFrame(st.session_state.logs)
                table_placeholder.dataframe(current_df, use_container_width=True)
                
                time.sleep(2) # Gap between emails in same company

            curr_idx += 1
            progress_bar.progress(curr_idx / total_groups)
            time.sleep(random.randint(d_min, d_max))

        st.success("Campaign Complete!")

# ==========================================
# 8. DOWNLOAD BUTTON
# ==========================================
st.divider()
if st.session_state.logs:
    final_df = pd.DataFrame(st.session_state.logs)
    
    # Simple Stats
    total = len(final_df)
    passed = len(final_df[final_df['Status'].str.contains("Sent")])
    
    st.metric("Session Summary", f"{total} Processed", f"{passed} Delivered")
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False)
    
    st.download_button("📥 Download History", buffer, "Campaign_History.xlsx")

# ==========================================
# 9. USER GUIDE
# ==========================================
with st.expander("📖 Guide: How to use?"):
    st.markdown("""
    1. **Global Subject Lines:** Section 2 mai saari subject lines paste karein. System har email k liye inme se koi ek random uthayega.
    2. **Body Templates:** Section 3 mai Tabs khol kar HTML files upload karein.
    3. **Multiple Campaigns:** Naye tab mai website khol kar dusri campaign chalayen.
    4. **Persistence:** Data tab tak save rahega jab tak aap "Clear All History" nahi dabate.
    """)
