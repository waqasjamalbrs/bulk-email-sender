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
# 1. PAGE SETUP
# ==========================================
st.set_page_config(page_title="SEO Outreach Pro", page_icon="📨", layout="wide")
st.markdown("""
    <style>
    .main { background-color: #f7f9fc; }
    .stButton>button { 
        background: linear-gradient(90deg, #1d976c, #93f9b9); 
        color: #004d2e; border: none; height: 50px; font-size: 18px; font-weight: bold;
        border-radius: 8px; width: 100%;
    }
    .uploaded-file-box { border: 2px dashed #4CAF50; padding: 10px; border-radius: 5px; background-color: #e8f5e9; }
    .instruction-box { background-color: #e3f2fd; padding: 15px; border-radius: 8px; border-left: 5px solid #2196f3; }
    .success-box { padding: 10px; background-color: #d4edda; color: #155724; border-radius: 5px; }
    .error-box { padding: 10px; background-color: #f8d7da; color: #721c24; border-radius: 5px; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. CORE FUNCTIONS
# ==========================================
PROVIDERS = {
    "Hostinger": {"smtp": "smtp.hostinger.com", "port": 465, "imap": "imap.hostinger.com", "i_port": 993},
    "Gmail": {"smtp": "smtp.gmail.com", "port": 465, "imap": "imap.gmail.com", "i_port": 993},
    "Outlook": {"smtp": "smtp.office365.com", "port": 587, "imap": "outlook.office365.com", "i_port": 993},
    "Custom": {"smtp": "", "port": 465, "imap": "", "i_port": 993}
}

def test_smtp_connection(conf, user, password):
    """Test login credentials without sending email"""
    try:
        if conf['port'] == 465:
            server = smtplib.SMTP_SSL(conf['smtp'], conf['port'])
        else:
            server = smtplib.SMTP(conf['smtp'], conf['port'])
            server.starttls()
        server.login(user, password)
        server.quit()
        return True, "Connection Successful! ✅"
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
# 3. SIDEBAR SETTINGS
# ==========================================
with st.sidebar:
    st.title("⚙️ Configuration")
    
    st.subheader("1. Email Service")
    p = st.selectbox("Select Provider", list(PROVIDERS.keys()))
    conf = PROVIDERS[p]
    if p == "Custom":
        conf['smtp'] = st.text_input("SMTP Host")
        conf['port'] = st.number_input("SMTP Port", 465)
    
    st.divider()
    st.subheader("2. Your Credentials")
    e_user = st.text_input("Your Email Address")
    e_pass = st.text_input("Password / App Password", type="password")
    sender_name = st.text_input("Sender Name (Display)", "Joseph Miller")
    
    # --- TEST CONNECTION BUTTON ---
    if st.button("🔌 Test Connection"):
        if e_user and e_pass:
            with st.spinner("Testing login..."):
                ok, msg = test_smtp_connection(conf, e_user, e_pass)
                if ok:
                    st.success(msg)
                else:
                    st.error(f"Failed: {msg}")
        else:
            st.warning("Enter Email & Password first.")
    
    st.divider()
    st.subheader("3. Safety Limits")
    limit = st.number_input("Stop after sending X emails", 50, 5000, 100)
    d_min = st.number_input("Min Delay (Seconds)", 5)
    d_max = st.number_input("Max Delay (Seconds)", 20)

# ==========================================
# 4. MAIN INTERFACE
# ==========================================
st.title("📨 Ultimate SEO Outreach Tool")
st.markdown("Supports: **HTML File Upload**, **Company Grouping**, & **Strict Variables**.")

col1, col2 = st.columns([1, 2])

# --- LEFT COLUMN: DATA UPLOAD ---
with col1:
    st.subheader("📂 Step 1: Upload Recipients Data")
    f = st.file_uploader("Upload Excel/CSV File", type=['xlsx', 'csv'])
    
    if f:
        st.success("File Uploaded Successfully!")
        try:
            if f.name.endswith('.csv'): 
                preview_df = pd.read_csv(f)
            else: 
                preview_df = pd.read_excel(f)
            st.dataframe(preview_df.head(3), hide_index=True)
            st.caption(f"Total Rows Found: {len(preview_df)}")
        except:
            pass

# --- RIGHT COLUMN: TEMPLATES ---
with col2:
    st.subheader("📝 Step 2: Templates Strategy")
    st.caption("Upload HTML files for body. System will rotate templates per company.")
    
    tabs = st.tabs(["Template 1", "Template 2", "Template 3", "Template 4", "Template 5"])
    tpls = []
    
    for i, tab in enumerate(tabs):
        with tab:
            st.markdown(f"**Template {i+1} Setup**")
            
            # Subject Input (Manual)
            raw_s = st.text_area(
                f"Subject Lines for T{i+1} (One per line)", 
                height=100, 
                key=f"s_{i}",
                placeholder="Collaboration for {Website}\nQuick question regarding {Company}"
            )
            
            # Body Input (FILE UPLOAD)
            st.markdown("**Email Body HTML:**")
            body_file = st.file_uploader(f"Upload .txt or .html for Template {i+1}", type=['html', 'txt'], key=f"f_{i}")
            
            final_body_content = ""
            
            # Read uploaded file content
            if body_file is not None:
                stringio = io.StringIO(body_file.getvalue().decode("utf-8"))
                final_body_content = stringio.read()
                st.markdown(f"<div class='uploaded-file-box'>✅ Loaded: {len(final_body_content)} characters</div>", unsafe_allow_html=True)
                with st.expander("👁️ Preview HTML Content"):
                    st.code(final_body_content, language='html')

            # Store Valid Templates
            if raw_s and final_body_content:
                s_list = [x.strip() for x in raw_s.split('\n') if x.strip()]
                tpls.append({"id": f"Template {i+1}", "subjects": s_list, "body": final_body_content})

# ==========================================
# 5. LOGIC & EXECUTION
# ==========================================
st.divider()

if st.button("🚀 START CAMPAIGN"):
    # --- VALIDATION ---
    if not f:
        st.error("❌ Please upload the Recipients Data file first.")
    elif not e_user or not e_pass:
        st.error("❌ Please enter Email and Password in Sidebar.")
    elif not tpls:
        st.error("❌ Please upload at least one Template Body file and add Subject lines.")
    else:
        # --- 1. READ FILE ---
        try:
            if f.name.endswith('.csv'): df = pd.read_csv(f)
            else: df = pd.read_excel(f)
        except Exception as e:
            st.error(f"Error reading recipients file: {e}")
            st.stop()
            
        # --- 2. PRE-PROCESS DATA ---
        all_contacts = []
        
        for _, row in df.iterrows():
            # Handle multiple emails
            raw_emails = str(row.get('Email', '')).split(',')
            
            # Variables
            name_val = str(row.get('Name', 'there')).strip()
            
            # Fetch Company/Website Columns
            company_col = str(row.get('Company', '')).strip()
            website_col = str(row.get('Website', '')).strip()
            
            # Clean 'nan'
            if company_col.lower() == 'nan': company_col = ""
            if website_col.lower() == 'nan': website_col = ""
            
            # Display Logic
            final_comp_txt = company_col if company_col else (website_col if website_col else "your company")
            final_web_txt = website_col if website_col else (company_col if company_col else "your website")
            
            for em in raw_emails:
                clean_em = em.strip()
                if "@" in clean_em:
                    # Grouping Logic
                    tech_domain = get_technical_domain(clean_em)
                    group_key = company_col if company_col else tech_domain
                    
                    all_contacts.append({
                        "GroupKey": group_key,
                        "DisplayComp": final_comp_txt,
                        "DisplayWeb": final_web_txt,
                        "Email": clean_em,
                        "Name": name_val,
                        "RowData": row
                    })
        
        # --- 3. GROUPING ---
        grouped_data = defaultdict(list)
        for c in all_contacts:
            grouped_data[c['GroupKey']].append(c)
        
        st.success(f"✅ Loaded {len(all_contacts)} emails. Grouped into {len(grouped_data)} unique Companies.")
        
        # --- 4. SENDING LOOP ---
        prog_bar = st.progress(0)
        status_text = st.empty()
        
        report_rows = []
        sent_total = 0
        cycle = itertools.cycle(tpls)
        curr_idx = 0
        total_groups = len(grouped_data)
        
        for g_key, contacts in grouped_data.items():
            if sent_total >= limit: 
                st.warning("🛑 Daily Limit Reached!"); break
            
            curr_tpl = next(cycle)
            group_results = []
            
            display_grp_name = contacts[0]['DisplayComp'] if contacts[0]['DisplayComp'] != "your company" else g_key
            status_text.markdown(f"**Processing:** `{display_grp_name}` | using {curr_tpl['id']}")
            
            for person in contacts:
                email = person['Email']
                name = person['Name']
                
                # Random Subject
                if curr_tpl['subjects']: 
                    subj_base = random.choice(curr_tpl['subjects'])
                else: 
                    subj_base = "Hello"
                
                # Personalization
                sub = subj_base.replace("{Name}", name).replace("{Company}", person['DisplayComp']).replace("{Website}", person['DisplayWeb'])
                bod = curr_tpl['body'].replace("{Name}", name).replace("{Company}", person['DisplayComp']).replace("{Website}", person['DisplayWeb'])
                
                # Extra Columns
                for col in df.columns:
                     val = str(person['RowData'].get(col, ''))
                     sub = sub.replace(f"{{{col}}}", val)
                     bod = bod.replace(f"{{{col}}}", val)
                
                # Send
                ok, msg = send_email_smtp(conf, e_user, e_pass, email, name, sender_name, sub, bod)
                
                if ok:
                    save_sent(conf, e_user, e_pass, msg)
                    sent_total += 1
                    group_results.append(f"✅ {email}")
                else:
                    group_results.append(f"❌ {email} ({msg})")
                
                time.sleep(2) 

            # Report Data
            ok_cnt = sum(1 for r in group_results if "✅" in r)
            fail_cnt = sum(1 for r in group_results if "❌" in r)
            
            report_rows.append({
                "Company / Website": display_grp_name,
                "Template Used": curr_tpl['id'],
                "Total Sent": ok_cnt,
                "Total Failed": fail_cnt,
                "Details": ", ".join(group_results)
            })
            
            curr_idx += 1
            prog_bar.progress(curr_idx / total_groups)
            time.sleep(random.randint(d_min, d_max))

        st.balloons()
        st.success("🎉 Campaign Finished!")
        
        # --- 5. DOWNLOAD REPORT ---
        if report_rows:
            res_df = pd.DataFrame(report_rows)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                res_df.to_excel(writer, index=False)
                worksheet = writer.sheets['Sheet1']
                worksheet.set_column('A:A', 25)
                worksheet.set_column('E:E', 60)
            
            st.download_button("📥 Download Final Report", buffer, "SEO_Report.xlsx")

# ==========================================
# 6. INSTRUCTIONS SECTION (NEW)
# ==========================================
st.divider()
with st.expander("📖 Guide: How to use this tool?", expanded=True):
    st.markdown("""
    <div class="instruction-box">
    <h3>🚀 Step-by-Step Instructions</h3>
    <ol>
        <li><strong>Prepare Excel File:</strong> Ensure your file has these headers: <code>Name</code>, <code>Email</code>, <code>Company</code> (or Website).</li>
        <li><strong>Configure Email:</strong> In the Sidebar, select your provider (e.g., Hostinger) and enter your Email/Password.</li>
        <li><strong>Test Connection:</strong> Click the 🔌 <b>Test Connection</b> button in the sidebar to verify your login.</li>
        <li><strong>Upload Data:</strong> Upload your .xlsx or .csv file in Step 1.</li>
        <li><strong>Setup Templates:</strong> 
            <ul>
                <li>Go to Template 1 tab.</li>
                <li>Enter Subject Lines (one per line). Use <code>{Name}</code> or <code>{Company}</code> placeholders.</li>
                <li>Upload your HTML file containing the email body.</li>
            </ul>
        </li>
        <li><strong>Start:</strong> Click <b>Start Campaign</b>. Don't close the tab until finished.</li>
    </ol>
    <p><strong>Note:</strong> If you use Gmail, you must use an <b>App Password</b>, not your regular password.</p>
    </div>
    """, unsafe_allow_html=True)
