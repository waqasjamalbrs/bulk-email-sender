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
# 1. PAGE CONFIGURATION
# ==========================================
st.set_page_config(page_title="Pro SEO Outreach", page_icon="🚀", layout="wide")
st.markdown("""
    <style>
    /* UI Polish */
    .main { background-color: #f7f9fc; }
    .stButton>button { 
        background: linear-gradient(90deg, #1d976c, #93f9b9); 
        color: #004d2e; border: none; height: 50px; font-size: 18px; font-weight: bold;
        border-radius: 8px;
    }
    .stTextInput>div>div>input { border-radius: 5px; }
    /* Success/Error Boxes */
    .report-box { padding: 10px; border-radius: 5px; margin-bottom: 5px; color: #155724; background-color: #d4edda; border-color: #c3e6cb; }
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

# Technical domain extractor (Sirf backend grouping k liye)
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
    
    st.divider()
    st.subheader("3. Safety Limits")
    limit = st.number_input("Stop after sending X emails", 50, 5000, 100)
    d_min = st.number_input("Min Delay (Seconds)", 5)
    d_max = st.number_input("Max Delay (Seconds)", 20)

# ==========================================
# 4. MAIN INTERFACE
# ==========================================
st.title("📨 Ultimate SEO Outreach Tool")
st.markdown("Supports: **Multi-Subject Rotation**, **Company Grouping**, & **Strict Variables**.")

col1, col2 = st.columns([1, 2])

# --- LEFT COLUMN: FILE ---
with col1:
    st.subheader("📂 Step 1: Upload Data")
    f = st.file_uploader("Upload Excel (.xlsx) or CSV", type=['xlsx', 'csv'])
    st.info("Required Columns: `Name`, `Email`. \n\nRecommended: `Company` or `Website`.")

# --- RIGHT COLUMN: TEMPLATES ---
with col2:
    st.subheader("📝 Step 2: Templates Strategy")
    st.caption("Add different templates. The system will rotate them per company.")
    
    tabs = st.tabs(["Template 1", "Template 2", "Template 3", "Template 4", "Template 5"])
    tpls = []
    
    for i, tab in enumerate(tabs):
        with tab:
            st.markdown(f"**Template {i+1} Setup**")
            
            # Subject Input
            raw_s = st.text_area(
                f"Subject Lines (One per line) - Randomly picked", 
                height=100, 
                key=f"s_{i}",
                placeholder="Collaboration opportunity for {Website}\nQuick question regarding {Company}\nFeature inquiry for {Name}"
            )
            
            # Body Input
            raw_b = st.text_area(
                f"Email Body (HTML Supported)", 
                height=200, 
                key=f"b_{i}",
                placeholder="<p>Hi {Name},</p>\n<p>I saw your work at {Company}...</p>"
            )
            
            # Store Valid Templates
            if raw_s and raw_b:
                # Split subjects by new line
                s_list = [x.strip() for x in raw_s.split('\n') if x.strip()]
                tpls.append({"id": f"Template {i+1}", "subjects": s_list, "body": raw_b})

# ==========================================
# 5. LOGIC & EXECUTION
# ==========================================
st.divider()

if st.button("🚀 START CAMPAIGN"):
    # --- VALIDATION ---
    if not f:
        st.error("❌ Please upload a file first.")
    elif not e_user or not e_pass:
        st.error("❌ Please enter your Email and Password in Sidebar.")
    elif not tpls:
        st.error("❌ Please add at least one Template (Subject + Body).")
    else:
        # --- 1. READ FILE ---
        try:
            if f.name.endswith('.csv'): df = pd.read_csv(f)
            else: df = pd.read_excel(f)
        except Exception as e:
            st.error(f"Error reading file: {e}")
            st.stop()
            
        # --- 2. PRE-PROCESS DATA (Strict Logic) ---
        all_contacts = []
        
        for _, row in df.iterrows():
            # Handle multiple emails (comma separated)
            raw_emails = str(row.get('Email', '')).split(',')
            
            # Basic Variables
            name_val = str(row.get('Name', 'there')).strip()
            
            # Strict Company/Website Logic
            company_col = str(row.get('Company', '')).strip()
            website_col = str(row.get('Website', '')).strip()
            
            # Clean up Pandas 'nan'
            if company_col.lower() == 'nan': company_col = ""
            if website_col.lower() == 'nan': website_col = ""
            
            # DETERMINE DISPLAY TEXT (For {Company} / {Website})
            # Agar Column khali hai to 'your company' use hoga.
            # Gmail/Hotmail domain kabhi use nahi hoga.
            
            final_comp_txt = company_col if company_col else (website_col if website_col else "your company")
            final_web_txt = website_col if website_col else (company_col if company_col else "your website")
            
            for em in raw_emails:
                clean_em = em.strip()
                if "@" in clean_em:
                    # GROUPING KEY LOGIC
                    # Group karne k liye hum Domain use kar sakte hain agar company name na ho.
                    # Ye sirf "Batching" k liye hai, text k liye nahi.
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
        
        st.success(f"✅ Loaded {len(all_contacts)} emails. Grouped into {len(grouped_data)} unique Companies/Websites.")
        
        # --- 4. SENDING LOOP ---
        prog_bar = st.progress(0)
        status_text = st.empty()
        
        report_rows = []
        sent_total = 0
        cycle = itertools.cycle(tpls)
        curr_idx = 0
        total_groups = len(grouped_data)
        
        # Loop through Companies
        for g_key, contacts in grouped_data.items():
            if sent_total >= limit: 
                st.warning("🛑 Daily Limit Reached! Stopping process."); break
            
            # Pick Template for this Company
            curr_tpl = next(cycle)
            group_results = []
            
            # UI Status Update
            # Show Company Name if available, else show Domain
            display_grp_name = contacts[0]['DisplayComp'] if contacts[0]['DisplayComp'] != "your company" else g_key
            status_text.markdown(f"**Processing:** `{display_grp_name}` | using {curr_tpl['id']}")
            
            # Loop through People in this Company
            for person in contacts:
                email = person['Email']
                name = person['Name']
                
                # Pick Random Subject
                if curr_tpl['subjects']: 
                    subj_base = random.choice(curr_tpl['subjects'])
                else: 
                    subj_base = "Hello" # Fallback
                
                # --- PERSONALIZATION ---
                # 1. Standard Tags
                sub = subj_base.replace("{Name}", name).replace("{Company}", person['DisplayComp']).replace("{Website}", person['DisplayWeb'])
                bod = curr_tpl['body'].replace("{Name}", name).replace("{Company}", person['DisplayComp']).replace("{Website}", person['DisplayWeb'])
                
                # 2. Extra Columns from Excel
                for col in df.columns:
                     val = str(person['RowData'].get(col, ''))
                     sub = sub.replace(f"{{{col}}}", val)
                     bod = bod.replace(f"{{{col}}}", val)
                
                # --- SENDING ---
                ok, msg = send_email_smtp(conf, e_user, e_pass, email, name, sender_name, sub, bod)
                
                if ok:
                    save_sent(conf, e_user, e_pass, msg)
                    sent_total += 1
                    group_results.append(f"✅ {email}")
                else:
                    group_results.append(f"❌ {email} ({msg})")
                
                # Small delay between colleagues (2 sec)
                time.sleep(2) 

            # Record Results
            ok_cnt = sum(1 for r in group_results if "✅" in r)
            fail_cnt = sum(1 for r in group_results if "❌" in r)
            
            report_rows.append({
                "Company / Website": display_grp_name,
                "Template Used": curr_tpl['id'],
                "Total Sent": ok_cnt,
                "Total Failed": fail_cnt,
                "Email Details": ", ".join(group_results)
            })
            
            # Update Progress
            curr_idx += 1
            prog_bar.progress(curr_idx / total_groups)
            
            # Big Random Delay between Companies
            time.sleep(random.randint(d_min, d_max))

        st.balloons()
        st.success("🎉 Campaign Finished!")
        
        # --- 5. REPORT DOWNLOAD ---
        if report_rows:
            res_df = pd.DataFrame(report_rows)
            st.dataframe(res_df)
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                res_df.to_excel(writer, index=False, sheet_name="Outreach Report")
                
                # Adjust column width for readability
                worksheet = writer.sheets['Outreach Report']
                worksheet.set_column('A:A', 25)
                worksheet.set_column('E:E', 60)
            
            st.download_button(
                label="📥 Download Final Report",
                data=buffer,
                file_name="SEO_Outreach_Report.xlsx",
                mime="application/vnd.ms-excel"
            )
