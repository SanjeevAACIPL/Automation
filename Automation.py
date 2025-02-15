import streamlit as st
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from apscheduler.schedulers.background import BackgroundScheduler
import datetime
import time
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options

import streamlit as st

st.set_page_config(page_title="Power BI Report Scheduler", page_icon="📊", layout="wide")


# Power BI API Credentials
TENANT_ID = "be260a84-0301-4296-8976-9b876609e454"
CLIENT_ID = "68826f88-a00f-4cfe-a12e-004f94b95cd0"
CLIENT_SECRET = "HMU8Q~TRM3j8bKnO3idozACVXkaKRFpVlg21KaPI"
WORKSPACE_ID = "f1fa71d4-3e0a-4992-a4f5-a12089caf254"
REPORT_ID = "63a24ff0-2da7-4d39-baa4-865d9987af35"

# Email Configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "no-reply@aacipl.com"
SENDER_PASSWORD = "fuwb omph fcze wdpq"

# Scheduler
scheduler = BackgroundScheduler()
schedules = []

# Store added recipients in session state
if "recipients" not in st.session_state:
    st.session_state["recipients"] = []

# Function to capture Power BI Screenshot
def capture_screenshot():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")

    edge_service = Service(executable_path=r"D:\Sanjeev\Python Projects\email automation for deployment\edgedriver_win64\msedgedriver.exe")


    
    driver = webdriver.Edge(service=edge_service, options=options)

    power_bi_embed_url = f"https://app.powerbi.com/reportEmbed?reportId={REPORT_ID}&groupId={WORKSPACE_ID}&autoAuth=true&ctid={TENANT_ID}"

    driver.get(power_bi_embed_url)
    time.sleep(10)  # Allow time for report to load

    screenshot_path = "powerbi_report_screenshot.png"
    driver.save_screenshot(screenshot_path)
    driver.quit()
    
    return screenshot_path

def send_email(recipient_email, subject, body):
    screenshot_path = capture_screenshot()
    if not screenshot_path:
        return "❌ Failed to capture screenshot"

    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = recipient_email
    msg["Subject"] = subject  # Use custom subject

    # Attach email body
    msg.attach(MIMEText(body, "plain"))  # Use custom body

    with open(screenshot_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename=PowerBI_Report.png")
        msg.attach(part)

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, recipient_email, msg.as_string())
        server.quit()
        st.success(f"✅ Email sent to {recipient_email}!")
    except Exception as e:
        st.error(f"❌ Email failed: {str(e)}")




# **Power BI Report Viewer**
st.header("📈 Live Power BI Report")
power_bi_embed_url = f"https://app.powerbi.com/reportEmbed?reportId={REPORT_ID}&groupId={WORKSPACE_ID}&autoAuth=true&ctid={TENANT_ID}"
st.markdown(f'<iframe width="100%" height="600px" src="{power_bi_embed_url}" frameborder="0" allowFullScreen></iframe>', unsafe_allow_html=True)

# **Recipient Management**
st.sidebar.header("📧 Manage Recipients")

# Input field for new recipient
new_recipient = st.sidebar.text_input("Add Recipient Email")

# **Add Recipient Button**
if st.sidebar.button("➕ Add Recipient"):
    if new_recipient.strip() and new_recipient not in st.session_state["recipients"]:
        st.session_state["recipients"].append(new_recipient.strip())
        st.sidebar.success(f"✅ {new_recipient} added!")
    elif new_recipient in st.session_state["recipients"]:
        st.sidebar.warning(f"⚠️ {new_recipient} is already added!")
    else:
        st.sidebar.warning("⚠️ Please enter a valid email address.")

# **List of Added Recipients**
st.sidebar.subheader("📜 Recipient List")
if st.session_state["recipients"]:
    for email in st.session_state["recipients"]:
        col1, col2 = st.sidebar.columns([3, 1])
        col1.write(f"📧 {email}")
        if col2.button("❌", key=f"remove_{email}"):
            st.session_state["recipients"].remove(email)
            st.sidebar.success(f"✅ {email} removed!")

# Editable Email Subject and Body
email_subject = st.sidebar.text_input("✉️ Email Subject", "Power BI Report Screenshot")
email_body = st.sidebar.text_area("📝 Email Body", "Please find the attached Power BI report screenshot.")

# **Schedule Email Section**
st.sidebar.header("📅 Schedule a Report")

# Time Selection
selected_hour = st.sidebar.selectbox("⏰ Select Hour", list(range(24)), index=9)
selected_minute = st.sidebar.selectbox("⏰ Select Minute", list(range(60)), index=0)

# Combine Hour and Minute into a single time object
selected_time = datetime.time(selected_hour, selected_minute)
st.sidebar.write(f"🔹 Selected Time: {selected_time.strftime('%H:%M')}")

# **Scheduling Options**
schedule_type = st.sidebar.radio("🔁 Select Frequency", ["Once", "Daily", "Weekly", "Monthly"], index=0)

# Conditional Fields for Weekly & Monthly
selected_weekday = None
selected_date = None

if schedule_type == "Weekly":
    selected_weekday = st.sidebar.selectbox("📆 Select Day", ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"])
elif schedule_type == "Monthly":
    selected_date = st.sidebar.slider("📅 Select Date", 1, 31, 1)

# **Schedule Button**
if st.sidebar.button("📌 Schedule Report"):
    if not email:
        st.sidebar.warning("⚠️ Please enter an email before scheduling!")
    else:
        hour = selected_time.hour
        minute = selected_time.minute

        if schedule_type == "Once":
            scheduler.add_job(send_email, "date", run_date=datetime.datetime.now(), args=[email, email_subject, email_body])

        elif schedule_type == "Daily":
            scheduler.add_job(send_email, "cron", hour=hour, minute=minute, args=[email, email_subject, email_body])

        elif schedule_type == "Weekly" and selected_weekday:
            weekdays_mapping = {
                "Monday": "mon",
                "Tuesday": "tue",
                "Wednesday": "wed",
                "Thursday": "thu",
                "Friday": "fri",
                "Saturday": "sat",
                "Sunday": "sun"
            }
            weekday_cleaned = weekdays_mapping.get(selected_weekday)

            if weekday_cleaned:
                scheduler.add_job(send_email, "cron", day_of_week=weekday_cleaned, hour=hour, minute=minute, args=[email, email_subject, email_body])
            else:
                st.sidebar.error("⚠️ Invalid weekday selected!")

        elif schedule_type == "Monthly" and selected_date:
            scheduler.add_job(send_email, "cron", day=selected_date, hour=hour, minute=minute, args=[email, email_subject, email_body])

        schedules.append({
            "email": email,
            "frequency": schedule_type,
            "time": selected_time.strftime("%H:%M"),
            "subject": email_subject,
            "body": email_body
        })
        st.sidebar.success(f"✅ {schedule_type} report scheduled for {email} at {selected_time.strftime('%H:%M')}.")


# Start the scheduler
if not scheduler.running:
    scheduler.start()
