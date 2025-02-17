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
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import os

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

scheduler = BackgroundScheduler()
schedules = []

if "recipients" not in st.session_state:
    st.session_state["recipients"] = []

def capture_screenshot():
    try:
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")

        driver = webdriver.Chrome(service=Service("/usr/bin/chromedriver"), options=chrome_options)
        power_bi_url = f"https://app.powerbi.com/reportEmbed?reportId={REPORT_ID}&groupId={WORKSPACE_ID}&autoAuth=true&ctid={TENANT_ID}"
        driver.get(power_bi_url)
        time.sleep(10)

        screenshot_path = "powerbi_report_screenshot.png"
        driver.save_screenshot(screenshot_path)
        driver.quit()
        return screenshot_path
    except Exception as e:
        st.error(f"❌ Screenshot capture failed: {str(e)}")
        return None

def send_email(recipient_email, subject, body):
    screenshot_path = capture_screenshot()
    if not screenshot_path:
        return "❌ Failed to capture screenshot"
    
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = recipient_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    with open(screenshot_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", "attachment; filename=PowerBI_Report.png")
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

st.header("📈 Live Power BI Report")
power_bi_url = f"https://app.powerbi.com/reportEmbed?reportId={REPORT_ID}&groupId={WORKSPACE_ID}&autoAuth=true&ctid={TENANT_ID}"
st.markdown(f'<iframe width="100%" height="600px" src="{power_bi_url}" frameborder="0" allowFullScreen></iframe>', unsafe_allow_html=True)

st.sidebar.header("📧 Manage Recipients")
new_recipient = st.sidebar.text_input("Add Recipient Email")
if st.sidebar.button("➕ Add Recipient"):
    if new_recipient.strip() and new_recipient not in st.session_state["recipients"]:
        st.session_state["recipients"].append(new_recipient.strip())
        st.sidebar.success(f"✅ {new_recipient} added!")
    elif new_recipient in st.session_state["recipients"]:
        st.sidebar.warning(f"⚠️ {new_recipient} is already added!")
    else:
        st.sidebar.warning("⚠️ Please enter a valid email address.")

st.sidebar.subheader("📜 Recipient List")
for email in st.session_state["recipients"]:
    col1, col2 = st.sidebar.columns([3, 1])
    col1.write(f"📧 {email}")
    if col2.button("❌", key=f"remove_{email}"):
        st.session_state["recipients"].remove(email)
        st.sidebar.success(f"✅ {email} removed!")

email_subject = st.sidebar.text_input("✉️ Email Subject", "Power BI Report Screenshot")
email_body = st.sidebar.text_area("📝 Email Body", "Please find the attached Power BI report screenshot.")

st.sidebar.header("📅 Schedule a Report")
selected_hour = st.sidebar.selectbox("⏰ Select Hour", list(range(24)), index=9)
selected_minute = st.sidebar.selectbox("⏰ Select Minute", list(range(60)), index=0)
selected_time = datetime.time(selected_hour, selected_minute)
st.sidebar.write(f"🔹 Selected Time: {selected_time.strftime('%H:%M')}")

schedule_type = st.sidebar.radio("🔁 Select Frequency", ["Once", "Daily", "Weekly", "Monthly"], index=0)
selected_weekday = None
selected_date = None
if schedule_type == "Weekly":
    selected_weekday = st.sidebar.selectbox("📆 Select Day", ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"])
elif schedule_type == "Monthly":
    selected_date = st.sidebar.slider("📅 Select Date", 1, 31, 1)

if st.sidebar.button("📌 Schedule Report"):
    for email in st.session_state["recipients"]:
        if schedule_type == "Once":
            scheduler.add_job(send_email, "date", run_date=datetime.datetime.now(), args=[email, email_subject, email_body])
        elif schedule_type == "Daily":
            scheduler.add_job(send_email, "cron", hour=selected_time.hour, minute=selected_time.minute, args=[email, email_subject, email_body])
        elif schedule_type == "Weekly" and selected_weekday:
            scheduler.add_job(send_email, "cron", day_of_week=selected_weekday[:3].lower(), hour=selected_time.hour, minute=selected_time.minute, args=[email, email_subject, email_body])
        elif schedule_type == "Monthly" and selected_date:
            scheduler.add_job(send_email, "cron", day=selected_date, hour=selected_time.hour, minute=selected_time.minute, args=[email, email_subject, email_body])
    st.sidebar.success(f"✅ {schedule_type} report scheduled for recipients at {selected_time.strftime('%H:%M')}.")

if not scheduler.running:
    scheduler.start()
