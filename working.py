from flask import Flask, render_template, request, redirect, url_for, jsonify
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from apscheduler.schedulers.background import BackgroundScheduler
import datetime
import time
import json
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from apscheduler.triggers.cron import CronTrigger
import os

app = Flask(__name__)

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
scheduler.start()
recipients = []
JOBS_FILE = "jobs.json"
LOG_FILE = "email_logs.json"

if not os.path.exists(LOG_FILE):
    with open(LOG_FILE, "w") as f:
        json.dump([], f)

def log_email(recipient, subject, status):
    with open(LOG_FILE, "r") as f:
        logs = json.load(f)
    logs.append({
        "recipient": recipient,
        "subject": subject,
        "status": status,
        "timestamp": str(datetime.datetime.now())
    })
    with open(LOG_FILE, "w") as f:
        json.dump(logs, f, indent=4)

def capture_screenshot():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    
    service = Service(EdgeChromiumDriverManager().install())
    driver = webdriver.Edge(service=service, options=options)
    
    power_bi_embed_url = f"https://app.powerbi.com/reportEmbed?reportId={REPORT_ID}&groupId={WORKSPACE_ID}&autoAuth=true&ctid={TENANT_ID}"
    driver.get(power_bi_embed_url)
    time.sleep(10)
    
    screenshot_path = "static/powerbi_report_screenshot.png"
    driver.save_screenshot(screenshot_path)
    driver.quit()
    
    return screenshot_path

def send_email(recipient_email, subject, body):
    screenshot_path = capture_screenshot()
    
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = recipient_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

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
        log_email(recipient_email, subject, "Sent")
        print(f"✅ Email sent to {recipient_email}")
    except Exception as e:
        log_email(recipient_email, subject, f"Failed: {str(e)}")
        print(f"❌ Email failed: {str(e)}")

def schedule_email(email, subject, body, hour, minute, frequency, days_of_week=None):
    job_id = f"job_{len(scheduler.get_jobs()) + 1}"
    
    # Day of the week mapping (Form input to Cron-compatible numbers)
    days_mapping = {
        "Monday": "0",
        "Tuesday": "1",
        "Wednesday": "2",
        "Thursday": "3",
        "Friday": "4",
        "Saturday": "5",
        "Sunday": "6"
    }
    
    if frequency == "daily":
        scheduler.add_job(send_email, "cron", id=job_id, hour=hour, minute=minute, args=[email, subject, body])
    elif frequency == "weekly" and days_of_week:
        # Convert days_of_week to cron-friendly format using the mapping
        days_of_week_cron = [days_mapping[day] for day in days_of_week if day in days_mapping]
        days_of_week_str = ','.join(days_of_week_cron)  # Create comma-separated string of cron-compatible day numbers
        scheduler.add_job(send_email, "cron", id=job_id, day_of_week=days_of_week_str, hour=hour, minute=minute, args=[email, subject, body])
    
    save_jobs()
    print(f"✅ Scheduled {frequency} email for {email} at {hour}:{minute} on {days_of_week_str if days_of_week else 'N/A'}")

    
    save_jobs()
    print(f"✅ Scheduled {frequency} email for {email} at {hour}:{minute} on {days_of_week_str if days_of_week else 'N/A'}")

def save_jobs():
    jobs = get_jobs()
    with open(JOBS_FILE, "w") as f:
        json.dump(jobs, f, indent=4)

def get_jobs():
    jobs = []
    for job in scheduler.get_jobs():
        if isinstance(job.trigger, CronTrigger):
            # Check if 'day_of_week' exists in the cron fields
            days_of_week_field = next((field for field in job.trigger.fields if field.name == "day_of_week"), None)
            frequency = "weekly" if days_of_week_field else "daily"
        else:
            frequency = "unknown"

        jobs.append({
            "id": job.id,
            "email": job.args[0],
            "subject": job.args[1],
            "body": job.args[2],
            "frequency": frequency
        })
    return jobs

@app.route("/schedule_email", methods=["POST"])
def schedule():
    email = request.form["email"]
    subject = request.form["subject"]
    body = request.form["body"]
    hour = int(request.form["hour"])
    minute = int(request.form["minute"])
    frequency = request.form["frequency"]
    days_of_week = request.form.getlist("days_of_week") if frequency == "weekly" else None
    
    schedule_email(email, subject, body, hour, minute, frequency, days_of_week)
    return redirect(url_for("index"))

@app.route("/jobs", methods=["GET"])
def list_jobs():
    jobs = get_jobs()
    return jsonify(jobs)

@app.route("/delete_job/<job_id>", methods=["POST"])
def delete_job(job_id):
    scheduler.remove_job(job_id)
    save_jobs()
    return redirect(url_for("index"))

@app.route('/')
def index():
    jobs = get_jobs()
    capture_screenshot()
    return render_template('index.html', recipients=recipients, report_image="static/powerbi_report_screenshot.png", jobs=jobs)

@app.route('/logs', methods=['GET'])
def view_logs():
    with open(LOG_FILE, "r") as f:
        logs = json.load(f)
    return jsonify(logs)

if __name__ == "__main__":
    app.run(debug=True)