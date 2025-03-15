import os
import json
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import datetime
import pytz
import time
from google.oauth2 import service_account
from googleapiclient.discovery import build
from openai import OpenAI

# Configuration
GOOGLE_SHEET_ID = os.environ['GOOGLE_SHEET_ID']
GOOGLE_SHEET_RANGE = 'Sheet1!A2:C'  # Adjust based on your sheet structure
EMAIL_USERNAME = os.environ['EMAIL_USERNAME']
EMAIL_PASSWORD = os.environ['EMAIL_PASSWORD']
OPENAI_API_KEY = os.environ['OPENAI_API_KEY']
CV_PATH = os.environ['CV_PATH']
EMAILS_PER_DAY = 10
STATUS_FILE = 'data/email_status.json'

# Initialize OpenAI client
openai_client = OpenAI(api_key=OPENAI_API_KEY)

def get_professor_data():
    """Fetch professor data from Google Sheets"""
    creds = service_account.Credentials.from_service_account_file(
        'service_account.json', 
        scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
    )
    
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=GOOGLE_SHEET_ID, range=GOOGLE_SHEET_RANGE).execute()
    values = result.get('values', [])
    
    if not values:
        print('No data found in the Google Sheet')
        return []
    
    professors = []
    for row in values:
        if len(row) >= 3:
            professors.append({
                'name': row[0],
                'research_domain': row[1],
                'email': row[2],
                'university': row[3] if len(row) > 3 else 'Unknown'
            })
    
    return professors

def load_email_status():
    """Load the status of sent emails"""
    if os.path.exists(STATUS_FILE):
        with open(STATUS_FILE, 'r') as f:
            return json.load(f)
    return {"last_index": 0, "sent_emails": []}

def save_email_status(status):
    """Save the status of sent emails"""
    # Ensure the directory exists
    os.makedirs(os.path.dirname(STATUS_FILE), exist_ok=True)
    
    with open(STATUS_FILE, 'w') as f:
        json.dump(status, f, indent=2)

def generate_email_content(professor, cv_content):
    """Generate personalized email using OpenAI"""
    prompt = f"""
    Create a personalized email for a summer research internship application to a professor.
    
    Professor details:
    Name: {professor['name']}
    Research Domain: {professor['research_domain']}
    University: {professor['university']}
    
    My CV highlights:
    {cv_content[:500]}...
    
    Email requirements:
    1. Subject Line: Specific and attention-grabbing
    2. Salutation: Formal address
    3. Introduction: Brief self-introduction
    4. Why Them: Mention specific aspects of their research
    5. My Qualifications: Highlight relevant skills and projects
    6. Ask: Clearly state seeking a research internship
    7. Mention CV attachment
    8. Polite closing
    
    Keep the email between 150-200 words. Make it personalized, mentioning specific research areas.
    Format the response as a JSON with 'subject' and 'body' fields.
    """
    
    response = openai_client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}]
    )
    
    email_content = json.loads(response.choices[0].message.content)
    return email_content['subject'], email_content['body']

def read_cv():
    """Read the CV file contents"""
    try:
        with open(CV_PATH, 'rb') as f:
            # Just return the first few KB to give AI a sense of the CV
            return f.read(5000).decode('utf-8', errors='ignore')
    except Exception as e:
        print(f"Error reading CV: {e}")
        return "CV could not be read. Please ensure it exists at the specified path."

def send_email(professor, subject, body):
    """Send an email to the professor"""
    msg = MIMEMultipart()
    msg['From'] = EMAIL_USERNAME
    msg['To'] = professor['email']
    msg['Subject'] = subject
    
    msg.attach(MIMEText(body, 'plain'))
    
    # Attach CV
    with open(CV_PATH, 'rb') as f:
        cv_attachment = MIMEApplication(f.read(), _subtype="pdf")
        cv_attachment.add_header('Content-Disposition', 'attachment', filename="CV.pdf")
        msg.attach(cv_attachment)
    
    # Send email
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(EMAIL_USERNAME, EMAIL_PASSWORD)
        server.send_message(msg)
    
    print(f"Email sent to {professor['name']} at {professor['email']}")
    return True

def main():
    """Main function to send emails"""
    professors = get_professor_data()
    if not professors:
        print("No professor data found.")
        return
    
    status = load_email_status()
    last_index = status["last_index"]
    sent_emails = status["sent_emails"]
    
    cv_content = read_cv()
    
    emails_sent = 0
    for i in range(last_index, min(last_index + EMAILS_PER_DAY, len(professors))):
        professor = professors[i]
        
        # Skip if email already sent
        if professor['email'] in sent_emails:
            continue
        
        try:
            subject, body = generate_email_content(professor, cv_content)
            
            # Send email
            success = send_email(professor, subject, body)
            
            if success:
                sent_emails.append(professor['email'])
                emails_sent += 1
                
                # Wait a bit between emails
                time.sleep(30)
        except Exception as e:
            print(f"Error sending email to {professor['name']}: {e}")
    
    # Update status
    status["last_index"] = min(last_index + EMAILS_PER_DAY, len(professors))
    status["sent_emails"] = sent_emails
    save_email_status(status)
    
    print(f"Sent {emails_sent} emails. Next start index: {status['last_index']}")

if __name__ == "__main__":
    main()
