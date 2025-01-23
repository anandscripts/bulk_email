import smtplib, ssl
from email.message import EmailMessage
import csv
import google.generativeai as genai
from dotenv import load_dotenv
import os
import gspread
from datetime import datetime

# Getting the API Key and the APP Password from the .env file
load_dotenv()
API_KEY = os.getenv("GEMINI_API")
APP_PASSWORD = os.getenv("APP_PASSWORD")

genai.configure(api_key=API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")

# Function to generate the email content with the details
def email_content(name, email, my_roles, recipient_role, company_name):
    prompt = f'''
    Generate a professional email for contacting a tech company. The email should address either the HR department or a manager, depending on the role specified. It must mention the company name and highlight the sender's willingness to contribute in one of several specified roles (e.g., Software Engineer, Python Developer, AI Engineer). The content should also mention that the sender's resume is attached and that they are inquiring about any current or upcoming job vacancies. Ensure the tone is formal, respectful and eager to work, and encourage the recipient to review the resume and consider the sender for relevant opportunities. Remember: don't include the subject line.

    Inputs:
    Name: {name}
    Email: {email}
    Sender's Roles: {my_roles}
    Recipient's Role: {recipient_role}
    Company Name: {company_name}

    Example Output: 
    Dear [Recipient's Role] at [Company Name],\n
    [Email content]
    Warm regards,
    [Name]
    [Email]
    '''
    summarized_text = model.generate_content(prompt)
    return summarized_text.text

# Your details
NAME = 'Anand'
EMAIL = 'emailid@gmail.com'
ROLES = 'Software Engineer, Python Developer, and AI Engineer'
RESUME_PATH = 'path/to/Resume.pdf'
SHEET_NAME = 'Bulk Emails'

def setup_google_sheets(sheet_name):
    gc = gspread.service_account(filename='credentials.json')
    sheet = gc.open(sheet_name).sheet1 
    return sheet

def log_email_sent(sheet, recipient_email, recipient_role, company_name):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
    sheet.append_row([company_name, recipient_role, recipient_email, timestamp])

sheet = setup_google_sheets(SHEET_NAME)

# SMTP Server configuration
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 465

# Open the CSV and get the details
with open("mail_ids.csv", "r") as csvfile:
    reader = csv.reader(csvfile)
    for row in enumerate(reader):
        recipient_email = row[1][0]
        recipient_role = row[1][1] 
        company_name = row[1][2]

        # Email message details
        email_message = EmailMessage()
        email_message['From'] = f'{NAME} <{EMAIL}>'
        email_message['Reply-To'] = EMAIL
        email_message['To'] = recipient_email

        email_message['Subject'] = f'Application for Career Opportunities at {company_name}'

        # Calling the email_content() to generate the content
        content = email_content(NAME, EMAIL, ROLES, recipient_role, company_name)
        email_message.set_content(content)

        # Open the Resume and Add to the Email as an attachment
        try:
            with open(RESUME_PATH, 'rb') as file:
                file_data = file.read()
                file_name = RESUME_PATH.split('/')[-1]  
                email_message.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)
        except Exception as e:
            print(f"Error attaching file: {e}")

        # Log in and Send the email
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
                server.login(EMAIL, APP_PASSWORD)
                server.send_message(email_message)
                log_email_sent(sheet, recipient_email, recipient_role, company_name)
            print(f"Email sent to {recipient_email}")
        except Exception as e:
            print(f"Failed to send email to {recipient_email}: {e}")
