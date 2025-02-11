import streamlit as st
import pandas as pd
import base64
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from email.mime.text import MIMEText
import os

SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# Function to authenticate and get Gmail service
def authenticate_gmail():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
        creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds

# Function to create email message
def create_message(to, subject, body):
    message = MIMEText(body)
    message["to"] = to
    message["subject"] = subject
    return {"raw": base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")}

# Function to send email
def send_email(service, to, subject, body):
    message = create_message(to, subject, body)
    send_message = service.users().messages().send(userId="me", body=message).execute()
    return send_message

# Streamlit UI
st.title("Bulk Email Sender using Gmail API")

if st.button("Authenticate with Gmail"):
    creds = authenticate_gmail()
    st.success("Authentication successful!")

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])
subject = st.text_input("Enter Email Subject")
body = st.text_area("Enter Email Body")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("Preview of Uploaded File:", df.head())
    
    if st.button("Send Emails"):
        creds = authenticate_gmail()
        if creds:
            service = build("gmail", "v1", credentials=creds)
            for index, row in df.iterrows():
                company = row["Company"]
                email = row["Email"]
                personalized_body = f"Dear {company},\n\n{body}"  # Personalizing email
                send_email(service, email, subject, personalized_body)
            st.success("Emails sent successfully!")
        else:
            st.error("Authentication failed. Please try again.")
