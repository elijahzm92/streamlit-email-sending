import streamlit as st
import pandas as pd
import base64
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from email.mime.text import MIMEText
import json

SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# Load credentials from Streamlit secrets
def authenticate_gmail():
    creds_dict = json.loads(st.secrets["google_credentials"])  # Load credentials from secrets
    creds = Credentials.from_authorized_user_info(creds_dict, SCOPES)
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
            for _, row in df.iterrows():
                company = row["Company"]
                email = row["Email"]
                personalized_body = f"Dear {company},\n\n{body}"
                send_email(service, email, subject, personalized_body)
            st.success("Emails sent successfully!")
        else:
            st.error("Authentication failed. Please check your credentials.")
