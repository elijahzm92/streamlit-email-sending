import streamlit as st
import pandas as pd
import base64
import os
import re
import time
import json
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

# Load secrets from Streamlit
if "credentials" in st.secrets:
    CLIENT_ID = st.secrets["credentials"]["client_id"]
    CLIENT_SECRET = st.secrets["credentials"]["client_secret"]
else:
    st.error("Missing credentials in Streamlit secrets")
    st.stop()

# Function to get authenticated service
def authenticate_gmail():
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_config({"installed": credentials_json["installed"]}, SCOPES)
            creds = flow.run_local_server(port=8501)
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return build("gmail", "v1", credentials=creds)

# Function to send email
def send_email(service, to_email, cc_email, subject, body, attachment):
    message = MIMEMultipart()
    message["to"] = to_email
    if cc_email:
        message["cc"] = cc_email
    message["subject"] = subject
    message.attach(MIMEText(body, "html"))
    
    if attachment:
        with open(attachment.name, "rb") as att:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(att.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={attachment.name}")
            message.attach(part)
    
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    try:
        service.users().messages().send(userId="me", body={"raw": raw_message}).execute()
        return "Success"
    except HttpError as error:
        return f"Error: {error}"

# Streamlit UI
st.title("Bulk Email Sender using Gmail API")
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
email_subject = st.text_input("Email Subject (Use {Company Name} for replacement)")
email_body = st.text_area("Email Body (Use {Company Name} and {Contact Name} for replacement)")
attachment = st.file_uploader("Upload an attachment (optional)")

if st.button("Send Emails"):
    if uploaded_file and email_subject and email_body:
        df = pd.read_excel(uploaded_file)
        if "Company Name" in df.columns and "Email" in df.columns:
            service = authenticate_gmail()
            for _, row in df.iterrows():
                company = row["Company Name"]
                to_email = row["Email"]
                cc_email = row["Alternate Email"] if "Alternate Email" in df.columns and not pd.isna(row["Alternate Email"]) else None
                contact_name = row["Contact Name"] if "Contact Name" in df.columns and not pd.isna(row["Contact Name"]) else ""
                
                subject = email_subject.replace("{Company Name}", company)
                body = email_body.replace("{Company Name}", company).replace("{Contact Name}", contact_name)
                
                status = send_email(service, to_email, cc_email, subject, body, attachment)
                st.write(f"Email to {to_email}: {status}")
        else:
            st.error("Excel file must contain 'Company Name' and 'Email' columns")
    else:
        st.error("Please upload an Excel file and fill email details")
