import streamlit as st
import json
import requests
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import base64
from email.mime.text import MIMEText
from urllib.parse import urlencode

SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# Load credentials from Streamlit secrets
def authenticate_gmail():
    creds_dict = {
        "client_id": st.secrets["google_auth"]["client_id"],
        "client_secret": st.secrets["google_auth"]["client_secret"],
        "auth_uri": st.secrets["google_auth"]["auth_uri"],
        "token_uri": st.secrets["google_auth"]["token_uri"],
        "redirect_uris": [st.secrets["google_auth"]["redirect_uri"]]
    }

    flow = Flow.from_client_config({"web": creds_dict}, scopes=SCOPES)
    flow.redirect_uri = creds_dict["redirect_uris"][0]

    auth_url, state = flow.authorization_url(prompt="consent", access_type="offline")
    
    st.write(f"[Click here to authenticate]({auth_url})")
    
    return state  # Store this state to validate OAuth callback

# Function to handle OAuth callback
def handle_callback():
    query_params = st.experimental_get_query_params()
    if "code" in query_params:
        code = query_params["code"][0]
        creds_dict = {
            "client_id": st.secrets["google_auth"]["client_id"],
            "client_secret": st.secrets["google_auth"]["client_secret"],
            "token_uri": st.secrets["google_auth"]["token_uri"],
            "redirect_uris": [st.secrets["google_auth"]["redirect_uri"]]
        }
        flow = Flow.from_client_config({"web": creds_dict}, scopes=SCOPES)
        flow.redirect_uri = creds_dict["redirect_uris"][0]

        token = flow.fetch_token(code=code)
        creds = Credentials(**token)
        st.success("Authentication successful!")
        return creds
    return None

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

# OAuth Authentication
if st.button("Authenticate with Gmail"):
    state = authenticate_gmail()

# Handle callback from OAuth redirect
creds = handle_callback()

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])
subject = st.text_input("Enter Email Subject")
body = st.text_area("Enter Email Body")

if uploaded_file is not None and creds:
    df = pd.read_excel(uploaded_file)
    st.write("Preview of Uploaded File:", df.head())

    if st.button("Send Emails"):
        service = build("gmail", "v1", credentials=creds)
        for _, row in df.iterrows():
            company = row["Company"]
            email = row["Email"]
            personalized_body = f"Dear {company},\n\n{body}"
            send_email(service, email, subject, personalized_body)
        st.success("Emails sent successfully!")
