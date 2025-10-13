import streamlit as st
import pandas as pd
import base64
import time
import re
import json
import random
import os
from datetime import datetime, timedelta
import pytz
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

# ========================================
# Streamlit Page Setup
# ========================================
st.set_page_config(page_title="Gmail Mail Merge", layout="wide")
st.title("📧 Gmail Mail Merge Tool (with Follow-up Replies + Draft Save + Auto Backup)")

# ========================================
# Gmail API Setup
# ========================================
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.labels",
    "https://www.googleapis.com/auth/gmail.compose",
]

CLIENT_CONFIG = {
    "web": {
        "client_id": st.secrets["gmail"]["client_id"],
        "client_secret": st.secrets["gmail"]["client_secret"],
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "redirect_uris": [st.secrets["gmail"]["redirect_uri"]],
    }
}

# ========================================
# Helper Functions
# ========================================
EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")

def extract_email(value: str):
    if not value:
        return None
    match = EMAIL_REGEX.search(str(value))
    return match.group(0) if match else None


def convert_bold(text):
    if not text:
        return ""
    text = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", text)
    text = re.sub(
        r"\[(.*?)\]\((https?://[^\s)]+)\)",
        r'<a href="\2" style="color:#1a73e8; text-decoration:underline;" target="_blank">\1</a>',
        text,
    )
    text = text.replace("\n", "<br>").replace("  ", "&nbsp;&nbsp;")
    return f"""
    <html>
        <body style="font-family: Verdana, sans-serif; font-size: 14px; line-height: 1.6;">
            {text}
        </body>
    </html>
    """


def get_or_create_label(service, label_name="Mail Merge Sent"):
    try:
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        for label in labels:
            if label["name"].lower() == label_name.lower():
                return label["id"]
        label_obj = {
            "name": label_name,
            "labelListVisibility": "labelShow",
            "messageListVisibility": "show",
        }
        created_label = service.users().labels().create(userId="me", body=label_obj).execute()
        return created_label["id"]
    except Exception as e:
        st.warning(f"Could not get/create label: {e}")
        return None


def safe_format_template(template: str, row: pd.Series):
    try:
        return template.format(**(row.to_dict()))
    except Exception:
        try:
            keys = re.findall(r"\{(.*?)\}", template)
            ctx = {k: str(row.get(k, "")) for k in keys}
            return template.format(**ctx)
        except Exception:
            return template


def ensure_creds_and_update_session(creds):
    try:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            st.session_state["creds"] = creds.to_json()
    except Exception as e:
        st.warning(f"Could not refresh credentials: {e}")
    return creds


def send_email_backup(service, to_email, csv_path):
    """Send backup CSV to the authenticated Gmail account."""
    try:
        msg = MIMEMultipart()
        msg["To"] = to_email
        msg["From"] = "me"
        msg["Subject"] = "📁 Mail Merge Backup CSV"

        body = MIMEText("Attached is the backup CSV file for your recent mail merge run.", "plain")
        msg.attach(body)

        with open(csv_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(csv_path))
        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(csv_path)}"'
        msg.attach(part)

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId="me", body={"raw": raw}).execute()

        st.info(f"📧 Backup CSV emailed to your Gmail inbox ({to_email}).")
    except Exception as e:
        st.warning(f"⚠️ Could not send backup email: {e}")

# ========================================
# OAuth Flow
# ========================================
if "creds" not in st.session_state:
    st.session_state["creds"] = None

if st.session_state["creds"]:
    creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
else:
    code = st.experimental_get_query_params().get("code", None)
    if code:
        flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
        flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
        flow.fetch_token(code=code[0])
        creds = flow.credentials
        st.session_state["creds"] = creds.to_json()
        st.rerun()
    else:
        flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
        flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
        auth_url, _ = flow.authorization_url(
            prompt="consent", access_type="offline", include_granted_scopes="true"
        )
        st.markdown(
            f"### 🔑 Please [authorize the app]({auth_url}) to send emails using your Gmail account."
        )
        st.stop()

creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
creds = ensure_creds_and_update_session(creds)
service = build("gmail", "v1", credentials=creds)

# ========================================
# Recover Last Backup (if exists)
# ========================================
if "last_saved_csv" in st.session_state:
    st.info("📂 Backup from last session available:")
    st.download_button(
        "⬇️ Download Last Saved CSV",
        data=open(st.session_state["last_saved_csv"], "rb"),
        file_name=st.session_state["last_saved_name"],
        mime="text/csv",
    )

# ========================================
# Upload Recipients
# ========================================
st.header("📤 Upload Recipient List")
st.info("⚠️ Upload maximum of **70–80 contacts** for safe Gmail sending.")

uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

if uploaded_file:
    df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith("csv") else pd.read_excel(uploaded_file)
    st.write("✅ Preview of uploaded data:")
    st.dataframe(df.head())
    st.info("📌 Include 'ThreadId' and 'RfcMessageId' columns for follow-ups if needed.")

    # ========================================
    # Email Template
    # ========================================
    st.header("✍️ Compose Your Email")
    subject_template = st.text_input("Subject", "Hello {Name}")
    body_template = st.text_area(
        "Body (supports **bold**, [link](https://example.com), and line breaks)",
        """Dear {Name},

Welcome to our **Mail Merge App** demo.

You can add links like [Visit Google](https://google.com)
and preserve formatting exactly.

Thanks,  
**Your Company**""",
        height=250,
    )

    # ========================================
    # Preview Section
    # ========================================
    st.subheader("👁️ Preview Email")
    if not df.empty and "Email" in df.columns:
        recipient_options = df["Email"].astype(str).tolist()
        selected_email = st.selectbox("Select recipient to preview", recipient_options)
        preview_row = df[df["Email"] == selected_email].iloc[0]
        preview_subject = safe_format_template(subject_template, preview_row)
        preview_body = safe_format_template(body_template, preview_row)
        st.markdown(f"**Subject:** {preview_subject}")
        st.markdown(convert_bold(preview_body), unsafe_allow_html=True)

    # ========================================
    # Label & Timing Options
    # ========================================
    st.header("🏷️ Label & Timing Options")
    label_name = st.text_input("Gmail label to apply (new emails only)", value="Mail Merge Sent")
    delay = st.slider("Delay between emails (seconds)", 30, 300, 30, 5)

    # ========================================
    # Send Mode
    # ========================================
    send_mode = st.radio(
        "Choose sending mode",
        ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"]
    )

    # ========================================
    # Main Send / Draft Button
    # ========================================
    if st.button("🚀 Send Emails / Save Drafts"):
        label_id = get_or_create_label(service, label_name)
        sent_count = 0
        skipped, errors = [], []

        with st.spinner("📨 Processing emails... please wait."):
            if "ThreadId" not in df.columns:
                df["ThreadId"] = None
            if "RfcMessageId" not in df.columns:
                df["RfcMessageId"] = None

            for idx, row in df.iterrows():
                to_addr = extract_email(str(row.get("Email", "")).strip())
                if not to_addr:
                    skipped.append(row.get("Email"))
                    continue

                try:
                    subject = safe_format_template(subject_template, row)
                    body_html = convert_bold(safe_format_template(body_template, row))
                    message = MIMEText(body_html, "html")
                    message["To"] = to_addr
                    message["Subject"] = subject
                    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
                    msg_body = {"raw": raw}

                    if send_mode == "💾 Save as Draft":
                        sent_msg = service.users().drafts().create(userId="me", body={"message": msg_body}).execute().get("message", {})
                    else:
                        sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()

                    time.sleep(random.uniform(delay * 0.9, delay * 1.1))

                    msg_detail = service.users().messages().get(
                        userId="me",
                        id=sent_msg.get("id", ""),
                        format="metadata",
                        metadataHeaders=["Message-ID"],
                    ).execute()
                    headers = msg_detail.get("payload", {}).get("headers", [])
                    message_id_header = next((h["value"] for h in headers if h["name"].lower() == "message-id"), None)

                    if send_mode == "🆕 New Email" and label_id:
                        service.users().messages().modify(
                            userId="me",
                            id=sent_msg["id"],
                            body={"addLabelIds": [label_id]},
                        ).execute()

                    df.loc[idx, "ThreadId"] = sent_msg.get("threadId", "")
                    df.loc[idx, "RfcMessageId"] = message_id_header or ""
                    sent_count += 1

                except Exception as e:
                    errors.append((to_addr, str(e)))

        # ========================================
        # ✅ Backup + Download
        # ========================================
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_label = re.sub(r'[^A-Za-z0-9_-]', '_', label_name)
        file_name = f"Updated_{safe_label}_{timestamp}.csv"
        file_path = os.path.join("/tmp", file_name)

        df.to_csv(file_path, index=False)
        st.session_state["last_saved_csv"] = file_path
        st.session_state["last_saved_name"] = file_name

        st.success(f"✅ Processed {sent_count} emails. Backup saved to server.")

        # Send backup to Gmail inbox
        send_email_backup(service, "me", file_path)

        # Manual download
        st.download_button(
            label="⬇️ Download Updated CSV (Local Copy)",
            data=open(file_path, "rb"),
            file_name=file_name,
            mime="text/csv",
        )

        if skipped:
            st.warning(f"⚠️ Skipped {len(skipped)} invalid emails.")
        if errors:
            st.error(f"❌ Failed for {len(errors)} emails.")
