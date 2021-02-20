from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from email.mime.text import MIMEText
import base64
import re
import logging
import sys

root_logger = logging.getLogger()
root_logger.setLevel(logging.DEBUG)
handler = logging.FileHandler('guests_mailer.log', 'w', 'utf-8')
handler.setFormatter(logging.Formatter('[%(asctime)s] %(name)s - %(message)s', '%Y-%m-%d %H:%M:%S'))
root_logger.addHandler(handler)

# If modifying these scopes, delete the file token.pickle.
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.modify',
]

def build_gmail_client():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('gmail', 'v1', credentials=creds, cache_discovery=False)

def reply_to_guests():
    try:
        """
        Finds all unread emails and reply to each guest with a canned response
        """
        gmail = build_gmail_client()

        unread_guest_inquiries = gmail.users().messages().list(
            userId='me',
            labelIds=['Label_4003694155315192171'],
            maxResults=50,
            q='is:unread'
        ).execute()

        message_ids = [msg['id'] for msg in unread_guest_inquiries['messages']]

        if len(message_ids) == 0:
            logging.info('No new emails')
            return None

        for message_id in message_ids:
            logging.debug('Processing message %s', message_id)
            message_payload = gmail.users().messages().get(userId='me', id=message_id).execute()['payload']

            if 'parts' in message_payload:
                raw_email_body = base64.urlsafe_b64decode(message_payload['parts'][0]['body']['data']).decode('utf-8')
                email_regex = '([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)'
                guest_email_address = re.search(email_regex, raw_email_body).group(1)
            else:
                guest_email_address = next(header for header in message_payload['headers'] if header['name'] == 'Reply-To')['value']

            if guest_email_address == 'membership@toastmasters.org':
                logging.error('Could not find guest email address for message %s', message_id)
                return None

            body_message = """Hi there!

This is Buddies' Vice President of Membership. I just received your contact information and am excited to help you experience our community.

We'd love to have you join us as a guest and hear more from you before, during or after the meeting. Let me share the meeting info with you below for access, and hope to see you in our next meeting!

We meet every Saturday 1PM UTC on Zoom: https://us02web.zoom.us/j/591440460?pwd=R3pMT0hzTU1TYytWT2lvdUZkd3BmUT09 The password is 'Buddies' and this link should work for every future meeting. This link can indicate what 1PM UTC means in your timezone: https://www.starts-at.com/event/3147989528

You can learn more about us in our main website: http://buddies.toastmost.org
or Twitter: https://twitter.com/BuddiesOnlineTM)
and Facebook! https://www.facebook.com/BuddiesOnlineToastmastersClub

🌻

Yours,
Vice President of Membership
"""

            message = MIMEText(body_message)
            message['to'] = guest_email_address
            message['from'] = 'buddiestmi@gmail.com'
            message['subject'] = 'We\'re Happy You Want to Join!'

            raw_message = { 'raw': base64.urlsafe_b64encode(message.as_string().encode('ascii')).decode('utf-8') }
            try:
                logging.debug('Sending email to "%s"', guest_email_address)
                gmail.users().messages().send(userId='me', body=raw_message).execute()

                logging.debug('Marking message: %s as read', message_id)
                gmail.users().messages().modify(userId='me', id=message_id, body={'removeLabelIds': ['UNREAD']}).execute()
                logging.error('Sent successfully!')
            except Exception as e_sending:
                logging.error('Unexpected error: %s', str(e_sending))
    except Exception as e:
        logging.error('Unexpected error: %s', str(e))

