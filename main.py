from __future__ import print_function

import os.path
import base64
import time
import logging as log
import sys

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from apiclient import errors

import win32api

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://mail.google.com/']

# ALLOWED_MIMETYPES = ['application/pdf']
ALLOWED_EXTENSIONS = ['pdf']


def get_attachments(message, service, user_id, store_dir="attachments/"):
    try:
        attachments = []
        parts = [message['payload']]
        while parts:
            part = parts.pop()
            if part.get('parts'):
                parts.extend(part['parts'])
            if part.get('filename'):
                if 'data' in part['body']:
                    file_data = base64.urlsafe_b64decode(part['body']['data'].encode('UTF-8'))
                    print('FileData for %s, %s found! size: %s' % (message['id'], part['filename'], part['size']))
                elif 'attachmentId' in part['body']:
                    attachment = service.users().messages().attachments().get(
                        userId=user_id, messageId=message['id'], id=part['body']['attachmentId']
                    ).execute()
                    file_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))
                    print('FileData for %s, %s found! size: %s' % (message['id'], part['filename'], attachment['size']))
                else:
                    file_data = None
                if file_data:
                    path = ''.join([store_dir, part['filename']])
                    with open(path, 'wb') as f:
                        f.write(file_data)
                    attachments.append(part['filename'])
        return attachments
    except errors.HttpError as error:
        print('An error occurred: %s' % error)
        return []


def main():
    log.info("VVF FAX starting...")

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        # Configure data to set message unread
        post_data = {
          "addLabelIds": [],
          "removeLabelIds": ['UNREAD']
        }

        # Access gmail
        service = build('gmail', 'v1', credentials=creds)
        # Get labelid of 'fax'
        fax_lableid = 'Label_4140688616999534095'

        num_files = 0

        # Load all messages with label 'fax' that are unread
        results = service.users().messages().list(userId='me', labelIds=[fax_lableid],
                                                  q='is:unread has:attachment').execute()

        if not results or results['resultSizeEstimate'] == 0:
            log.info(f"No new messages found, stopping...")
        else:
            log.info(f"Found {results['resultSizeEstimate']} new messages")
            for msg in results['messages']:
                msg_obj = service.users().messages().get(userId='me', id=msg['id']).execute()
                log.info(f"Processing message '{msg_obj['snippet']}'...")

                attachments = get_attachments(msg_obj, service, 'me')
                log.info(f"Found {len(attachments)} attachments to print")

                for filename in attachments:
                    ext = filename.split(".")[-1].lower()
                    if ext not in ALLOWED_EXTENSIONS:
                        log.info(f'Skipping {filename} because \'{ext}\' is not an allowed extension')
                    else:
                        log.info(f"Printing file {filename}...")

                        win32api.ShellExecute(0, 'print', f"{filename}", None, './attachments', 0)
                        time.sleep(30)

                # Mark message as read
                log.info("Done, marking message as read...")
                service.users().messages().modify(userId='me', id=msg['id'], body=post_data).execute()

        log.info(f"Run finished successfully, {num_files} files sent to printer")

    except HttpError as error:
        log.error(f'An error occurred: {error}')


if __name__ == '__main__':
    log.basicConfig(filename='vvf_fax.log', encoding='utf-8', level=log.DEBUG, format='%(asctime)s %(message)s')
    log.getLogger().addHandler(log.StreamHandler(sys.stdout))
    main()
