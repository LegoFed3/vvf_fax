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

import win32print
import win32api

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://mail.google.com/']

ALLOWED_MIMETYPES = ['application/pdf']


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
        # Configure default printer
        printer_name = win32print.GetDefaultPrinter()  # verify that it matches with the name of your printer
        print_defaults = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}  # Doesn't work with PRINTER_ACCESS_USE
        handle = win32print.OpenPrinter(printer_name, print_defaults)
        level = 2
        attributes = win32print.GetPrinter(handle, level)
        # attributes['pDevMode'].Duplex = 1  #no flip
        # attributes['pDevMode'].Duplex = 2  #flip up
        attributes['pDevMode'].Duplex = 3  # flip over
        win32print.SetPrinter(handle, level, attributes, 0)
        # win32print.GetPrinter(handle, level)['pDevMode'].Duplex

        # Configure data to set message unread
        post_data = {
          "addLabelIds": [],
          "removeLabelIds": ['UNREAD']
        }

        # Access gmail
        service = build('gmail', 'v1', credentials=creds)
        # Get labelid of 'fax'
        fax_lableid = 'Label_4140688616999534095'
        # results = service.users().labels().list(userId='me').execute()
        # labels = results.get('labels', [])
        # print(labels)

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
                payload = msg_obj['payload']
                attachment_ids = set()
                for part in payload['parts']:
                    if 'attachmentId' in part['body'] and part['mimeType'] in ALLOWED_MIMETYPES:
                        attachment_ids.add((part['body']['attachmentId'], part['filename']))
                log.info(f"Found {len(attachment_ids)} attachments to print")
                for atc_id, filename in attachment_ids:
                    atc_obj = service.users().messages().attachments().get(userId='me', messageId=msg['id'],
                                                                           id=atc_id).execute()
                    file_data = base64.urlsafe_b64decode(atc_obj['data'].encode('UTF-8'))
                    log.info(f"Printing file {filename}...")
                    num_files += 1
                    ext = filename.split(".")[-1]
                    with open(f"./tmp.{ext}", "wb") as fo:
                        fo.write(file_data)

                        # win32api.ShellExecute(0, 'print', 'manual1.pdf', '.', '/manualstoprint', 0)
                        win32api.ShellExecute(0, 'print', f"tmp.{ext}", None, '.', 0)
                        time.sleep(30)
                        # os.remove(f"./tmp.{ext}")
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
