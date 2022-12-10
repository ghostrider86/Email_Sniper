import smtplib
import time
import imaplib
import email
import traceback
import base64
import os
import pickle
#import pandas as pd
import datetime as dt
import logger
import sys

# Gmail API utils
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
# for encoding/decoding messages in base64
from base64 import urlsafe_b64decode, urlsafe_b64encode
# for dealing with attachement MIME types
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from mimetypes import guess_type as guess_mime_type

class GmailApi(object):
    """description of class"""
    def __init__(self):
        self.SMTP_Server = 'imap.gmail.com'
        self.SMTP_PORT = 993
        # Request all access (permission to read/send/receive emails, manage the inbox, and more)
        self.SCOPES = ['https://mail.google.com/']
        self.email_count = 1
        self.email_bytes = 0

    def set_email(self,email_address):
        self.FROM_EMAIL = email_address

    def set_feedback(self, disp, logger):
        self.Display = disp
        self.log = logger

    def set_save_path (self,path):
        self.save_path = path
    
    def set_copy_delete(self, is_copy, is_delete):
        self.do_copy = is_copy
        self.do_delete = is_delete

    def get_search_count(self):
        return self.email_count

    def reset_search_count(self):
        self.email_count = 1

    def get_search_size(self):
        return self.get_size_format(self.email_bytes)

    def logprint(self,msg, msg2="", msg3="", msg4="", msg5="", msg6=""):
        if not msg6 == "":
            print(msg,msg2,msg3,msg4,msg5,msg6)
            self.log.write(msg+" "+msg2+msg3+msg4+msg5+mg6)
        elif not msg5 == "":
            print(msg,msg2,msg3,msg4,msg5)
            self.log.write(msg+" "+msg2+msg3+msg4+msg5)
        elif not msg4 == "":
            print(msg,msg2,msg3,msg4)
            self.log.write(msg+" "+msg2+msg3+msg4)
        elif not msg3 == "":
            print(msg,msg2,msg3)
            self.log.write(msg+" "+msg2+msg3)
        elif not msg2 == "":
            print(msg,msg2)
            self.log.write(msg+" "+msg2)
        else :
            print(msg)
            self.log.write(msg)


    def gmail_authenticate(self,email):
        creds = None
        # the file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first time
        fname = email+".pickle"
        if os.path.exists(fname):
            with open(fname, "rb") as token:
                creds = pickle.load(token)
                self.log.write("loaded existing credentials")
        # if there are no (valid) credentials availablle, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                    self.log.write("refreshed credentials")
                except:
                    if os.path.exists(fname):
                        os.remove(fname)
                        return self.gmail_authenticate(email)
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', self.SCOPES)
                self.log.write("asking user for credentials")
                creds = flow.run_local_server(port=0)
            # save the credentials for the next run
            with open(fname, "wb") as token:
                pickle.dump(creds, token)
                self.log.write("saving credentials")
        return build('gmail', 'v1', credentials=creds)

    def get_email_pickle_name(self):
        try:
            idx = self.FROM_EMAIL.index('@')
        except:
            return "token"
        return self.clean(self.FROM_EMAIL[0:idx])
    
    def get_service(self):
        # get the Gmail API service for the active email        
        service = self.gmail_authenticate(self.get_email_pickle_name())
        return service

    def search_messages(self,service, query):
        self.Display("searching gmail...")
        self.log.write("  starting gmail search")
        result = service.users().messages().list(userId='me',q=query).execute()
        messages = [ ]
        if 'messages' in result:
            messages.extend(result['messages'])
        while 'nextPageToken' in result:
            page_token = result['nextPageToken']
            self.log.write("  fetching more messages")
            result = service.users().messages().list(userId='me',q=query, pageToken=page_token).execute()
            if 'messages' in result:
                messages.extend(result['messages'])
        return messages

    def delete_messages(self,service, query):
        messages_to_delete = self.search_messages(service, query)
        # it's possible to delete a single message with the delete API, like this:
        # service.users().messages().delete(userId='me', id=msg['id'])
        # but it's also possible to delete all the selected messages with one query, batchDelete
        try:
            return service.users().messages().batchDelete(
              userId='me',
              body={
                  'ids': [ msg['id'] for msg in messages_to_delete]
              }
            ).execute()
        except Exception as e:
            self.logprint(str(e))


    # utility functions
    def get_size_format(self,b, factor=1024, suffix="B"):
        """
        Scale bytes to its proper byte format
        e.g:
            1253656 => '1.20MB'
            1253656678 => '1.17GB'
        """
        for unit in ["", "K", "M", "G", "T", "P", "E", "Z"]:

            if b < factor:
                return f"{b:.2f}{unit}{suffix}"
            b /= factor
        return f"{b:.2f}Y{suffix}"

    def safe_timestamp(self, raw):
        try:
            dto = dt.datetime.strptime(raw,'%a, %d %b %Y %H:%M:%S %z')
            return dt.datetime.strftime(dto,'%Y%b%dT%H.%M.%S')
        except:
            try:
                idx = raw.index('(')
                cooked = raw[0:idx]
                try:
                    dto = dt.datetime.strptime(cooked,'%a, %d %b %Y %H:%M:%S %z')
                    return dt.datetime.strftime(dto,'%Y%b%dT%H.%M.%S')
                except:
                    return "_date_"
            except:
                try:
                    idx = raw.index("GMT")
                    cooked = raw[0:idx]
                    try:
                        dto = dt.datetime.strptime(cooked,'%a, %d %b %Y %H:%M:%S')
                        return dt.datetime.strftime(dto,'%Y%b%dT%H.%M.%S')
                    except:
                        return "_date_"
                except:
                    return "_date_"

    def safe_filename(self, raw):
        try:
            idx1 = raw.index('<')
            idx2 = raw.index('>')
            cooked = raw[idx1+1:idx2]
            return cooked
        except:
            return raw

    def parse_parts(self, service, parts, folder_name, message_id, target):
        """
        Utility function that parses the content of an email partition
        """
        target = self.clean(target)
        if parts:
            for part in parts:
                filename = part.get("filename")
                mimeType = part.get("mimeType")
                body = part.get("body")
                data = body.get("data")
                file_size = body.get("size")
                self.email_bytes += file_size
                part_headers = part.get("headers")
                if part.get("parts"):
                    # recursively call this function when we see that a part
                    # has parts inside
                    self.parse_parts(service, part.get("parts"), folder_name, message_id, target)
                if mimeType == "text/plain":
                    # if the email part is text plain
                    if data:

                        self.Display(" --> text file:" + target+".txt")
                        if self.do_copy:
                            rawtext = urlsafe_b64decode(data)
                            self.logprint(rawtext.decode())

                            try:
                                filepath = os.path.join(folder_name, target+".txt")
                                with open(filepath, "wb") as f:
                                    f.write(rawtext)
                            except:
                                curr_dir = os.getcwd()
                                cdr = os.chdir(folder_name)
                                try:
                                    with open(target+".txt", "wb") as f:
                                        f.write(rawtext)
                                except:
                                    with open("message.txt", "wb") as f:
                                        f.write(rawtext)
                                cdr = os.chdir(curr_dir)
                elif mimeType == "text/html":
                    # if the email part is an HTML content
                    # save the HTML file and optionally open it in the browser
                    if not filename:
                        filename = "index.html"
                    self.Display(" --> html file:" + filename)
                    if self.do_copy:
                        try:
                            filepath = os.path.join(folder_name, target+'_'+filename)
                            self.logprint("Saving HTML to ", filepath)
                            with open(filepath, "wb") as f:
                                f.write(urlsafe_b64decode(data))
                        except:
                            curr_dir = os.getcwd()
                            cdr = os.chdir(folder_name)
                            self.logprint("Switched to ", target+'_'+filename)
                            try:
                                with open(target+'_'+filename, "wb") as f:
                                    f.write(urlsafe_b64decode(data))
                            except:
                                self.logprint("Switched to ", filename)
                                with open(filename, "wb") as f:
                                    f.write(urlsafe_b64decode(data))
                            cdr = os.chdir(curr_dir)
                else:
                    
                    # attachment other than a plain text or HTML
                    for part_header in part_headers:
                        part_header_name = part_header.get("name")
                        part_header_value = part_header.get("value")
                        if part_header_name == "Content-Disposition":
                            if "attachment" in part_header_value:
                                self.logprint ("handling mimeType:",mimeType)
                                # we get the attachment ID
                                # and make another request to get the attachment itself
                                self.Display(" --> file:" + filename)
                                if self.do_copy:
                                    self.logprint("Saving the file:", target + '_' + filename, "size:", self.get_size_format(file_size))
                                    attachment_id = body.get("attachmentId")
                                    attachment = service.users().messages() \
                                                .attachments().get(id=attachment_id, userId='me', messageId=message_id['id']).execute()
                                    data = attachment.get("data")
                                    try:
                                        filepath = os.path.join(folder_name, target + '_' + filename)
                                        if data:
                                            with open(filepath, "wb") as f:
                                                f.write(urlsafe_b64decode(data))

                                    except:
                                        curr_dir = os.getcwd()
                                        cdr = os.chdir(folder_name)
                                        if data:
                                            self.logprint("Switched to ",target+'_'+filename)
                                            try:
                                                with open(target + '_' + filename, "wb") as f:
                                                    f.write(urlsafe_b64decode(data))
                                            except:
                                                try:
                                                    self.logprint("Switched to ",filename)
                                                    with open(filename, "wb") as f:
                                                        f.write(urlsafe_b64decode(data))
                                                except:
                                                    self.logprint("  switched to f.dat")
                                                    with open("f.dat", "wb") as f:
                                                        f.write(urlsafe_b64decode(data))
                                        cdr = os.chdir(curr_dir)


    def read_message(self, service, message_id):
        """
        This function takes Gmail API `service` and the given `message_id` and does the following:
            - Downloads the content of the email
            - Prints email basic information (To, From, Subject & Date) and plain/text parts
            - Creates a folder for each email based on the subject
            - Downloads text/html content (if available) and saves it under the folder created as index.html
            - Downloads any file that is attached to the email and saves it in the folder created
        """
        self.email_count += 1
        msg = service.users().messages().get(userId='me', id=message_id['id'], format='full').execute()
        # parts can be the message body, or attachments
        payload = msg['payload']
        headers = payload.get("headers")
        parts = payload.get("parts")
        folder_name = self.save_path #path "email"
        target = 'gm_'
        if not os.path.isdir(folder_name):
            os.mkdir(folder_name)
        if headers:
            # this section prints email basic info & creates a folder for the email
            for header in headers:
                name = header.get("name")
                value = header.get("value")
                if name == 'Delivered-To' or name == 'delivered-to':
                    # we print the delivered-to address
                    self.logprint("Delivered-To:", value)
                    #self.Display("Delivered-To: "+value)
                elif name == 'Content-Type':
                    self.logprint(name+":",value)
                elif name == 'From' or name == 'from':
                    # we print the From address
                    self.logprint("From:", value)
                    #self.Display("From: "+value)
                    target += self.safe_filename(value)
                elif name == "To" or name == "to" or name == 'to':
                    # we print the To address
                    self.logprint("To:", value)
                    #self.Display("To: "+value)
                elif name == "Subject" or name == 'subject':
                    # make a directory with the name of the subject
                    folder_name = self.save_path + os.sep +self.clean(value)
                    if len(folder_name) > 255:
                        folder_name = folder_name[0:254]

                    # we will also handle emails with the same subject name
                    folder_counter = 0
                    while os.path.isdir(folder_name):
                        folder_counter += 1
                        # we have the same folder name, add a number next to it
                        if folder_name[-1].isdigit() and folder_name[-2] == "_":
                            folder_name = f"{folder_name[:-2]}_{folder_counter}"
                        elif folder_name[-2:].isdigit() and folder_name[-3] == "_":
                            folder_name = f"{folder_name[:-3]}_{folder_counter}"
                        else:
                            folder_name = f"{folder_name}_{folder_counter}"
                    if self.do_copy:
                        os.mkdir(folder_name)
                    
                    self.logprint("Subject:", value)
                    #self.Display(" subject: "+value+"  folder:"+folder_name)
                    #
                    #target += '_' + value
                elif name == "Date" or name == 'date':
                    # we print the date when the message was sent
                    self.logprint("Date:", value)
                    #self.Display(" date: "+value)
                    target += '_' + self.safe_timestamp(value)  
                #else:
                #    print(name+":", value)

        self.parse_parts(service, parts, folder_name, message_id,self.clean(target))

    def clean(self, text):
        # clean text for creating a folder
        step1 = "".join(c if c.isalnum() else "_" for c in text)
        step2 = step1.replace("__","_")
        step1 = step2.replace("__","_")
        step2 = step1.replace("__","_")
        if (len(step2) > 255):
            return step2[0:254]
        return step2
