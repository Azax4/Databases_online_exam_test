from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import os.path
import smtplib
import csv
import sys
import io
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.remote.webelement import WebElement
from selenium.common.exceptions import NoSuchElementException
from optparse import OptionParser
import json, pdb, time, urllib, base64, hashlib, hmac, os
from googleapiclient.http import MediaIoBaseDownload


# The ID of the document you want to copy.
DOCUMENT_ID = '********'


def startup():
    creds = None
    # If modifying these scopes, delete the file token.pickle.
    DRIVE_SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/documents']

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', DRIVE_SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('docs', 'v1', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)
    return service, drive_service


def create_exam_copy(student_jhed):
    copy_title = student_jhed+"_Finals_Databases_Fall2020"
    body = {
        'name': copy_title
    }
    drive_response = drive_service.files().copy(
        fileId=DOCUMENT_ID, body=body).execute()
    document_copy_id = drive_response.get('id')
    return document_copy_id



def send_emails():
    fromaddr = 'smtp.gmail.com:587'
    ifile = open('roster-615.csv',"r")
    outfile = open('emailstatus.csv','w')
    examdetails = open('documentstatus.csv','w')
    reader = csv.reader(ifile)
    #Put in your gmail address here
    username = 'ankurkejriwal4jhu@gmail.com'
    #Password for your gmail account here
    password = '******'

    server = smtplib.SMTP('smtp.gmail.com:587')
    server.starttls()
    server.login(username,password)

    rownum = 0
    for row in reader:
        if rownum == 0:
            header = row
        else:
            if len(row)==0:
                continue
            email_addr = row[3]
            name = row[1]
            try:
                link_new = create_exam_copy(row[2])
                examdetails.write(row[2]+','+email_addr+','+link_new+','+'success\n')
            except:
                link_new = "error"
                examdetails.write(row[2]+','+email_addr+','+link_new+','+'failure\n')
                print("failure for user: " + row[2])
            link = "https://docs.google.com/document/d/"+link_new+"/edit"
            msg1 = '\nSubject: Database Fall 2020 Midterm\n'
            msg = """
            Subject: Database Fall 2020 Finals
            Hi,
            Reach out to CAs on piazza and the class zoom link.

            There is also information on the course website

            If all else fails, you can message Evan at johndoe1@jhu.edu or text ******** (if in the US). But you'll get a faster reply using the above methods of communication.

            The prof is available at  ***-***-**** or prof-email-id@jhu.edu

            The link for your exam is below

            All the best!
            """ + link
            try:
              server.sendmail(fromaddr, email_addr, msg)
              outfile.write(str(rownum)+","+email_addr+",sent\n")
            except Exception as inst:
              outfile.write(str(rownum)+","+email_addr+",failed\n")
              print("failure for user: " + email_addr)
              print(inst)
        rownum = rownum + 1
    server.quit()
    ifile.close()
    outfile.close()


x1 = []
def callback(request_id, response, exception):
    if exception:
        # Handle error
        print(exception)
    else:
        # print("Permission Id: %s" % response.get('id'))
        x1.append(response.get('id'))
    # print(x)
    return x1

def give_write_permission():
    # x = []
    batch = drive_service.new_batch_http_request(callback=callback)
    ifile = open('documentstatus.csv',"r")
    outfile = open('permission_given_status.csv','w')
    reader = csv.reader(ifile)
    y = []
    for row in reader:
            if len(row)==0:
                continue
            email_addr = row[1]
            file_id = row[2]
            y.append(file_id)
            user_permission = {
                'type': 'user',
                'role': 'writer',
                'emailAddress': email_addr
            }
            # print(email_addr)
            batch.add(drive_service.permissions().create(
                    fileId=file_id,
                    body=user_permission,
                    fields='id',
            ))            
    try:
        batch.execute()
        print("Permissions set to write")
    except:
        print("could not change permissions")
    for i in range(len(y)):
        outfile.write(y[i]+','+x1[i]+'\n')

x2 = []
def callback2(request_id, response, exception):
    if exception:
        # Handle error
        print(exception)
    else:
        # print("Permission Id: %s" % response.get('id'))
        x2.append(response.get('id'))
    # print(x2)
    return x2



def give_write_permission2():
    # x = []
    batch = drive_service.new_batch_http_request(callback=callback2)
    ifile = open('documentstatus.csv',"r")
    outfile = open('permission_given_status2.csv','w')
    reader = csv.reader(ifile)
    y = []
    for row in reader:
            if len(row)==0:
                continue
            email_addr = row[1]
            file_id = row[2]
            y.append(file_id)
            user_permission = {
                'type': 'anyone',
                'role': 'writer'
            }
            # print(email_addr)
            batch.add(drive_service.permissions().create(
                    fileId=file_id,
                    body=user_permission,
                    fields='id',
            ))            
    try:
        batch.execute()
        print("Permissions set to write")
    except:
        print("could not change permissions")
    for i in range(len(y)):
        outfile.write(y[i]+','+x2[i]+'\n')




def revoke_write_permission():
    batch = drive_service.new_batch_http_request()
    ifile = open('permission_given_status.csv',"r")
    outfile = open('permission_revoked_status.csv','w')
    y = []
    x = []
    reader = csv.reader(ifile)
    for row in reader:
            if len(row)==0:
                continue
            permission_id = row[1]
            file_id = row[0]
            y.append(file_id)
            x.append(permission_id)
            # print(permission_id)
            batch.add(drive_service.permissions().delete(
                    fileId=file_id,
                    permissionId = permission_id,
            ))            
    try:
        batch.execute()
        print("Permission to write revoked")
    except Exception as inst:
        print(inst)
    for i in range(len(x)):
        outfile.write(y[i]+','+x[i]+'\n')


def revoke_write_permission2():
    batch = drive_service.new_batch_http_request()
    ifile = open('permission_given_status2.csv',"r")
    outfile = open('permission_revoked_status2.csv','w')
    y = []
    x = []
    reader = csv.reader(ifile)
    for row in reader:
            if len(row)==0:
                continue
            permission_id = row[1]
            file_id = row[0]
            y.append(file_id)
            x.append(permission_id)
            # print(permission_id)
            batch.add(drive_service.permissions().delete(
                    fileId=file_id,
                    permissionId = permission_id,
            ))            
    try:
        batch.execute()
        print("Permission to write revoked")
    except Exception as inst:
        print(inst)
    for i in range(len(x)):
        outfile.write(y[i]+','+x[i]+'\n')



def dump():
    ifile = open('documentstatus.csv',"r")
    chromedriver = "D:\\chromedriver.exe"
    
    output_path = "C:\\Users\Ankur Kejriwal\Desktop\db_f202\downloads"
    options = webdriver.ChromeOptions()
    profile = {"plugins.plugins_list": [{"enabled": False,
                                         "name": "Chrome PDF Viewer"}],
               "download.default_directory": output_path,
               "plugins.always_open_pdf_externally": True,
               "user-data-dir":"C:\\Users\\Ankur Kejriwal\\AppData\\Local\\Google\\Chrome\\User Data\\Default",
               "download.extensions_to_open": ""}
    options.add_experimental_option("prefs", profile)
    # options.add_argument("user-data-dir=C:\\Users\\Ankur Kejriwal\\AppData\\Local\\Google\\Chrome\\User Data\\Default")
    driver = webdriver.Chrome(executable_path=chromedriver, chrome_options = options)
    reader = csv.reader(ifile)
    for row in reader:
            if len(row)==0:
                continue
            # email_addr = row[1]
            file_id = row[2]
            lnk = 'https://docs.google.com/document/d/'+file_id+'/export?format=pdf'
            driver.get(lnk)
            time.sleep(4)



service, drive_service = startup()
send_emails()
give_write_permission()
revoke_write_permission()
give_write_permission2()
revoke_write_permission2()
dump()
