import json
from pathlib import Path
import os
import re
import sys
import time
from httplib2 import Http
import selenium
import pickle
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from unidecode import unidecode
from whoswho import who
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2 import service_account
from googleapiclient.discovery import build

APPLICATION_NAME = 'whatsapp_automation'
MAXLOGFILE = 10000*1024 
base_path = os.getcwd()

# Global Options
LOADWAIT15 = 15
LOADWAIT30 = 30
LOADWAIT60 = 60
STEPWAIT = 2
CONFIRMATIONWAIT = 5
LOADWAIT10 = 10
PAGE_WAIT_TIME = 5

# Initialize Logging
fname = os.path.join(base_path, APPLICATION_NAME + ".log")
if os.path.isfile(fname) and os.stat(fname).st_size > MAXLOGFILE:
    newfpath = os.path.join(base_path,"BackupLogs/")
    timestr = time.strftime("%Y%m%d-%H%M%S")
    os.replace(fname, newfpath + APPLICATION_NAME + "_" + timestr + ".log")

timestr = time.strftime("%Y-%m-%d %H:%M:%S")
logFile = open(fname,"a")
logFile.write(APPLICATION_NAME + " Session Started at: " + timestr)
logFile.write('\n')
logFile.close()

def CustomLog(logtext,type):
    global logFile
    global fname

    if os.path.isfile(fname) and os.stat(fname).st_size > MAXLOGFILE:
        newfpath = os.path.join(base_path,"BackupLogs/")
        Path(newfpath).mkdir(parents=True, exist_ok=True)
        timestr = time.strftime("%Y%m%d-%H%M%S")
        os.replace(fname, newfpath + APPLICATION_NAME + "_" + timestr + ".log")

    logFile = open(fname,"a")

    timestr = time.strftime("%Y-%m-%d-%H:%M:%S")
    strtext = timestr + " " + type + ": " + logtext

    print(strtext)
    logFile.write(strtext)
    logFile.write('\n')
    logFile.close()

userdataDir = os.path.join(base_path, "UserData")
userProfile = "Default"
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--user-data-dir=" + userdataDir)
chrome_options.add_argument("--profile-directory=" + userProfile)
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

messages = [
    ", Message 1",
    ", Message 2",
    ", Message 3"
]

operation_type = input("Input Question\n")
SCOPES = ['https://www.googleapis.com/auth/drive','https://www.googleapis.com/auth/spreadsheets']
sheet_id = 'Sheet ID'
range =  ['Sheet Name!A:G', 'Sheet Name With!A:G', 'Sheet Name!A:G']

try:
    creds = None
    token_pickle = os.path.join(base_path,"token.pickle")
    if os.path.exists(token_pickle):
        with open(token_pickle, 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            client_secret = os.path.join(base_path,"client_secret.json")
            flow = InstalledAppFlow.from_client_secrets_file(client_secret, SCOPES) # here enter the name of your downloaded JSON file
            creds = flow.run_local_server(port=8080)
        with open(token_pickle, 'wb') as token:
            pickle.dump(creds, token)

    try:
        service = build('sheets', 'v4', credentials=creds)
    except:
        DISCOVERY_SERVICE_URL = 'Paste Your Sheet Url Here'
        service = build('sheets', 'v4', credentials=creds, discoveryServiceUrl=DISCOVERY_SERVICE_URL)

    def update_google_sheet(idx, status, contact, r):
        try:
            r = r.split("!")[0]
            idx = idx + 2
            new_contact_row = [None, None, None, None, None, None, None]
            
            new_contact_row[4] = status
            new_contact_row[5] = str(datetime.now())

            body = {'values': [new_contact_row]}
            service.spreadsheets().values().update(spreadsheetId=sheet_id, range=r + "!A" + str(idx) + ':G' + str(idx), valueInputOption='RAW', body=body).execute()

            logtext = "Updated Google Sheet against the number: " + str(contact['phone'])
            CustomLog(logtext,"INFO")
        
        except Exception as e: 
            logtext = "Failed to update Google Sheet against the number: " + str(contact['phone']) + ' - ' + str(e)
            CustomLog(logtext,"ERROR")

    try:

        if operation_type == '1':
            r = range[0]

        elif operation_type == '2':
            r = range[1]
            
        elif operation_type == '3':
            r = range[2]
        
        else:
            raise ()

        try:
            driver = webdriver.Chrome(chrome_options=chrome_options)
            driver.get('https://web.whatsapp.com/')

            try:
                result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=r).execute()
                sheet_rows = result.get('values', [])

                header = sheet_rows[0]

                try:
                    WebDriverWait(driver, LOADWAIT60).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/div[3]/header/div[2]/div/span/div[3]/div/span')))
                    
                    try:
                        logtext = "WhatsApp Web logged in"
                        CustomLog(logtext,"INFO")

                        time.sleep(STEPWAIT)
                        WebDriverWait(driver, LOADWAIT30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="side"]/div[1]/div/div/div[2]')))
                        elm_search = driver.find_element(By.XPATH, '//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]')

                        for idx, contact_row in enumerate(sheet_rows[1:]):
                            try:
                                status_check = contact_row[4]
                                if contact_row[4] == "":
                                    raise()
                            
                            except:
                                contact = dict(zip(header, contact_row))
                                try:
                                    # search_contact
                                    time.sleep(STEPWAIT)
                                    elm_search.clear()
                                    elm_search.send_keys(re.sub('[^0-9+]', '', str(contact['phone'])), Keys.ENTER)
                                    time.sleep(CONFIRMATIONWAIT)
                                    
                                    try:
                                        elm_no_chats = WebDriverWait(driver, LOADWAIT15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="pane-side"]/div[1]/div/span')))
                                        elm_no_chats = elm_no_chats.text

                                        if elm_no_chats == 'No chats, contacts or messages found':
                                            logtext = "No contact found against the number: " + str(contact['phone'])
                                            CustomLog(logtext,"INFO")

                                            update_google_sheet(idx, logtext, contact, r)

                                        else:
                                            raise()
                                    
                                    except:
                                        try:
                                            # search and click contact
                                            elm_contact_name = WebDriverWait(driver, LOADWAIT15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main"]/header/div[2]/div/div/span')))
                                            elm_contact_name = elm_contact_name.text

                                            if who.match(str(elm_contact_name).lower(), str(contact["name"]).lower()):
                                                
                                                message_text = "Hi " + str(contact["name"]) + messages[int(operation_type) - 1]

                                                elm_message = WebDriverWait(driver, LOADWAIT15).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p')))
                                                elm_message.clear()
                                                elm_message.send_keys(message_text, Keys.ENTER)
                                                time.sleep(STEPWAIT)

                                                logtext = "Message sent successfully to : " + str(contact['name']) + "\nMessage: " + message_text
                                                CustomLog(logtext,"INFO")

                                                update_google_sheet(idx, logtext, contact, r)
                                            
                                            else:
                                                logtext = "No contact found against the name: " + str(contact['name'])
                                                CustomLog(logtext,"INFO")
                                        
                                        except Exception as e:
                                            logtext = "WhatsApp failed for Contact: " + str(contact['name']) + " Number: " + str(contact['phone']) + " - " + str(e)
                                            CustomLog(logtext,"ERROR")
                                    
                                except:
                                    logtext = "Contact is invalid"
                                    CustomLog(logtext,"ERROR")

                    except Exception as e:
                        logtext = "Failed to interact with search bar"
                        CustomLog(logtext,"ERROR")

                except:
                    WebDriverWait(driver, LOADWAIT60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div/div/div[3]/div[1]/div/div/div[2]/div/canvas')))
                    
                    logtext = "WhatsApp Web is not logged in waiting on Admin to log in - Script will need to be restarted after logging in"
                    CustomLog(logtext,"INFO")

                    WebDriverWait(driver, 900).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/div[3]/header/div[2]/div/span/div[3]/div/span')))
                    driver.quit()

            except Exception as e:
                logtext = "Connection failed with Google Sheet - " + str(e)
                CustomLog(logtext,"ERROR")
                
        except Exception as e:
            logtext = "Failed to launch Whatsapp Web - " + str(e)
            CustomLog(logtext,"ERROR")

    except:
        logtext = "to run the script you need to pass the operation type as argument - 1: Follow Up, 2: Reminder, 3: Follow Up Reminder"
        CustomLog(logtext,"ERROR")

except Exception as e:
    logtext = "Error occured while reading GCP credentials - " + str(e)
    CustomLog(logtext,"ERROR")

if sys.platform == "win32":
    input("Script ended - Press any key to close this windows")

























































