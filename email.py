import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import pyautogui
import pandas as pd
from datetime import datetime
import time
import fsspec



username = "smokey@redbudroots.com"
password = "Rep0rt$&!411"
flowhub = 'flowhub_report'

def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)

imap = imaplib.IMAP4_SSL("outlook.office.com",993)
imap.login(username,password)
#Selects which folder to access
status,messages = imap.select('Flowhub')

N = 1

messages = int(messages[0])

for i in range(messages, messages-N, -1):
    # fetch the email message by ID
    res, msg = imap.fetch(str(i), "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            # parse a bytes email into a message object
            msg = email.message_from_bytes(response[1])
            # decode the email subject
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                # if it's a bytes, decode to str
                subject = subject.decode(encoding)
            # decode email sender
            From, encoding = decode_header(msg.get("From"))[0]
            if isinstance(From, bytes):
                From = From.decode(encoding)
            print("Subject:", subject)
            print("From:", From)
            # if the email message is multipart
            if msg.is_multipart():
                # iterate over email parts
                for part in msg.walk():
                    # extract content type of email
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    try:
                        # get the email body
                        body = part.get_payload(decode=True).decode()
                    except:
                        pass
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        # print text/plain emails and skip attachments
                        print(body)
                    elif "attachment" in content_disposition:
                        # download attachment
                        filename = part.get_filename()
                        if filename:
                            folder_name = clean(subject)
                            if not os.path.isdir(folder_name):
                                # make a folder for this email (named after the subject)
                                os.mkdir(folder_name)
                            filepath = os.path.join(folder_name, filename)
                            # download attachment and save it
                            open(filepath, "wb").write(part.get_payload(decode=True))
            else:
                # extract content type of email
                content_type = msg.get_content_type()
                # get the email body
                body = msg.get_payload(decode=True).decode()
                if content_type == "text/plain":
                    # print only text email parts
                    print(body)
            if content_type == "text/html":
                # if it's HTML, create a new HTML file and open it in browser
                folder_name = clean(subject)
                if not os.path.isdir(folder_name):
                    # make a folder for this email (named after the subject)
                    os.mkdir(folder_name)
                filename = "index.html"
                filepath = os.path.join(folder_name, filename)
                # write the file
                open(filepath, "w").write(body)
                # open in the default browser
                webbrowser.open(filepath)
            print("="*100)
# close the connection and logout
imap.close()
imap.logout()

time.sleep(3)

#Clicks link
pyautogui.click(153,246,duration=0.75)
#Clicks Download
pyautogui.click(275,992,duration=1.25)
#Clicks Open File
pyautogui.click(320,844,duration=2.00)

#pyautogui.click(1020,1050,duration=2.00)
#pyautogui.hotkey('win', 'up')



#Clicks FILE in Excel
pyautogui.click(33,66,duration=2.00)
#Clicks SaveAs in Excel
pyautogui.click(91,402,duration=0.75)
#Clicks File Name and enters file name
pyautogui.click(702,190,duration=0.75)
pyautogui.typewrite(flowhub)
#Saves
pyautogui.click(1581,222,duration=0.75)
time.sleep(3)
file = pd.read_csv ("C://Users//Gonzalez//Downloads//flowhub_report.csv")   
df = pd.DataFrame(file, columns= ['Brand','Category','Cost Per Item','Total Discount Amount',
    'Inventory Exp Date','Item Subtotal','Location','Price','Product Name','Product Type','Quantity Sold', 'Price',
    'Receipt ID','Strain Name','Transaction Date','Unit of Measure'])

date = df['Transaction Date'][0]
datetime_obj = datetime.strptime(date,'%m/%d/%Y %H:%M')

#Sale Year
year = datetime_obj.year
#Sale month
month = datetime_obj.month
#Sale Day
day = datetime_obj.day
#Sale week
#Sale Yr Week
datetime_obj.isocalendar()[1]

for i in df['Transaction Date']:
    datetime_obj = datetime.strptime(i,'%m/%d/%Y %H:%M')
    month = datetime_obj.month
    day = datetime_obj.day
    year = datetime_obj.year
       
    week_number = datetime_obj.isocalendar()[1]
    
    df['Sale Year'] = year
    df['Sale Month'] = month
    df['Sale Yr Week'] = week_number

print(df)
df.to_excel('smokey.xlsx',index=False) 
