import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import pyautogui
import pandas as pd
from datetime import datetime
import pyodbc
import time
import warnings
import numpy as np

warnings.filterwarnings("ignore")


username = email_address
password = password
flowhub = 'flowhub_report'
server = 'localhost\SQLEXPRESS'
table = 'master'

def blank(x):
    if not x:
        field = None
    else:
        field = x
    return field

def flowsql(servername,database):
    conn =  pyodbc.connect('Driver={ODBC Driver 17 for SQL Server};Server=' + servername + ';Database=' + database + ';Trusted_Connection=yes')
    cursor = conn.cursor()

    for row in range(1):
        cursor.execute('''
                INSERT INTO [dbo].[Flowhub] (brand,category,costPerItem,totalDiscountAmount,inventoryExpDate,itemSubtotal,location,price,productName,productType,
                    quanitySold,receiptID,strainName,transactionDate,unitOfMeasure,saleYear,saleMonth,saleWeek,saleYrWeek) 
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ''',
                (
                field1,     #brand
                field2,     #category
                field3,     #costPerItem
                field4,     #totalDiscountAmount
                field5,     #inventoryExpDate
                field6,     #itemSubtotal
                field7,     #location
                field8,     #price
                field9,     #productName
                field10,    #productType
                field11,    #quanitySold
                field12,    #receiptID
                field13,    #strainName
                field14,    #transactionDate
                field15,    #unitOfMeasure
                field16,    #saleYear
                field17,    #saleMonth
                field18,    #saleWeek
                field19,    #saleYrWeek
                )
            )

        cursor.commit()


def testing():
    print('Brand: ' + str(field1))
    print('Category: ' + field2)
    print('Cost Per Item: ' + str(field3))
    print('Total Discount Amount: ' + str(field4))
    print('Inventory Exp Date: ' + str(field5))
    print('Item Subtotal: ' + str(field6))
    print('Location: ' + field7)
    print('Price: ' + field8)
    print('Product Name: ' + field9)
    print('Product Type: ' + str(field10))
    print('Quanity Sold: ' + str(field11))
    print('Receipt ID: ' + field12)
    print('Strain Name: ' + str(field13))
    print('Transaction Date: ' + str(field14))
    print('Unit Of Measure: ' + field15)
    print('Sale Year: ' + str(field16))
    print('Sale Month: ' + str(field17))
    print('Sale Week: ' + str(field18))
    print('Sale Yr Week: ' + str(field19))
    print('\n')
    

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

def mouseMove():
    #Clicks link
    pyautogui.click(133,194,duration=0.75)
    #Clicks Download
    pyautogui.click(211,1010,duration=1.25)
    #Clicks Open File
    pyautogui.click(247,886,duration=2.00)
    #pyautogui.hotkey('win', 'up')
    #Clicks FILE in Excel
    pyautogui.click(25,48,duration=2.00)
    #Clicks SaveAs in Excel
    pyautogui.click(69,317,duration=0.75)
    #Clicks File Name and enters file name
    pyautogui.click(687,153,duration=0.75)
    pyautogui.typewrite(flowhub)
    #Saves
    pyautogui.click(1261,171,duration=0.75)
    #Close Excel
    pyautogui.click(1887,22,duration=0.75)

mouseMove()
time.sleep(5)

data = pd.read_csv(r'C:\Users\rbgonzalez\Downloads\flowhub_report.csv')   
df = pd.DataFrame(data, columns= ['Brand','Category','Cost Per Item','Total Discount Amount',
    'Inventory Exp Date','Item Subtotal','Location','Price','Product Name','Product Type','Quantity Sold',
    'Receipt ID','Strain Name','Transaction Date','Unit of Measure'])

#Converts dates to datetime datatype
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
saleYrWeek = datetime_obj.isocalendar()[1]

#Calulates how many rows are in the file for the loop
index = df.index
rows = len(index)

#Loops in Inventory Exp Date and converts dates to datetime datatype

#if rows != 0:
    #i = 0

for i in range(rows):
   

    if pd.isnull(df['Inventory Exp Date'][i]) == True:
        #df['Inventory Exp Date'][i] = None 
        pass
    else:
        date = df['Inventory Exp Date'][i]
        itest = datetime.strptime(date,'%m/%d/%Y')
        df['Inventory Exp Date'][i] = itest

    #Pulls Sale dates from Transaction Date and creates new columns
    
    df['Sale Year'] = year
    df['Sale Month'] = month
    df['Sale Week'] = day
    df['Sale Yr Week'] = saleYrWeek

    #Converts Transaction Date to datetime for SQL
    
    date = df['Transaction Date'][i]
    tdate = datetime.strptime(date,'%m/%d/%Y %H:%M')
    df['Transaction Date'][i] = tdate

    
    #Creates fields for SQL entry

    field1   = df['Brand'].values[i]
    field2   = df['Category'].values[i]
    field3   = df['Cost Per Item'].values[i]
    field4   = df['Total Discount Amount'][i]    
    field5   = df['Inventory Exp Date'].values[i]       
    field6   = df['Item Subtotal'].values[i]
    field7   = df['Location'].values[i]
    field8   = df['Price'].values[i]
    field9   = df['Product Name'].values[i]
    field10  = df['Product Type'].values[i]
    field11  = df['Quantity Sold'].values[i]
    field12  = df['Receipt ID'].values[i]
    field13  = df['Strain Name'].values[i]
    field14  = df['Transaction Date'].values[i]
    field15  = df['Unit of Measure'].values[i]
    field16  = df['Sale Year'].values[i]
    field17  = df['Sale Month'].values[i]
    field18  = df['Sale Week'].values[i]
    field19  = df['Sale Yr Week'].values[i]

    field1 = str(field1)
    field3 = format(field3,'.2f')
    field6 = format(field6,'.2f')
    field8  = format(field8,'.2f')
    field10 = str(field10)
    field11 = int(field11)
    field13 = str(field13)
    field16 = int(field16)
    field17 = int(field17)
    field18 = int(field18)
    field19 = int(field19)
    
    if np.isnan(field4) == True:
        field4 = None
    else:
        pass

    if pd.isnull(df['Inventory Exp Date'][i]) == True:
        field5 = None
    else:
        pass

    #print(i)
    #testing()
    flowsql(server,table)

df.to_excel('smokey.xlsx',index=False)
os.remove(r"C:\Users\rbgonzalez\Downloads\flowhub_report.csv")
