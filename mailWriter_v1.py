'''
=====================================
|           E-Mail Writer           |
|                V1                 |
=====================================
'''
import pandas as pd
import openpyxl as openxl # Used to open and analyze the .xlsx document
import mailWriterSettings as mws # Loads user settings
#import array as ary
def logo(): #prints the splash screen
    print(" ================================================= \n |                                               | \n |     |------------------|                      | \n |     |\_              _/|                      | \n |     |  \_          _/  |                      | \n |     |    \__    __/    |     E-Mail Writer    | \n |     |       \__/       |          V1          | \n |     |                  |                      | \n |     |                  |                      | \n |     |------------------|                      | \n |                                               | \n =================================================")
    pass   
import smtplib
from email.mime.multipart import MIMEMultipart # Imports email sending library
from email.mime.text import MIMEText
logo()
print("Please configure the program in mailWriterSettings.py!")
progress = mws.status #transfers settings parameters and translates them for the program

if(progress==True):
    print("Logging into gmail server...")
server = smtplib.SMTP("smtp.gmail.com", 587) # initiates connection to the mail server
server.starttls()
email = mws.gmail #transfers settings parameters and translates them for the program
password = mws.passkey #transfers settings parameters and translates them for the program
server.login(email, password )
if(progress==True):
    print("Log in successful!")
def Emailer(text, subject, recipient):
    msg = MIMEMultipart()
    msg['From'] = email
    msg['To'] = recipient
    msg['Subject'] = subject
    msg.attach(MIMEText(text, 'plain'))
    server.send_message(msg)
    pass # Code that acts as the sender. It recieves mail contents from the generator and sends it. 
filepath = "test.xlsx"
#filepath = mws.fileLocation #transfers settings parameters and translates them for the program
#data = pd.read_excel("H:\YTVIDS\worker_dep.xlsx")
#print(data)
workbook = openxl.load_workbook(filepath) #opens the excel file at the entered file location
worksheet = workbook.active
mail_list = [] # Eindimensionales Array erstellen
for row in worksheet.iter_rows(values_only=True): # creates a two dimensional array for each seperate line in the table
    mail_list.append(list(row)) #turns the one dimensional array into a two dimensional array by adding another one dimensional array into every array slot. This also fills the array.
#example = mail_list[2][1]
#print(example)
workbook.close() #closes excel file
if(mws.mailAmount==0):
    mail_length = len(mail_list) #calculates how many emails have to be sent by analysing lenght of arrays
else:
    mail_length=mws.mailAmount + 1
for mail in range(1, mail_length):
    if(progress==True):
        print("Generating Email " + str(mail) + " out of " + str(mail_length - 1))
    adresse = mail_list[mail][int(mws.mailAdress)-1] #divides the array created with the excel file and sorts the data.
    name = mail_list[mail][int(mws.mailName)-1]
    #if(name=null)
    firma = mail_list[mail][int(mws.mailCompany)-1]
    gender = mail_list[mail][int(mws.mailGender)-1]
    if(gender == "F"):
        anrede = "Sehr geehrte Frau "
        pass
    elif (gender == "M"):
        anrede = "Sehr geehrter Herr "
        pass
    elif(gender =="all"):
        anrede = "Sehr geehrte"
        pass
    else:
        anrede = "Sehr geehrte/r Herr/Frau "
        pass
    body = anrede + name + ", \n"
    if(mws.addText==True): #these if statements check what "program" the user has selected
        body = body + mws.message
    if(mws.addSignature==True):
        body = body + "\n \n" + mws.signature
    #if(mws.addFile==True):
        
    #if(mws.addImage==True):

    #print(body)

    betreff = mws.subject
    if(progress==True):
        print("Sending Email " + str(mail) + " out of " + str(mail_length - 1))
    Emailer(body, betreff, adresse)
    pass # Code that generates the mail and passes the contents along to the sender.
if(progress==True):
    print("All Emails sent successfully!")
    print("Logging out...")
server.quit() #closes connection to the mail server
if(progress==True):
    print("Log out successful. Goodbye!")