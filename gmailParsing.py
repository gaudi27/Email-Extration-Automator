'''Author: George Z Audi
Date: April 4th 2024'''
'''this is the code for parsing through 
emails to automate getting information from emails.'''



#TODO - bug fix UI
#TODO - change the email recognition to subject instead of sender

#libraries
import os
import email
import imaplib
#getting username and password to be able to use said email
import yaml
#pasting gmail info into excel
import excelPaster
#storing emails in txt file
import StoreEmail

def resource_path(relative_path):
    """ Get the absolute path to the resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = os.path.join(os.path.dirname(__file__))
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def start_parsing_emails():
    #a list made to extract the whole emails
    emailBody = []
    #stores the emails so it doesnt add an email to spreadsheet more than once
    INFO = []
    processed_email_ids = set()

    #opens the yaml file with username and password and uses them to log in to email
    config_path = resource_path("usernameAndPassword.yml")
    with open(config_path) as f:
        text = f.read()
    
        
    #setting username and password to variables
    info = yaml.load(text, Loader=yaml.FullLoader)
    username, password = info["user"], info["password"]
    
    #now I want to make the code ever running so it auto updates 
    while True:
    
        #imap connection to email
        imapGmail = imaplib.IMAP4_SSL('imap.gmail.com')
        #logging into gmail account
        imapGmail.login(username, password)
        #location of the emails you want parsed
        imapGmail.select('Inbox')
        
        
        
        #enter the email or contrants of the type of information you need 
        #to get from the emails
        key = "FROM"
        gmail = "gzaudi738@gmail.com"
        
        
        #gets the data from the inbox
        _, data = imapGmail.search(None, key, gmail)
        
        '''now that I have the login and the type of email I want
        I now want to extract the information from the emails'''
        
        #get the IDs of the emails that are applicable to what is needed
        getIDs = data[0].split()
        
        
        #going through the list of emails and putting them into the emailBody list
        for i in getIDs:
            if i not in processed_email_ids:
                typ, data = imapGmail.fetch(i, '(RFC822)')
                emailBody.append(data)
                processed_email_ids.add(i)  # Add email ID to processed set
        
        
        
        '''I now have all the messages but with alot of unneeded data
        I want to extract only the text'''
        
        
        for emls in emailBody[::-1]:
            for response in emls:
                if type(response) is tuple:
                    my_eml = email.message_from_bytes((response[1]))
                    #ADDING KEYS AND VALUES TO THE DICTIONARY "data" TO THEN PRINT IT TO EXCEL
                    dic = {}
                    dic["sender"] = my_eml['from']
                    dic[""] = ""
                    dic["subject"] = my_eml['subject']
                    dic[""] = ""
                    for part in my_eml.walk():  
                        #print(part.get_content_type())
                        if part.get_content_type() == 'text/plain':
                            dic["body"] = part.get_payload()
                    StoreEmail.needNewFile()
                    if StoreEmail.EmailStorage(dic) == False:
                        excelPaster.Paster(dic)
