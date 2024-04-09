'''Author: George Z Audi
Date: April 4th 2024'''
'''this is the code for parsing through 
emails to automate getting information from emails.'''


#libraries
import email
import imaplib
#getting username and password to be able to use said email
import yaml

#opens the yaml file with username and password and uses them to log in to email
with open("usernameAndPassword.yml") as f:
    text = f.read()
    
#setting username and password to variables
info = yaml.load(text, Loader=yaml.FullLoader)
username, password = info["user"], info["password"]


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

#a list made to extract the whole emails
emailBody = []
#going through the list of emails and putting them into the emailBody list
for i in getIDs:
    typ, data = imapGmail.fetch(i, '(RFC822)')
    emailBody.append(data)
    
'''I now have all the messages but with alot of unneeded data
I want to extract only the text'''

for emls in emailBody[::-1]:
    for response in emls:
        if type(response) is tuple:
            my_eml = email.message_from_bytes((response[1]))
            print("_________________________________________")
            print ("subj:", my_eml['subject'])
            print ("from:", my_eml['from'])
            print ("body:")
            for part in my_eml.walk():  
                #print(part.get_content_type())
                if part.get_content_type() == 'text/plain':
                    print (part.get_payload())
