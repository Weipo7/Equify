import datetime
import email
import imaplib
import mailbox
import os
import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


user = encrypted
password = encrypted


mail = imaplib.IMAP4_SSL('imap.gmail.com')
mail.login(user, password)
mail.list()
mail.select('inbox')
result, data = mail.uid('search', None, "UNSEEN") # (ALL/UNSEEN)
i = len(data[0].split())
for x in range(i):
    latest_email_uid = data[0].split()[x]
    result, email_data = mail.uid('fetch', latest_email_uid, '(RFC822)')
    raw_email = email_data[0][1]
    raw_email_string = raw_email.decode('utf-8')
    email_message = email.message_from_string(raw_email_string)

    # Header Details
    date_tuple = email.utils.parsedate_tz(email_message['Date'])
    if date_tuple:
        local_date = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
        local_message_date = "%s" %(str(local_date.strftime("%a, %d %b %Y %H:%M:%S")))
    email_from = str(email.header.make_header(email.header.decode_header(email_message['From'])))
    email_to = str(email.header.make_header(email.header.decode_header(email_message['To'])))
    subject = str(email.header.make_header(email.header.decode_header(email_message['Subject'])))

    for part in email_message.walk():
        if part.get_content_type() == "text/plain":
            body = part.get_payload(decode=True)
            file_name = "email_" + str(x) + ".txt"
            output_file = open(file_name, 'w')
            output_file.write("From: %s\nTo: %s\nDate: %s\nSubject: %s\n\nBody: \n\n%s" %(email_from, email_to,local_message_date, subject, body.decode('utf-8')))
            output_file.close()
        else:
            continue

folder_path = r'C:\Users\btbnn\Documents\Equify-Backend'
files = os.listdir(folder_path)
num_files=(len(files))

def automated_email(first, last, mail, interest):

    subject = first+" "+last+"'s Equify Document"
    body = "Hello "+first+"! \nBelow is your customized Equify Opportunity Document"

    message = MIMEMultipart()
    message["From"] = user
    message["To"] = mail
    message["Subject"] = subject
    
    # Add body to email
    message.attach(MIMEText(body, "plain"))
    if(interest=="Virtualreality(VR)"):
        filename = "Virtual-Reality-Equify.pdf"  # In same directory as script
    
    elif(interest=="ArtificialIntelligence"):
        filename = "Artificial-Intelligence-Equify.pdf"
    
    elif(interest=="IoT"):
        filename = "IoT-Equify.pdf"
    
    elif(interest=="BlockChain"):
        filename = "Blockchain-Equify.pdf"
    
    else:
        filename="Quantum-Computing-Equify.pdf"
    print()
    # Open PDF file in binary mode
    with open(filename, "rb") as attachment:

        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
       
    encoders.encode_base64(part)
    
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )
    
    message.attach(part)
    text = message.as_string()

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(user, password)
        server.sendmail(user, mail, text)
        
for i in range(0,(num_files)-8):
    text_file="email_"+str(i)+".txt"
    handle=open(text_file,"r")
    form=handle.read()
    form=form.replace(" ","")
    form=form.replace("\n","")
    if (form.find("<noreply@123formbuilder.com>")!=-1):
        first_name=form[form.find("Subject:")+8:form.find("-", form.find("Subject:"))]
        last_name=form[form.find("-", form.find("Subject:"))+1:form.find("Body", form.find("-", form.find("Subject:")))]
        email=form[form.find("Email",form.find("Body", form.find("-", form.find("Subject:"))))+5:form.find("Grade", form.find("Email",form.find("Body", form.find("-", form.find("Subject:")))))]
        interest=form[form.find("Interest",form.find("Grade", form.find("Email",form.find("Body", form.find("-", form.find("Subject:"))))))+8:form.find("Message", form.find("Interest",form.find("Grade", form.find("Email",form.find("Body", form.find("-", form.find("Subject:")))))))]
        handle.close()
        os.remove(text_file)
        automated_email(first_name, last_name, email, interest)
    else:
        handle.close()
        os.remove(text_file)
        
 
