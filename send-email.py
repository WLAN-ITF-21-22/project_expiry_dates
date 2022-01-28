# send e-mail with python
import smtplib
import datetime as dt
import time
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders

# create the email content
message = """Geachte,
In de bijlage kan u het document vinden met de overschot van deze week.
"""

sender = 'supermarkt.test@gmail.com'
password = '2ccs02AH3'
# put the email of the receiver here
receiver = 'vzwsupermarkt.test@gmail.com'

msg = MIMEMultipart()
msg['From'] = sender
msg['To'] = receiver
msg['Subject'] = 'Briefing'

# send the email via Gmail server

server = smtplib.SMTP('smtp.gmail.com:587') # Gmail rewriting port 25 to port 587
server.starttls() 							# Support SMPT AUTH
server.login(sender, password)
server.sendmail(msg['From'], msg['To'], msg.as_string())

#attachmement
# Name of document
current_date = datetime.now()
name = 'Rapport vervaldata - {}'.format(current_date.date())
msg.attach(MIMEText(message, 'plain'))
pdfname = 'C:/Users/bress/OneDrive/Bureaublad/project/github/AH3_project/Reports/{}.pdf'.format(name)
binary_pdf = open(pdfname, 'rb')

payload = MIMEBase('application', 'octa-stream', Name=pdfname)
payload.set_payload((binary_pdf).read())

encoders.encode_base64(payload)

# add header with pdf name
payload.add_header('Content-Decomposition', 'attachment', filename=pdfname)
msg.attach(payload)

#use gmail with port
session = smtplib.SMTP('smtp.gmail.com', 587)

#enable security
session.starttls()

#login with mail_id and password
session.login(sender, password)

text = msg.as_string()

#send_time = dt.datetime.now()
#time.sleep(send_time.timestamp() - time.time())

session.sendmail(sender, receiver, text)
session.quit()
server.quit()
print('Mail Sent')
