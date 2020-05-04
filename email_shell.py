import smtplib
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from string import Template


def sendEmailWithAttachment(my_email, my_password, cus_name, cus_email, filedir, filename):

	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.starttls() # securing the server
	server.login(my_email, my_password)

	msg = MIMEMultipart()
	msg['From']		= my_email
	msg['To']		= cus_email
	msg['Subject']	= 'Your purchase receipt'

	with open('email_body.txt', 'r', encoding='utf-8') as email_body:
		email_message = email_body.read()
		email_message = Template(email_message)

	message = email_message.substitute(PERSON_NAME=cus_name)

	msg.attach(MIMEText(message, 'plain'))

	fullfilepath = filedir+filename

	# Open PDF file in binary mode
	with open(fullfilepath, "rb") as attachment:
	    # Add file as application/octet-stream
	    # Email client can usually download this automatically as attachment
	    part = MIMEBase("application", "octet-stream")
	    part.set_payload(attachment.read())

	# Encode file in ASCII characters to send by email    
	encoders.encode_base64(part)

	# Add header as key/value pair to attachment part
	part.add_header(
	    "Content-Disposition",
	    f"attachment; filename= {filename}",
	)

	# Add attachment to message and convert message to string
	msg.attach(part)

	server.send_message(msg)
	print('Email sent to: %s' % cus_name)

	del msg

	server.quit()