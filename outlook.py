import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os.path

class Outlook:
	def __init__(self, sender, username, password):
		'''
		Instantiate the class to connect to outlook
		'''
		self.sender = sender
		self.server = smtplib.SMTP('smtp.office365.com', 587)
		
		try:
			self.server.ehlo()
			self.server.starttls()
			self.server.login(username, password)
		except:
			print('SMPT server connection error')
		

	def send_email(self, recipient, subject, message, attachment_locations = []):
		msg = MIMEMultipart()
		msg['From'] = self.sender
		msg['To'] =  recipient
		msg['Subject'] = subject
		
		msg.attach(MIMEText(message, 'plain'))
		
		if len(attachment_locations) > 0:
			for attachment_location in attachment_locations:
				filename = os.path.basename(attachment_location)
				attachment = open(attachment_location, "rb")
				part = MIMEBase('application', 'octet-stream')
				part.set_payload(attachment.read())
				encoders.encode_base64(part)
				part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
				msg.attach(part)
		
		text = msg.as_string()
		
		try:
			server = self.server
			server.sendmail(self.sender, recipient, text)
		except:
			print("server connection error")
		return True

	def quit_connection(self):
		self.server.quit()

	

