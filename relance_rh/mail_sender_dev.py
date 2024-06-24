import base64
import os
import smtplib
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from relance_rh.config.config_loader_dev import load_config


class MailSender:

	def __init__(self, use_mailhog=False):
		self.config = load_config()
		self.use_mailhog = use_mailhog
		self.sender_email = self.config['email']['sender_email']

	def encode_image_to_base64(self, image_path):
		try:
			with open(image_path, "rb") as image_file:
				encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
			return encoded_string
		except FileNotFoundError:
			print(f"Error: The file at {image_path} was not found.")
		except Exception as e:
			print(f"An error occurred while encoding the image: {e}")
		return None

	def send_email_after_3_moths(self, information):
		try:
			image_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "./asset/logo.png"))

			sender_email = self.sender_email
			smtp_server = self.config['email']['smtp_server']
			smtp_port = self.config['email']['smtp_port']
			login = self.config['email']['login']
			password = self.config['email']['password']

			receiver_email = information['email']
			first_name = information['first_name']
			msg = MIMEMultipart('related')
			msg['From'] = sender_email
			msg['To'] = receiver_email
			msg['Subject'] = " De belles opportunités !"
			# The email body
			# HTML body with styled content and a logo
			html_content = f"""\
	                    <html>
	                    <head></head>
	                    <body>
	                        <p style="font-family: Arial, sans-serif; font-size: 14px;">
	                            Bonjour {first_name},<br><br><br>

	                            <p style="color: #0D0D0D">Chargée des ressources humaines<p>
	                            <img src="cid:logo_akema" alt="logo_akema" style="width:110px; height:24px;">

	                        </p>
	                        </body>
	                        </html>
	                        """
			msg.attach(MIMEText(html_content, 'html'))

			# Attach image as related content
			image_base64 = self.encode_image_to_base64(image_path)
			if image_base64:
				img = MIMEImage(base64.b64decode(image_base64), name=os.path.basename(image_path))
				img.add_header('Content-ID', '<logo_akema>')
				img.add_header('Content-Disposition', 'inline', filename=os.path.basename(image_path))
				msg.attach(img)

			if self.use_mailhog:
				# Use smtplib for testing
				with smtplib.SMTP('localhost', 1025) as server:
					server.send_message(msg)
			else:
				# Send via SMTP server
				with smtplib.SMTP(smtp_server, smtp_port) as server:
					server.starttls()
					server.login(login, password)
					server.send_message(msg)

			return "Email OK"
		except Exception as e:
			return f"Email KO {e}"




