""" 
    Envia los emails para el pago de promociones a los supermercadistas a partir de un excel

    Copyright (C) 2021  Mario Mori

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""

import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

class SendEmail:
    
    def __init__(self, username=None, sender=None, password=None, mail_server=None, mail_use_ssl=True, mail_use_tls=False, mail_port=None):
        self.username = username
        self.sender = sender
        self.password = password
        self.mail_server = mail_server
        self.mail_use_ssl = mail_use_ssl
        self.mail_use_tls = mail_use_tls
        self.mail_port = mail_port

    def send_email(self, body="", html="", recipent="", subject="", attach=None):

        if recipent != "":
            message = MIMEMultipart("alternative")
            message["Subject"] = subject
            message["From"] = self.sender
            message["To"] = recipent
            if attach != None:
                with open(attach, 'rb') as attachment:
                    part = MIMEBase('application','octet-stream')
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename={attach}",
                )
                message.attach(part)
        else:
            raise "There is no recipent"

        # Create the plain text and HTML version of your message
        # Turn these into plain/html MIMEText objects
        part1 = MIMEText(body, "plain")
        part2 = MIMEText(html, "html")

        # Add HTML/text-plain parts to MIMEMultipart objects
        # The email client will try to render the last part first
        if body != "":
            message.attach(part1)
        if html != "":
            message.attach(part2)

        # Create a secure SSL context
        if self.mail_use_ssl:
            context = ssl.create_default_context()
        with smtplib.SMTP_SSL(self.mail_server, self.mail_port, context=context) as server:
            server.login(self.username, self.password)
            server.sendmail(self.sender, recipent, message.as_string())
