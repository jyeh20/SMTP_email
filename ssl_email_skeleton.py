"""

Much of this code followed the tutorial found at https://realpython.com/python-send-email/#option-1-setting-up-a-gmail-account-for-development

"""

import smtplib, ssl, getpass
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd

class TlsConnectionEmail():
    def __init__(self, sender_email, receivers, subject, port=465, smtp_server="smtp.gmail.com"):
        self.sender_email = sender_email
        self.receivers = receivers
        self.subject = subject
        self.smtp_server = smtp_server
        self.password = getpass.getpass(prompt="Password: ")
        self.context = ssl.create_default_context()
        self.port = port

    def format_message(self, receiver_name, receiver_email):
        """ Format your message here in both plain text and HTML"""
        message = MIMEMultipart("alternative")
        message["Subject"] = self.subject
        message["From"] = self.sender_email
        message["To"] = receiver_email

        text = f"""\
        Hi {receiver_name},
        How are you?
        Real Python has many great tutorials:
        www.realpython.com"""
        html = f"""\
        <html>
        <body>
            <p>Hi {receiver_name},<br>
            How are you?<br>
            <a href="http://www.realpython.com">Real Python</a>
            has many great tutorials.
            </p>
        </body>
        </html>
        """

        plain_text = MIMEText(text, "plain")
        html_text = MIMEText(html, "html")

        message.attach(plain_text)
        message.attach(html_text)

        return message

    def email(self):
        """ Send the email(s) using secure connection """
        with smtplib.SMTP_SSL(self.smtp_server, self.port, context=self.context) as server:
            server.login(self.sender_email, self.password)

            for _, receiver in self.receivers.iterrows():
                name, email = receiver['First Name'], receiver['Email Address']
                server.sendmail(self.sender_email, email, self.format_message(name, email).as_string())


if __name__=='__main__':
    sender = "x@gmail.com"
    subject = "Testing Jonathan's Email bot"

    df = pd.read_excel('./sample.xlsx')

    tls = TlsConnectionEmail(sender, df, subject).email()