"""

Much of this code followed the tutorial found at https://realpython.com/python-send-email/#option-1-setting-up-a-gmail-account-for-development

"""

import smtplib, ssl, getpass
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

class Server:
    def __init__(self, sender_email, receiver_emails, subject, smtp_server="smtp.gmail.com"):
        self.sender_email = sender_email
        self.receiver_emails = receiver_emails
        self.subject = subject
        self.smtp_server = smtp_server
        self.password = getpass.getpass(prompt="Password: ")
        self.context = ssl.create_default_context()

    def format_message(self):
        pass

    def email(self):
        pass

class TlsConnectionEmail(Server):
    def __init__(self, sender_email, receiver_emails, subject, port=465):
        super().__init__(sender_email, receiver_emails, subject)
        self.port = port

    def format_message(self, receiver):
        """ Put your message here in both plain text and HTML """
        message = MIMEMultipart("alternative")
        message["Subject"] = self.subject
        message["From"] = self.sender_email
        message["To"] = receiver

        text = """\
        Hi,
        How are you?
        Real Python has many great tutorials:
        www.realpython.com"""

        html = """\
        <html>
        <body>
            <p>Hi,<br>
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
        for receiver in self.receiver_emails:
            with smtplib.SMTP_SSL(self.smtp_server, self.port, context=self.context) as server:
                server.login(self.sender_email, self.password)
                server.sendmail(self.sender_email, receiver, self.format_message(receiver).as_string())

if __name__=='__main__':
    sender = "x@gmail.com"
    receivers = ["test@gmail.com", "test2@gmail.com", "test3@gmail.com"]
    subject = "test email skeleton"

    tls = TlsConnectionEmail(sender, receivers, subject).email()