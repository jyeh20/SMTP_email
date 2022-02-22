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
        """ Put your message here in both plain text and HTML"""
        message = MIMEMultipart("alternative")
        message["Subject"] = self.subject
        message["From"] = self.sender_email
        message["To"] = receiver_email

        text = f"""\
        Hello {receiver_name},
            My name is Jonathan and I am currently the Co-President for LMU’s chapter of ACM (Association of Computing Machinery). I am reaching out because we are hosting an upcoming alumni night (nicknamed ‘Alumnight’)! The event will be a casual networking event where alumni can talk to current students about their post-LMUCS experiences. No preparation or presentations are needed. This event will be on Friday, March 11th, 6-8pm on campus. Food will be provided.
            Please let me know if you have any questions! If you are interested in attending, please fill out this form: https://forms.gle/cBuFNJgi2DSHAggE9 by March 4th.

        Best,
        Jonathan Yeh
        LMU ACM Co-President || DRN Resident Advisor
        Computer Science Major
        Class of 2022
        """

        html = f"""\
        <html>
        <head>
        <style>
            .name {{
                color: #51a7f9;
                font-weight: bold;
                font-family: Arial;
            }}
        </style>
        </head>
        <body>
            <p>Hello {receiver_name},<br>
            <p>My name is Jonathan and I am currently the Co-President for LMU’s chapter of ACM (Association of Computing Machinery). I am reaching out because we are hosting an upcoming alumni night (nicknamed ‘Alumnight’)! The event will be a casual networking event where alumni can talk to current students about their post-LMUCS experiences. No preparation or presentations are needed. This event will be on <b>Friday, March 11th, 6-8pm on campus.</b> Food will be provided.</p>
            <p>Please let me know if you have any questions by replying to this email! If you are interested in attending, please fill out <a href="https://forms.gle/cBuFNJgi2DSHAggE9">this form</a> by <b>March 4th.</b></p>
            </p><br>
            Best,<br>
            <div class="name">Jonathan Yeh</div>
            LMU ACM Co-President || DRN Resident Advisor<br>
            Computer Science Major<br>
            Class of 2022</p><br>
            <img src="https://s3.amazonaws.com/lmuemailsignature/email-sig-logo-update.png" alt="LMU Logo">
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

                print(f"sending to {receiver['First Name']}, {receiver['Email Address']}")
                server.sendmail(self.sender_email, "jmyeh51@gmail.com", self.format_message(name, email).as_string())
                print(f"sent to {receiver['First Name']}, {receiver['Email Address']}")


if __name__=='__main__':
    sender = "lmuacm0@gmail.com"
    subject = "Testing lmuacm gmail"

    df = pd.read_excel('./Senior_Project_Invitees.xlsx')
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

    rows_to_drop = []
    for index, row in df.iterrows():
        try:
            if df["UG"][index] >= 2022 or not df["Email Address"][index]:
                rows_to_drop.append(index)
        except Exception as e:
            rows_to_drop.append(index)
            continue

    df.drop(labels=rows_to_drop, axis=0, inplace=True)
    df.drop(columns=["UG", "Last Name"], inplace=True)

    tls = TlsConnectionEmail(sender, df, subject).email()