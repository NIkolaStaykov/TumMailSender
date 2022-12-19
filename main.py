import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import json
import shutil
from pathlib import Path
from typing import BinaryIO

import pandas as pd
import re
from tqdm import tqdm


class DataManager:
    def __init__(self, student_records_path: Path, submissions_path: Path):
        self.records = pd.read_csv(student_records_path)
        self.name_pattern = re.compile(r"([\w-]+[ ]?){2,7}(?=_\d)")
        self.submissions_path = submissions_path
        self.submissions_info = {}
        self.extract_submissions_info()

        self.sent_submissions_folder = submissions_path / "../Sent"
        self.sent_submissions_folder.resolve()
        if not self.sent_submissions_folder.exists():
            self.sent_submissions_folder.mkdir()

    def names_from_submission_folder(self, submission_folder: Path) -> (str, str):
        names = self.name_pattern.search(submission_folder.as_posix()).group(0).split()
        num_names = len(names)
        for i in range(1, num_names):
            vorname = " ".join(names[:i])
            nachname = " ".join(names[i:])
            if self.query_records(vorname, nachname):
                return vorname, nachname
        else:
            raise Exception(f"Submission folder {submission_folder} is named dumb")

    def query_records(self, vorname: str, nachname: str) -> bool:
        if ((self.records["Vorname"] == vorname) & (self.records["Nachname"] == nachname)).any():
            return True
        return False

    def mail_address_from_database(self, vorname: str, nachname: str) -> str:
        condition_mask = (self.records["Vorname"] == vorname) & (self.records["Nachname"] == nachname)
        receiver_mail = self.records[condition_mask]["E-Mail-Adresse"]
        return receiver_mail

    def extract_submissions_info(self):
        for submission_folder in self.submissions_path.iterdir():
            vorname, nachname = self.names_from_submission_folder(submission_folder)
            mail_address = self.mail_address_from_database(vorname, nachname)
            self.submissions_info[submission_folder] = {"vorname": vorname,
                                                        "nachname": nachname,
                                                        "mail": mail_address}

    def move_to_sent(self, submission_folder: Path):
        new_path = self.sent_submissions_folder / submission_folder.name
        self.submissions_info.pop(submission_folder)
        submission_folder.rename(new_path)

class Mail:
    def __init__(self, sender: str, receiver: str, body: str, subject: str):
        self.message = MIMEMultipart()
        self.sender = sender
        self.receiver = receiver
        self.create_message(body, subject)

    def create_message(self, body: str, subject: str):
        self.message['From'] = self.sender
        self.message['To'] = self.receiver
        self.message['Subject'] = subject
        self.message.attach(MIMEText(body, 'plain'))

    def attach_pdf(self, pdf_path: Path):
        # open the file in bynary
        binary_pdf: BinaryIO = open(pdf_path, 'rb')

        payload = MIMEBase('application', 'octate-stream', Name=pdf_path.as_posix())
        payload.set_payload(binary_pdf.read())

        # enconding the binary into base64
        encoders.encode_base64(payload)

        # add header with pdf name
        payload.add_header('Content-Decomposition', 'attachment', filename=pdf_path.as_posix())
        self.message.attach(payload)


class MailSendingService:
    def __init__(self, sender, sender_password, server='postout.lrz.de', port=587):
        self.sender = sender
        self.password = sender_password
        self.server = server
        self.port = port
        self.session = self.start_session()

    def start_session(self):
        # connect to mail server
        # https://doku.lrz.de/pages/viewpage.action?pageId=19103885
        session = smtplib.SMTP(self.server, self.port)

        # enable security
        session.starttls()

        # login with mail_id and password
        session.login(self.sender, self.password)
        return session

    def send_mail(self, mail: Mail):
        text = mail.message.as_string()
        self.session.sendmail(mail.sender, mail.receiver, text)

    def close_connection(self):
        self.session.quit()


def main():
    config_path = "config.json"
    config = json.load(open(config_path))

    sender, password = config["sender_data"].values()
    student_records_path = Path(config["records_path"])
    submissions_path = Path(config["submissions_path"])

    mail_service = MailSendingService(sender, password)
    data_manager = DataManager(student_records_path, submissions_path)

    subject = "Individual submission correction"
    body = open("mail_body.txt").read()

    for (submission_folder, receiver_info) in tqdm(data_manager.submissions_info.copy().items()):
        mail = Mail(sender, receiver_info["mail"], body, subject)
        pdf_path = list(submission_folder.glob("./*.pdf"))[0]
        mail.attach_pdf(pdf_path)
        mail_service.send_mail(mail)
        data_manager.move_to_sent(submission_folder)

    mail_service.close_connection()
    print("Done")


if __name__ == "__main__":
    main()
