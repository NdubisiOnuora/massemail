# -*- coding: utf-8 -*-

import csv
import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from os.path import basename
import smtplib

from config import Config


class Mailer(object):
    def __init__(self, mail_config):
        self.config = mail_config
        print('Setting up mail server object.')
        self._setup_mail_server()

    def _setup_mail_server(self):
        # Send the message via local SMTP server.
        self.mail_server = smtplib.SMTP(self.config.smtp_server_name, self.config.smtp_server_port)
        self.mail_server.ehlo()
        self.mail_server.starttls()
        self.mail_server.login(self.config.sender_email, self.config.sender_password)

    def create_message(self, sender, recipient_data, attachments, text_template, html_template):
        """
        :param sender:
        :param recipient_data:
        :param attachments:
        :param text_template:
        :param html_template:
        :return:
        """
        recipient_email = recipient_data['email'].strip()
        recipient_first_name = recipient_data['first_name'].strip()
        recipient_company_name = recipient_data['company'].strip()

        # Create message container - the correct MIME type is multipart/alternative.
        msg = MIMEMultipart('alternative')
        msg['Subject'] = self.config.email_subject
        msg['From'] = sender
        msg['To'] = recipient_email

        # Create the body of the message (a plain-text and an HTML version).
        text = text_template.format(name=recipient_first_name, company=recipient_company_name,
                                    main_contact_name=self.config.main_contact_name)
        html = html_template.format(name=recipient_first_name, company=recipient_company_name,
                                    main_contact_name=self.config.main_contact_name)

        # Record the MIME types of both parts - text/plain and text/html.
        part1 = MIMEText(text, 'plain')
        part2 = MIMEText(html, 'html')

        # Attach parts into message container.
        # According to RFC 2046, the last part of a multipart message, in this case
        # the HTML message, is best and preferred.
        msg.attach(part1)
        msg.attach(part2)

        # Attach the attachments
        for a in attachments:
            with open(a, 'rb') as f:
                part = MIMEApplication(f.read(), Name=basename(a))
            part['Content-Disposition'] = 'attachment; filename="{basename}"'.format(basename=basename(a))
            msg.attach(part)

        return msg

    def send_message(self, message_recipient_tuples, bcc_list, current_date=None):
        for mrt in message_recipient_tuples:
            recip_data = mrt[0]
            msg = mrt[1]
            all_recips = [recip_data['email']] + self.config.blind_copy_email

            self._send_message(all_recips, msg)
            log_msg = 'Sending message to {recip_name}({recip_email}) for {company} | BCC: {bcc_list}'
            log_msg = log_msg.format(recip_name=recip_data['name'], recip_email=recip_data['email'],
                                     company=recip_data['company'], bcc_list=bcc_list)
            print(log_msg)

        print('Sent e-mail to {num} recipients'.format(num=len(message_recipient_tuples)))

        if not current_date:
            current_date = datetime.datetime.now()
        self.create_report(message_recipient_tuples, current_date)

    def _send_message(self, all_recipients, message):
        self.mail_server.sendmail(self.config.sender_email, all_recipients, message.as_string())

    def create_report(self, message_recipient_tuples, current_date):
        report_file_base_name, extension = self.config.report_file_name.split('.')
        report_file_name = '{base_file_name}_{year}{month}{day}T{hour}{minute}{second}.{ext}'.format(
            base_file_name=report_file_base_name,
            year=current_date.year,
            month=current_date.month if current_date.month > 9 else '0{}'.format(str(current_date.month)),
            day=current_date.day if current_date.day > 9 else '0{}'.format(str(current_date.day)),
            hour=current_date.hour if current_date.hour > 9 else '0{}'.format(str(current_date.hour)),
            minute=current_date.minute if current_date.minute > 9 else '0{}'.format(str(current_date.minute)),
            second=current_date.second if current_date.second > 9 else '0{}'.format(str(current_date.second)),
            ext=extension)
        with open(report_file_name, 'a') as csvfile:
            csv_writer = csv.writer(csvfile)
            for mrt in message_recipient_tuples:
                recip_data = mrt[0]
                csv_writer.writerow(
                    [recip_data['company'], self.config.main_contact_name, recip_data['name'], recip_data['email'],
                     current_date])
        print('Created report and saved to {output}.'.format(output=report_file_name))

    def run(self):
        # Read the CSV file containing recipient information
        # Load the data into a list of dictionaries for each row
        # Create each message for each row
        # Send all messages at once (so as not to arouse suspicion)

        with open(self.config.email_template_text_file_name, 'r') as f:
            email_text_template = f.read()

        with open(self.config.email_template_html_file_name, 'r') as f:
            email_html_template = f.read()

        # Read the CSV File
        print('Reading recipient file.')
        with open(self.config.recipient_data_file_name) as csvfile:
            reader = csv.DictReader(csvfile, fieldnames=('first_name', 'name', 'email', 'company'))
            all_recipient_data = [row for row in reader]

        # Create all messages
        msg_recip_tuples = []
        for recipient_data in all_recipient_data:
            created_message = self.create_message(self.config.sender_email, recipient_data,
                                                  self.config.attachment_file_names, email_text_template, email_html_template)
            message_recip_tuple = (recipient_data, created_message)
            msg_recip_tuples.append(message_recip_tuple)

        bcc_list = ','.join(self.config.blind_copy_email)
        print('Preparing to send messages to recipients')
        self._run(msg_recip_tuples, bcc_list)

    def _run(self, message_recipient_tuples, bcc_list, current_date=None):
        self.send_message(message_recipient_tuples, bcc_list, current_date)


if __name__ == '__main__':
    config = Config('config.json')
    mailer = Mailer(config)

    start_time = datetime.datetime.now()
    mailer.run()
    end_time = datetime.datetime.now()
    diff = end_time - start_time
    print('Start Time={}'.format(start_time))
    print('End Time={}'.format(end_time))
    print('Difference=({seconds} seconds)'.format(seconds=diff.total_seconds()))
