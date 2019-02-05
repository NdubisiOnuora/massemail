import csv
import datetime
import json
from os.path import basename
import smtplib

from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

config = {}
with open('config.json') as f:
    config = json.load(f)

# sender == my email address
# recipient == receiver's email address
sender_email = config['sender_email']
sender_password = config['sender_password']
blind_copy_email = config['blind_copy_email']
email_template_text_file_name = config['email_template_text_file_name']
email_template_html_file_name = config['email_template_html_file_name']
email_subject = config['email_subject']
recipient_data_file_name = config['recipient_data_file_name']
report_file_name = config['report_file_name']
main_contact_name = config['main_contact_name']
attachment_file_names = config['attachment_file_names']
smtp_server_name = config['smtp_server']
smtp_server_port = config['smtp_port']


def create_message(sender, recipient_data, attachments, text_template, html_template):
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
    msg['Subject'] = email_subject
    msg['From'] = sender
    msg['To'] = recipient_email

    # Create the body of the message (a plain-text and an HTML version).
    text = text_template.format(name=recipient_first_name, company=recipient_company_name, main_contact_name=main_contact_name)
    html = html_template.format(name=recipient_first_name, company=recipient_company_name, main_contact_name=main_contact_name)

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


def setup_mail_server():
    # Send the message via local SMTP server.
    mail_server = smtplib.SMTP(smtp_server_name, smtp_server_port)
    mail_server.ehlo()
    mail_server.starttls()
    mail_server.login(sender_email, sender_password)

    return mail_server


if __name__ == '__main__':

    # Read the CSV file containing recipient information
    # Load the data into a list of dictionaries for each row
    # Create each message for each row
    # Send all messages at once (so as not to arouse suspicion)

    email_text_template = ''
    email_html_template = ''

    with open(email_template_text_file_name, 'r') as f:
        email_text_template = f.read()

    with open(email_template_html_file_name, 'r') as f:
        email_html_template = f.read()

    all_recipient_data = []

    # Read the CSV File
    print('Reading recipient file.')
    with open(recipient_data_file_name) as csvfile:
        reader = csv.DictReader(csvfile, fieldnames=('first_name', 'name', 'email', 'company'))
        all_recipient_data = [row for row in reader]

    # Create all messages
    msg_recip_tuples = []
    for recipient_data in all_recipient_data:
        message_recip_tuple = (recipient_data, create_message(sender_email, recipient_data, attachment_file_names, email_text_template, email_html_template))
        msg_recip_tuples.append(message_recip_tuple)

    print('Setting up mail server object.')
    mail_server = setup_mail_server()

    bcc_list = ','.join(blind_copy_email)
    print('Preparing to send messages to recipients')
    for mrt in msg_recip_tuples:
        recip_data = mrt[0]
        msg = mrt[1]
        all_recips = [recip_data['email']] + blind_copy_email
        mail_server.sendmail(sender_email, all_recips, msg.as_string())
        log_msg = 'Sending message to {recip_name}({recip_email}) for {company} | BCC: {bcc_list}'
        log_msg = log_msg.format(recip_name=recip_data['name'], recip_email=recip_data['email'],
                                 company=recip_data['company'], bcc_list=bcc_list)
        print(log_msg)

    print('Sent e-mail to {num} recipients'.format(num=len(msg_recip_tuples)))

    current_date = datetime.datetime.now()

    report_file_base_name, extension = report_file_name.split('.')
    report_file_name = '{base_file_name}_{year}{month}{day}T{hour}{minute}{second}.{ext}'.format(base_file_name=report_file_base_name,
                                                                                                 year=current_date.year,
                                                                                                 month=current_date.month,
                                                                                                 day=current_date.day,
                                                                                                 hour=current_date.hour,
                                                                                                 minute=current_date.minute,
                                                                                                 second=current_date.second,
                                                                                                 ext=extension)
    with open(report_file_name, 'wb') as csvfile:
        csv_writer = csv.writer(csvfile)
        for mrt in msg_recip_tuples:
            recip_data = mrt[0]
            csv_writer.writerow([recip_data['company'], main_contact_name, recip_data['name'], recip_data['email'], current_date])
    print('Created report and saved to {output}.'.format(output=report_file_name))
