# -*- coding: utf-8 -*-

import datetime
import smtplib
from threading import Thread

from mailer import Mailer


class ThreadedMailer(Mailer):
    def __init__(self, mail_config, num_threads):
        super(ThreadedMailer, self).__init__(mail_config)
        self.num_threads = num_threads

    def _setup_mail_server(self):
        # Send the message via local SMTP server.
        mail_server = smtplib.SMTP(self.config.smtp_server_name, self.config.smtp_server_port)
        mail_server.ehlo()
        mail_server.starttls()
        mail_server.login(self.config.sender_email, self.config.sender_password)

        return mail_server

    def _send_message(self, all_recipients, message):
        mail_server = self._setup_mail_server()
        try:
            mail_server.sendmail(self.config.sender_email, all_recipients, message.as_string())
        except Exception as e:
            print('Problems sending messages to {}'.format(all_recipients))
            print(e)

    def _run(self, message_recipient_tuples, bcc_list, current_date=None):
        self.all_threads = []
        start = 0
        increment = len(message_recipient_tuples) // self.num_threads
        current_date = datetime.datetime.now()
        for i in range(self.num_threads):
            end = len(message_recipient_tuples) if start + increment >= len(message_recipient_tuples) else start + increment
            if end - start == 0 or end > len(message_recipient_tuples):
                break
            print('Start={} End={}'.format(start, end))
            thread = Thread(target=self.send_message, args=(message_recipient_tuples[start:end], bcc_list, current_date))
            thread.start()
            self.all_threads.append(thread)

            start += increment

        # The thread will complete the rest of the work
        print('Start={} End={}'.format(start, len(message_recipient_tuples)))
        self.send_message(message_recipient_tuples[start:len(message_recipient_tuples)], bcc_list, current_date)

        for t in self.all_threads:
            t.join()
