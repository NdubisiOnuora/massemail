# massemail
Helps One to Send Mass E-Mails Especially for Sponsorship Requests

## What to Update in config.json
`sender_email`: your e-mail address

`sender_password`: your e-mail password

`email_subject`: the subject of your e-mail

`recipient_data_file_name`: the CSV file that the script will read from; it is in the following form
* Contact First Name, Contact Full Name, Contact E-Mail Address, Contact Company
* Example: Ndubisi,Ndubisi Onuora,ndubisionuora@gmail.com,Cox Automotive

`main_contact_name`: your first name

`attachment_file_names`: the names of attachments relative to the massemail directory

## Files to Update
* templates\email_template_html.html
* templates\email_template_text.txt


## Python Version
Python 2.7.14

## Packages
None; just standard library.

## How to Run
1) git clone https://github.com/NdubisiOnuora/massemail.git
2) Go to the directory "massemail"
3) python mailer.py (the application will take time to run)
4) Check the report_file_name_YYYYMMDDThhmmss (Example: massemail\output\sponsorship_submissionT20190204234310) to check a
report of successful submissions to recipients. 
