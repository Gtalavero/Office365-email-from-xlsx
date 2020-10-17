# -*- coding: utf-8 -*-
import openpyxl
import smtplib
import sys, os
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

xlsx_file = Path(os.environ["FILE_PATH"], os.environ["FILE_NAME"])
wb_obj = openpyxl.load_workbook(xlsx_file)

# Read the active sheet:
sheet = wb_obj.active

# Email config
email_from = ''
password = ''
email_to = ''

def send_email():
    msg = MIMEText(message)
    msg['Subject'] = subject
    msg['From'] = email_from
    msg['To'] = email_to

    s = smtplib.SMTP('smtp.office365.com', 25, timeout=20)
    # Other port if doesn't work
    #s = smtplib.SMTP("smtp.outlook.office365.com", 587, timeout=20)
    s.starttls()
    s.login(email_from, password)
    s.sendmail(email_from , email_to, msg.as_string())
    s.quit()

#Row to begin to read
row = 3
col1 = "D" + str(row)
col2 = "B" + str(row)
col3 = "A" + str(row)
subject = f"Metting {sheet[col1].value} 22/10"
message = f"""Dear {sheet[col2].value},
You must give me {sheet[col3].value} dollars."""
send_email()

#https://www.marsja.se/your-guide-to-reading-excel-xlsx-files-in-python/