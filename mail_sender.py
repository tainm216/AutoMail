# mail_sender.py

import yagmail
import time


def send_mail(sender_email, app_password, receiver_email, subject, body,
              pdf_file):

    yag = yagmail.SMTP(user=sender_email, password=app_password)

    yag.send(to=receiver_email,
             subject=subject,
             contents=[body],
             attachments=pdf_file)

    print(f"Đã gửi tới {receiver_email}")

    time.sleep(2)
