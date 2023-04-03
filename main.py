import os
from datetime import date, datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import pandas as pd
import logging
import re

EMAIL_ADDRESS = os.environ.get('EMAIL_ADDRESS')
EMAIL_PASSWORD = os.environ.get('EMAIL_PASSWORD')
SUCCESS_FILE_PATH = 'emails/email.success.txt'

logging.basicConfig(filename='demo.log', level=logging.DEBUG)

today = date.today().strftime("%m/%d/%Y")
success_count = 0
if os.path.isfile(SUCCESS_FILE_PATH):
    with open(SUCCESS_FILE_PATH, 'r') as f:
        last_success_date, last_success_count = f.read().split(',')
        if last_success_date == today:
            success_count = int(last_success_count)

email_queue = pd.read_excel('emails/email.queue.xlsx', header=0)
email_sent = pd.read_excel('emails/email.sent.xlsx', header=0)

emails_to_send = 65
emails_attempted = 0

logging.debug(f'------------------ BATCH RUNNING ON: {datetime.now().strftime("%m/%d/%Y %H:%M:%S")} ------------------')


def send_email_promo(sender, receiver, password):
    try:
        smtp = smtplib.SMTP('smtp.gmail.com', 587)
        smtp.starttls()

        smtp.login(sender, password)

        message = MIMEMultipart()
        message['Subject'] = 'FREE Day of Surveillance'
        message['From'] = sender
        message['To'] = receiver

        html = '''
            <html>
                <head>
                    <style>
                        .page {
                            background-color: #f2f2f2;
                        }
                        .container {
                            margin: auto;
                            max-width: 600px;
                            padding: 20px;
                        }
                        .unsubscribe {
                            text-align: center;
                            font-size: 14px;
                            color: #999999;
                            margin-top: 20px;
                        }
                    </style>
                </head>
                <body>
                    <div class="page">
                        <div class="container">
                            <img style="width: 100%; height:auto;" src="cid:advert_img">
                        </div>
                        <div class="unsubscribe">
                            To unsubscribe, reply "unsubscribe" to this email.
                        </div>
                    </div>
                </body>
            </html>
        '''

        html_body = MIMEText(html, 'html')
        message.attach(html_body)

        with open('images/OTI Promo.png', 'rb') as img:
            image = MIMEImage(img.read())
            image.add_header('Content-ID', '<advert_img>')
            message.attach(image)

        smtp.sendmail(sender, receiver, message.as_string())
        logging.debug(f'Email sent to: {receiver}')
        smtp.quit()
        return True

    except smtplib.SMTPException as e:
        logging.debug(f'an error occurred sending an email to: {receiver}')
        logging.debug(e)
        logging.debug(' ')
        try:
            smtp.quit()
            return False
        except:
            return False


if __name__ == '__main__':
    for i, email in email_queue.iterrows():
        is_sent = False

        if success_count >= 1850:
            logging.debug('Maximum emails sent for today')
            break

        if emails_attempted >= emails_to_send:
            break

        email = email.values[0].strip()

        if email_sent['Email'].str.lower().str.contains('^' + re.escape(email.lower()) + '$', regex=True).any():
            logging.debug(f'{email} is a duplicate')
            email_queue.drop(index=i, inplace=True)
            continue

        is_sent = send_email_promo(EMAIL_ADDRESS, email, EMAIL_PASSWORD)

        if is_sent:
            email_queue.drop(index=i, inplace=True)
            email_sent.loc[len(email_sent)] = [email, 'Yes', today, None]
            success_count += 1
            with open(SUCCESS_FILE_PATH, 'w') as f:
                f.write(f'{today},{success_count}')

        emails_attempted += 1

    email_queue.to_excel('emails/email.queue.xlsx', index=False)
    email_sent.to_excel('emails/email.sent.xlsx', index=False)

logging.debug('------------------ BATCH FINISHED ------------------')
