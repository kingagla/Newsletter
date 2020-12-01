import smtplib
import ssl
import mammoth
import os, sys
import pandas as pd
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication


def attach_picture(message, directory, filename):
    img = open(os.path.join(directory, filename), 'rb')
    image = MIMEImage(img.read(), filename.split('.')[1])
    img.close()
    image.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(image)
    return message


def attach_pdf(message, directory, filename):
    file = open(os.path.join(directory, filename), 'rb')
    pdf_attachment = MIMEApplication(file.read(), 'pdf')
    file.close()
    pdf_attachment.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(pdf_attachment)


def send_mail(from_, mail_password, directory, mails_names_file, text_mail_file, subject, *attachments):
    mails = pd.read_csv(os.path.join(directory, mails_names_file))
    for mail, name in mails.values:
        try:
            # convert message to html
            f = open(os.path.join(directory, text_mail_file), 'rb')
            document = mammoth.convert_to_html(f)
            msg = (document.value.encode('utf8'))
            f.close()

            # personalize message
            msg = msg.decode('utf8').format(name)

            # create multi-part message and define each element
            message = MIMEMultipart()
            message['Subject'] = subject
            message['From'] = f'Kinga Glabinska'
            message['To'] = mail

            # attach text of message
            message.attach(MIMEText(msg, 'html', _charset='utf-8'))

            # attach files if any
            for attachment in attachments:
                if attachment.endswith('.pdf'):
                    attach_pdf(message, directory, attachment)
                else:
                    attach_picture(message, directory, attachment)

            # send email
            smtpserver = smtplib.SMTP("smtp.gmail.com", 587)
            smtpserver.ehlo()
            smtpserver.starttls()
            smtpserver.ehlo()
            smtpserver.login(from_, mail_password)
            smtpserver.sendmail(from_, mail, message.as_string())
            smtpserver.close()

            print('Mail sent to', mail)
        except Exception as e:
            print('Mail did not send to', mail, 'due to', e)


if __name__ == '__main__':
    from_ = 'my_mail@gmail.com'
    mail_password = input('Please type your password: ')
    directory = './files'
    mails_names_file = 'mails.csv'
    text_mail_file = 'urodziny.docx'
    subject = '30 urodziny restauracji'
    send_mail(from_, mail_password, directory, mails_names_file, text_mail_file, subject, 'logo.jpg', 'menu.pdf')
