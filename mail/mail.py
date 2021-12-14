##########################################################################
# File:         mailer_invite.py                                         #
# Purpose:      Automatic send 專題說明會 invitation mails to professors   #
# Last changed: 2015/06/21                                               #
# Author:       Yi-Lin Juang (B02 學術長)                                 #
# Edited:       2021/07/01 Eleson Chuang (B08 Python大佬)                 #
#               2018/05/22 Joey Wang (B05 學術長)                         #
# Copyleft:     (ɔ)NTUEE                                                 #
##########################################################################
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
# from email.mime.image import MIMEImage
# import email.message

import sys
import time
from os import listdir
from os.path import join
import csv

from string import Template
import configparser as cp


def connectSMTP(user, pw):
    # Send the message via NTU SMTP server.
    # For students ID larger than 09
    s = smtplib.SMTP_SSL("smtps.ntu.edu.tw", 465)
    # For students ID smaller than 08 i.e. elders
    # s = smtplib.SMTP('mail.ntu.edu.tw', 587)
    s.set_debuglevel(False)
    # Uncomment this line to go through SMTP connection status.
    s.ehlo()
    if s.has_extn("STARTTLS"):
        s.starttls()
        s.ehlo()
    s.login(user, pw)
    print("SMTP Connected!")
    return s


def disconnect(server):
    server.quit()


def read_list(file_name):
    obj = list()
    with open(file_name, "r", encoding="utf-8") as f:
        for line in f:
            t = line.split()
            if t is not None:
                obj.append(t)
    return obj


def send_mail(msg, server):
    server.sendmail(msg["From"], msg["To"], msg.as_string())
    print("Sent message from {} to {}".format(msg["From"], msg["To"]))


def attach_files_method1(folder, msg):
    """
    This method will attach all the files in the ./attach folder.
    """
    dir_path = join(folder, "attach")
    files = listdir(dir_path)
    print(folder)
    for f in files:  # add files to the message
        file_path = join(dir_path, f)
        attachment = MIMEApplication(open(file_path, "rb").read())
        attachment.add_header('Content-Disposition', 'inline', filename=f)
        msg.attach(attachment)


def attach_files_method2(msg):
    """
    Reading attachment, put file_path in args
    """
    for argvs in sys.argv[1:]:
        attachment = MIMEApplication(open(str(argvs), "rb").read())
        attachment.add_header("Content-Disposition", "attachment", filename=str(argvs))
        msg.attach(attachment)


def send(folder):
    print(folder)
    # load email account info
    config = cp.ConfigParser()
    config.read(join(folder, "config.ini"), encoding="utf-8")  # reading sender account information

    rec_file = join(folder, "recipients.csv")

    account_info = config["ACCOUNT"]
    acc_user = account_info["user"]
    acc_pw = account_info["pw"]

    message_info = config["MESSAGE"]
    mes_from = message_info["from"]
    mes_subject = message_info["subject"]
    mes_content = join(folder, "letter.html")

    # recipient  = recipient's email address
    sender = "{}@ntu.edu.tw".format(acc_user)
    # 2. Sender email address (yours).
    # recipients = read_list(rec_file)
    # Uncomment this line to send to yourself. (for TESTING)
    server = connectSMTP(acc_user, acc_pw)
    count = 0

    # with open('password.csv', 'r', newline = '') as name:
    #     names = csv.reader(name)
    #     name_list = list()
    #     for n in names:
    #         name_list.append(n)

    with open(rec_file, 'r', newline='', encoding="utf-8") as csvfile:
        recipients = csv.reader(csvfile)
        column = {"name": -1, "id": -1}

        for recipient in recipients:
            if count % 10 == 0 and count > 0:
                print("{} mails sent, resting...".format(count))
                time.sleep(10)  # for mail server limitation
            if count % 130 == 0 and count > 0:
                print("{} mails sent, resting...".format(count))
                time.sleep(20)  # for mail server limitation
            if count % 260 == 0 and count > 0:
                print("{} mails sent, resting...".format(count))
                time.sleep(20)  # for mail server limitation

            if count == 0:
                column["name"] = recipient.index("name")
                column["id"] = recipient.index("id")
                count += 1
                continue

            msg = MIMEMultipart("alternative")

            # 讓寄件人顯示為本來信箱
            # msg["From"] = sender
            # 讓寄件人顯示改成學術部
            from_text = mes_from
            msg['From'] = formataddr((Header(from_text, 'utf-8').encode(), sender))

            # set subject
            msg["Subject"] = mes_subject

            '''
            # msg.preamble = "Multipart massage.\n"
            for n in name_list:
                if n[0] == recipient[0]:
                    name = n[1]
                    break
            '''

            # letter content
            user = recipient[column["name"]]
            content = mes_content
            template = Template(open(mes_content, "r", encoding="utf-8").read())
            # print(Path(content).read_text())
            body = template.substitute({"user": user})
            msg.attach(MIMEText(body, "html"))  # HTML郵件內容

            '''
            part = "<html>'{}同學您好：\n\n'.format(name)</html>"
            # part1 = MIMEText(text, "plain")
            part2 = MIMEText(html, "html")
            # msg.attach(part1)
            msg.attach(part2)
            '''

            # ./attach folder METHOD
            attach_files_method1(folder, msg)

            # sys.argv METHOD
            # attach_files_METHOD2(msg)

            identity = recipient[column["id"]]
            msg["To"] = (identity + "@ntu.edu.tw")

            send_mail(msg, server)
            count += 1

        disconnect(server)
        return "{} mails sent. Exiting...".format(count - 1)
