#!/usr/bin/env python
# -*- coding:utf-8 -*-
import datetime
from smtplib import SMTP_SSL
from email.mime.text import MIMEText
from email.header import Header
from apscheduler.schedulers.blocking import BlockingScheduler
from generate_html import MailTable


def send_mail(subject, html):
    mail_info = {
        "from": "taomercy@qq.com",
        "to": ["virdis.wu@autodesk.com"],
        "cc": [],
        "hostname": "smtp.qq.com",
        "username": "taomercy@qq.com",
        "password": "ibbwbcdeogiajjjg",
        "mail_subject": subject,
        "mail_text": html,
        "mail_encoding": "utf-8"
    }

    smtp = SMTP_SSL(mail_info["hostname"])
    # smtp.set_debuglevel(1)

    smtp.ehlo(mail_info["hostname"])
    smtp.login(mail_info["username"], mail_info["password"])
    msg = MIMEText(mail_info["mail_text"], _subtype="html", _charset=mail_info["mail_encoding"])
    msg["Subject"] = Header(mail_info["mail_subject"], mail_info["mail_encoding"])
    msg["From"] = mail_info["from"]
    msg["To"] = ",".join(mail_info["to"])
    msg["Cc"] = ",".join(mail_info["cc"])

    smtp.sendmail(mail_info["from"], mail_info["to"] + mail_info["cc"], msg.as_string())
    smtp.quit()
    print("Send success.")


def monitor_task():
    send_date = datetime.datetime.now()
    print("Send date:", send_date)
    mt = MailTable(send_date)
    subject = mt.get_subject()
    context = mt.get_html()
    send_mail(subject, context)


if __name__ == '__main__':
    scheduler = BlockingScheduler(timezone='Asia/Shanghai')
    scheduler.add_job(monitor_task, 'cron', day=1, hour=9, minute=30)
    # scheduler.add_job(monitor_task, 'interval', seconds=5)
    try:
        print("monitor start ...")
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        pass
