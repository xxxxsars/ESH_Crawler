from bs4 import BeautifulSoup
import configparser
import datetime
import os
import pandas as pd
import subprocess
import win32com.client


CONFIG_PATH = "setting.ini"


def load_setting() -> configparser:
    config = configparser.ConfigParser()
    config.read(CONFIG_PATH)
    return config


def date_condition() -> dict[str, str]:
    days_ago = int(load_setting()["Crawler"]["days_ago"])

    assert isinstance(days_ago, int), "Your [days ago] values had some error."
    now = datetime.datetime.now()

    # now = datetime.datetime.strptime('2021/08/01', "%Y/%m/%d")
    ago = datetime.timedelta(days=int(days_ago))

    end_date = now.strftime("%Y/%m/%d")

    start_date = (now - ago).strftime("%Y/%m/%d")

    return {"start_date": start_date, "end_date": end_date}


def send_outlook_html_mail(recipients: list[str], subject: str = 'No Subject', body: str = 'Blank',
                           send_or_display: str = 'Display', copies: list[str] = None):
    if len(recipients) > 0 and isinstance(recipients, list):
        outlook = win32com.client.Dispatch("Outlook.Application")
        ol_msg = outlook.CreateItem(0)
        str_to = ""

        for recipient in recipients:
            str_to += recipient + ";"
        ol_msg.To = str_to

        if copies is not None and len(copies) > 0:
            str_cc = ""
            for cc in copies:
                str_cc += cc + ";"

            ol_msg.CC = str_cc

        ol_msg.Subject = subject
        ol_msg.HTMLBody = body

        if send_or_display.upper() == 'SEND':
            ol_msg.Send()
        else:
            ol_msg.Display()
    else:
        print('Recipient email address - NOT FOUND')


def faster_read_excel(xlsx_path: str, usecols: list[str] = None) -> pd:
    tmp_csv = "tmp.csv"

    call = ["python", "./xlsx2csv.py", xlsx_path, tmp_csv]
    try:
        subprocess.call(call)  # On Windows use shell=True
    except:
        print('Failed with {}'.format(xlsx_path))

    dataframe = pd.read_csv(tmp_csv, usecols=usecols)

    # if removed tmp file had error will ignore it.
    try:
        os.remove(tmp_csv)
    except:
        pass

    return dataframe

def send_mail(detail_link:str):
    setting = load_setting()
    recipient_map = setting["Recipient"]
    recipient_list = [recipient_map[recipient] for recipient in recipient_map]

    copy_map = setting["Copy"]
    copy_list = [copy_map[copy] for copy in copy_map]

    MAIL_SUBJECT = '環安申請單逾期警告'

    html = open("mail_templates.html", encoding='utf-8')
    html_content = html.read()

    soup = BeautifulSoup(html_content, 'html.parser')
    soup.find(id="detail-link", href=True)["href"] = detail_link

    mail_body = str(soup)

    send_outlook_html_mail(recipients=recipient_list, subject=MAIL_SUBJECT, body=mail_body,
                                       send_or_display='SEND', copies=copy_list)