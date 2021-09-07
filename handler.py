import configparser
import datetime
import os
import pandas as pd
import subprocess
import win32com.client

from xlsx2csv import Xlsx2csv

def load_setting(config_path:str) -> configparser:
    config = configparser.ConfigParser()
    #assert os.path.exists(config_path), f"Please verify the '{config_path}' existed. "
    config.read(config_path)
    return config


def date_condition(days_ago:int) -> dict[str, str]:
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

    try:
        Xlsx2csv(xlsx_path).convert(tmp_csv, sheetid=1)
    except:
        print('Failed with {}'.format(xlsx_path))

    try:
        dataframe = pd.read_csv(tmp_csv, usecols=usecols)
    except:
        dataframe = pd.DataFrame()


    # if removed tmp file had error will ignore it.
    try:
        os.remove(tmp_csv)
    except:
        pass

    return dataframe


def read_history_esh(history_xlsx_path:str):

    if os.path.exists(history_xlsx_path):
        raw_dataframe = faster_read_excel(history_xlsx_path)
        return raw_dataframe
    else:
        return pd.DataFrame()


