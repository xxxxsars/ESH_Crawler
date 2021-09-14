import configparser
import datetime
import os
import re
import requests
import pandas as pd
import subprocess
import win32com.client

from xlsx2csv import Xlsx2csv


def load_setting(config_path: str) -> configparser:
    config = configparser.ConfigParser()
    assert os.path.exists(config_path), f"Please verify the '{config_path}' existed. "
    config.read(config_path)
    return config


def date_condition(days_ago: int) -> dict[str, str]:
    assert isinstance(days_ago, int), "Your [days ago] values had some error."
    now = datetime.datetime.now()

    # now = datetime.datetime.strptime('2021/08/01', "%Y/%m/%d")
    ago = datetime.timedelta(days=int(days_ago))
    end_date = now.strftime("%Y/%m/%d")
    start_date = (now - ago).strftime("%Y/%m/%d")

    return {"start_date": start_date, "end_date": end_date}


def send_outlook_html_mail(recipients: list[str], subject: str = 'No Subject', body: str = 'Blank',
                           send_or_display: str = 'Display', copies: list[str] = None) -> None:
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


def read_history_esh(config_path: str):
    files = show_all_xlsx(config_path)
    if files:
        merge_df = pd.DataFrame()
        for file in show_all_xlsx('setting.ini'):
            merge_df = pd.concat([merge_df, faster_read_excel(file)])
        # remove duplicates data
        merge_df = merge_df.drop_duplicates(ignore_index=True, subset=['異常單系統編號'], keep='last')
        return merge_df
    else:
        return pd.DataFrame()


def upload_files(config_path: str) -> None:
    config = load_setting(config_path)
    url = config["Link"]["upload_url"]
    file_path = config["Crawler"]["xlsx_name"]

    if os.path.exists(file_path):
        files = {"files": open(file_path, 'rb')}
        data = {"site": "L3D", "path": "M2"}
        res = requests.post(url, files=files, data=data)
        assert res.status_code == 202, "Upload file had error."
    else:
        raise FileExistsError(f"Your {file_path} was not existed.")


def show_all_xlsx(config_path: str) -> list[str]:
    xlsx_files: list[str] = []

    config = load_setting(config_path)
    split_file_path = os.path.split(config["Crawler"]["xlsx_name"])

    root_path = split_file_path[0]
    file_prefix = os.path.splitext(split_file_path[1])[0]

    # enhance the performance of iterative files
    _, _, files = next(os.walk(root_path))

    for f in files:
        full_path = os.path.join(root_path, f)
        if os.path.isfile(full_path) and re.search(f'^{file_prefix}', f) and re.search("xlsx$", f):
            xlsx_files.append(full_path)

    assert len(xlsx_files) > 0, "Your root path didn't had any matched files."

    return xlsx_files


if __name__ == "__main__":
    print(load_setting("setting.ini"))