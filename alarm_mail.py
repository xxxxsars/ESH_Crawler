import datetime
import handler
import json
import os
import pandas as pd
import re
import sys
import time

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.select import Select

class EmptyDataFrameException(Exception): pass


LUNCH_DATE = "2021/08/24"

class Alarm_mail():
    def __init__(self,config_path):
        self.config_path = config_path
        self.config = handler.load_setting(config_path)
        self.crawler = self.config["Crawler"]

    def _overdue_mail(self, detail_link: str):

        recipient_map = self.config["Recipient"]
        recipient_list = [recipient_map[recipient] for recipient in recipient_map]

        copy_map = self.config["Copy"]
        copy_list = [copy_map[copy] for copy in copy_map]

        MAIL_SUBJECT = '環安申請單逾期警告'

        html = open("mail_templates.html", encoding='utf-8')
        html_content = html.read()

        soup = BeautifulSoup(html_content, 'html.parser')
        soup.find(id="detail-link", href=True)["href"] = detail_link

        mail_body = str(soup)

        handler.send_outlook_html_mail(recipients=recipient_list, subject=MAIL_SUBJECT, body=mail_body,
                                       send_or_display='SEND', copies=copy_list)


    def send_overdue_mail(self):

        history_dataframe = handler.read_history_esh(self.config_path)
        if history_dataframe.empty:
            raise EmptyDataFrameException(f'Please check your {self.crawler["xlsx_name"]} existed or had data.')

        month_ago_date = datetime.datetime.now() - datetime.timedelta(days=int(self.crawler["alert_days"]))
        alert_data = (history_dataframe[
            (history_dataframe["表單狀態"] != "已結案") & (pd.to_datetime(history_dataframe['發現日期']) < month_ago_date)])
        alert_mail_links = alert_data["詳細資訊"].values.tolist()
        print(alert_mail_links)
        if alert_mail_links:
            # TODO: modify to the for-loop mail
            self._send_mail(alert_mail_links[0])
        else:
            print("All esh form had been done.")

        # TODO: if saved excel file failed, it will save another file name with timestamp



if __name__ =="__main__":

    a = Alarm_mail("setting.ini")
    a.send_overdue_mail()