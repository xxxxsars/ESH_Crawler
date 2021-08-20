import argparse
import datetime
import sys

import handler
import os
import pandas as pd
import re
import time

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.select import Select

# handle the characters that fail to convert xlsx file to csv .
ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]|\n|\xa0')


class EmptyDataFrameException(Exception): pass


class Crawler:
    def __init__(self, config_path):
        self.config = handler.load_setting(config_path)

        # initial selenium driver
        self.driver = None

    def __enter__(self):
        # read the setting.ini data
        self.link = self.config["Link"]
        self.user = self.config["User"]
        self.crawler = self.config["Crawler"]

        # read the history esh xlsx file data
        self.raw_dataframe = self._read_history_esh()

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.driver:
            self.driver.close()
            self.driver.quit()
        if exc_type == None and exc_val == None and exc_tb == None:
            sys.exit(0)

    def _read_history_esh(self):
        history_xlsx_path = self.crawler["xlsx_name"]
        if os.path.exists(history_xlsx_path):
            raw_dataframe = handler.faster_read_excel(self.crawler["xlsx_name"])
            return raw_dataframe
        else:
            return pd.DataFrame()

    def _clean_unless_data(self):

        start_date = datetime.datetime.strptime(handler.date_condition(int(self.crawler["days_ago"]))["start_date"],
                                                "%Y/%m/%d")
        over_years_date = datetime.datetime.now() - datetime.timedelta(days=int(self.crawler["keep_years"]) * 365)

        remove_index = self.raw_dataframe[(pd.to_datetime(self.raw_dataframe['發現日期']) > start_date) | (
                pd.to_datetime(self.raw_dataframe['發現日期']) < over_years_date)].index
        self.raw_dataframe.drop(remove_index, inplace=True)

    def _login(self):
        op = webdriver.ChromeOptions()
        # op.add_argument('headless')
        self.driver = webdriver.Chrome(options=op)
        self.driver.set_page_load_timeout(self.crawler["time_out_seconds"])

        self.driver.get(self.link["esh_url"])

        input_account = self.driver.find_element_by_id("txtInputId")
        input_account.send_keys(self.user["account"])

        input_password = self.driver.find_element_by_id("txtPassword")
        input_password.send_keys(self.user["password"])

        btn_login = self.driver.find_element_by_id("ext-gen22")
        btn_login.click()

        time.sleep(2)

    def _esh_conditions(self):
        input_fab = Select(self.driver.find_element_by_id("ddlFab"))
        input_fab.select_by_value("118")

        date_info = handler.date_condition(int(self.crawler["days_ago"]))

        input_start_date = self.driver.find_element_by_id("str_tm1")
        input_start_date.send_keys(date_info["start_date"])

        input_start_date = self.driver.find_element_by_id("str_tm2")
        input_start_date.send_keys(date_info["end_date"])

        soup = BeautifulSoup(self.driver.page_source, 'html.parser')
        for option in soup.find(id="ddlStatus").findAll("option"):
            option_value = (option["value"])
            input_status = Select(self.driver.find_element_by_id("ddlStatus"))
            input_status.select_by_value(option_value)

        self.driver.find_element_by_id("btnSearch").click()

    def _esh_crawler(self):
        soup = BeautifulSoup(self.driver.page_source, 'html.parser')

        # Waiting for web page search to complete.
        while True:
            table_element = soup.find(id="QualityAbnorGrid")
            if table_element:
                break

        detail_url_prefix = "http://augic8/ESH/ESH/"
        column: list[str] = []
        rows: list[str] = []

        for index, tr in enumerate(table_element.find("tbody").find_all("tr")):
            row: list[str] = []

            for td in tr.find_all('td'):
                value = ILLEGAL_CHARACTERS_RE.sub(r'', td.text)
                if value == "View":
                    detail_url = detail_url_prefix + (td.find("a", href=True)['href'])
                    row.append(detail_url)
                else:
                    row.append(value)

            # The first line was the column name
            if index == 0:
                column = row
                column[0] = "詳細資訊"
            else:
                rows.append(row)

        df = pd.DataFrame(rows, columns=column)

        if self.raw_dataframe.empty == False:
            self._clean_unless_data()

        # appended esh df to history dataframe
        new_dataframe = self.raw_dataframe.append(df, ignore_index=True)
        new_dataframe.to_excel(self.crawler["xlsx_name"], index=False)

    def _send_mail(self, detail_link: str):

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

    def run_crawler(self):

        self._login()
        self._esh_conditions()
        self._esh_crawler()

    def send_alert_mail(self):

        history_dataframe = self._read_history_esh()
        if history_dataframe.empty:
            raise EmptyDataFrameException(f'Please check your {self.crawler["xlsx_name"]} existed or had data.')

        month_ago_date = datetime.datetime.now() - datetime.timedelta(days=int(self.crawler["alert_days"]))
        alert_data = (history_dataframe[
            (history_dataframe["表單狀態"] != "已結案") & (pd.to_datetime(history_dataframe['發現日期']) < month_ago_date)])
        alert_mail_links = alert_data["詳細資訊"].values.tolist()
        if alert_mail_links:
            # TODO: modify to the for-loop mail
            self._send_mail(alert_mail_links[0])
        else:
            print("All esh form had been done.")



