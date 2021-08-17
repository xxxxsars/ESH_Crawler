import datetime
import handler
import os
import pandas as pd
import re
import time

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.select import Select

ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]|\n|\xa0')

class Crawler:

    def __init__(self):
        config = handler.load_setting()
        # read the setting.ini data
        self.link = config["Link"]
        self.user = config["User"]
        self.crawler = config["Crawler"]

        # read the history esh xlsx file data
        self.raw_dataframe = self._read_history_esh()

        # initial selenium driver
        self.driver = None

    def _read_history_esh(self):
        history_xlsx_path = self.crawler["xlsx_name"]
        if os.path.exists(history_xlsx_path):
            raw_dataframe = handler.faster_read_excel(self.crawler["xlsx_name"])
            return raw_dataframe
        else:
            return pd.DataFrame()

    def _clean_unless_data(self):

        start_date = datetime.datetime.strptime(handler.date_condition()["start_date"], "%Y/%m/%d")
        over_years_date = datetime.datetime.now() - datetime.timedelta(days=int(self.crawler["keep_years"]) * 365)

        remove_index = self.raw_dataframe[(pd.to_datetime(self.raw_dataframe['發現日期']) > start_date)  | (pd.to_datetime(self.raw_dataframe['發現日期']) < over_years_date)].index
        self.raw_dataframe.drop(remove_index, inplace=True)

    def _login(self):
        op = webdriver.ChromeOptions()
        # op.add_argument('headless')
        self.driver = webdriver.Chrome(options=op)

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

        date_info = handler.date_condition()

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
        self._clean_unless_data()

        # appended esh df to history dataframe
        new_dataframe = self.raw_dataframe.append(df, ignore_index=True)
        new_dataframe.to_excel(self.crawler["xlsx_name"], index=False)

    def run_crawler(self):
        try:
            self._login()
            self._esh_conditions()
            self._esh_crawler()
        except Exception as e:
            raise Exception(str(e))

        finally:
            self.driver.quit()

    def send_alert_mail(self):
        history_dataframe = self._read_history_esh()
        month_ago_date = datetime.datetime.now() - datetime.timedelta(days=int(self.crawler["alert_days"]))
        alert_data = (history_dataframe[
            (history_dataframe["表單狀態"] != "已結案") & (pd.to_datetime(history_dataframe['發現日期']) > month_ago_date)])
        alert_mail_links = alert_data["詳細資訊"].values.tolist()
        handler.send_mail(alert_mail_links[2])


if __name__ == "__main__":

    crawler = Crawler()
    crawler.send_alert_mail()






