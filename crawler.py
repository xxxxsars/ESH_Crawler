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

# handle the characters that fail to convert xlsx file to csv .
ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]|\n|\xa0')


class Crawler:
    def __init__(self, config_path, debug=False):
        # Debug mode will show the browser and didn't close the final browser window
        self.config_path = config_path
        self.debug = debug
        self.config = handler.load_setting(config_path)

        # initial selenium driver
        self.driver = None

    def __enter__(self):
        # read the setting.ini data
        self.link = self.config["Link"]
        self.user = self.config["User"]
        self.crawler = self.config["Crawler"]

        # read the history esh xlsx file data
        self.raw_dataframe = handler.read_history_esh(self.config_path)

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.driver and self.debug == False:
            self.driver.close()
            self.driver.quit()
        if exc_type == None and exc_val == None and exc_tb == None:
            sys.exit(0)

    def _clean_unless_data(self):

        start_date = datetime.datetime.strptime(handler.date_condition(int(self.crawler["days_ago"]))["start_date"],
                                                "%Y/%m/%d")
        over_years_date = datetime.datetime.now() - datetime.timedelta(days=int(self.crawler["keep_years"]) * 365)

        remove_index = self.raw_dataframe[(pd.to_datetime(self.raw_dataframe['發現日期']) > start_date) | (
                pd.to_datetime(self.raw_dataframe['發現日期']) < over_years_date)].index
        self.raw_dataframe.drop(remove_index, inplace=True)

    def _login(self):
        op = webdriver.ChromeOptions()
        if self.debug == False:
            op.add_argument('headless')
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
        input_fab.select_by_value("117")

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

        column: list[str] = []
        rows: list[str] = []

        for index, tr in enumerate(table_element.find("tbody").find_all("tr")):
            row: list[str] = []

            for td in tr.find_all('td'):
                value = ILLEGAL_CHARACTERS_RE.sub(r'', td.text)
                if value == "View":
                    form_id = re.search("\w+$", (td.find("a", href=True)['href'])).group(0)
                    row.append(form_id)
                else:
                    row.append(value)

            # The first line was the column name
            if index == 0:
                column = row
                column[0] = "異常單系統編號"

                # cut the category to category and sub category
                category_index = column.index("異常類別")
                sub_category_index = category_index + 1
                column.insert(sub_category_index, "異常子類別")
            else:
                category_value = (row[category_index]).split("-")
                row[category_index] = category_value[0]
                row.insert(sub_category_index, category_value[1])
                rows.append(row)

        df = pd.DataFrame(rows, columns=column)
        return df

    def _add_status(self, crawler_df):
        with open("keyword.json", encoding="utf-8") as fin:
            allow_keyword_prefix = list(json.load(fin).keys())
            description_status = []
            for description in (crawler_df["異常現象敍述"]):
                match = re.search("^(\[.+\])", description)
                if match:
                    if match.group(0) in allow_keyword_prefix:
                        description_status.append("ok")
                    else:
                        description_status.append("NG-關建字")
                else:
                    description_status.append("NG-標示")

            count_rows = len(crawler_df.index)
            reply_days = crawler_df["回覆天數"].values.tolist()
            form_status = crawler_df["表單狀態"].values.tolist()
            status = ["OK" for _ in range(count_rows)]

            crawler_df['1'] = reply_days
            crawler_df['2'] = form_status
            crawler_df['3'] = status
            crawler_df["關鍵字"] = description_status

            return crawler_df

    def run_crawler(self):
        self._login()
        self._esh_conditions()
        crawler_df = self._esh_crawler()
        crawler_df = self._add_status(crawler_df)

        if self.raw_dataframe.empty == False:
            self._clean_unless_data()

        # TODO: if saved excel file failed, it will save another file name with timestamp
        # appended esh df to history dataframe
        new_dataframe = self.raw_dataframe.append(crawler_df, ignore_index=True)
        new_dataframe.to_excel(self.crawler["xlsx_name"], sheet_name='ESH_RawData', index=False)

        print("crawler running had been done!")

if __name__ == "__main__":
    with Crawler("setting.ini") as cw:
        cw.run_crawler()
