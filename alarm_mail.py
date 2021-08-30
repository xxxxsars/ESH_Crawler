import datetime
import handler
import os
import pandas as pd

from bs4 import BeautifulSoup


class EmptyDataFrameException(Exception): pass


class Alarm_mail:
    def __init__(self, config_path, history_xlsx_path=None):
        self.config_path = config_path
        self.history_xlsx_path = history_xlsx_path
        self.config = handler.load_setting(self.config_path)
        self.crawler = self.config["Crawler"]
        self.detail_url_prefix = self.config["Link"]["detail_url"]
        self.lunch_date = datetime.datetime.strptime(self.config["System"]["lunch_date"], "%Y/%m/%d")
        self.mail_suffix = self.config["System"]["mail_suffix"]

        # If there is no default path, the 'xlsx name' parameter in the configuration file will be read.
        if history_xlsx_path is None:
            history_xlsx_path = self.crawler['xlsx_name']

        self.history_dataframe = handler.read_history_esh(self.history_xlsx_path)

        if self.history_dataframe.empty:
            raise EmptyDataFrameException(f'Please check your {self.crawler["xlsx_name"]} existed or had data.')

    def _mail(self, mail_subject: str, mail_type: str, detail_link: str, recipient_list: list[str]):

        html_path = f"templates/{mail_type}.html"

        if not os.path.exists(html_path):
            raise FileExistsError(f"Your '{html_path}' not existed. ")

        html = open(html_path, encoding='utf-8')
        html_content = html.read()

        soup = BeautifulSoup(html_content, 'html.parser')
        soup.find(id="detail-link", href=True)["href"] = detail_link

        mail_body = str(soup)

        handler.send_outlook_html_mail(recipients=recipient_list, subject=mail_subject, body=mail_body,
                                       send_or_display='SEND')

    def send_overdue_mail(self):
        mail_subject = '環安申請單逾期警告'

        month_ago_date = datetime.datetime.now() - datetime.timedelta(days=int(self.crawler["alert_days"]))
        alert_data = self.history_dataframe[
            (self.history_dataframe["表單狀態"] != "已結案") & (
                    pd.to_datetime(self.history_dataframe['發現日期']) < month_ago_date) & (
                    pd.to_datetime(self.history_dataframe['發現日期']) > self.lunch_date)]
        system_ids = alert_data["異常單系統編號"].values.tolist()
        recipient_list = [department + self.mail_suffix for department in alert_data["責任單位"].values.tolist()]
        if not alert_data.empty:
            for index in range(len(alert_data)):
                detail_url = self.detail_url_prefix + str(system_ids[index])
                self._mail(mail_subject, "overdue", detail_url, [recipient_list[index]])
        else:
            print("All esh form had been done.")

    def send_wrong_format_mail(self):
        mail_subject = '環安申請單格式錯誤警告'
        alert_data = (self.history_dataframe[
            (self.history_dataframe["關鍵字"] != "ok") & (
                    pd.to_datetime(self.history_dataframe['發現日期']) > self.lunch_date)])
        recipient_list = [department + self.mail_suffix for department in alert_data["提出單位"].values.tolist()]
        system_ids = alert_data["異常單系統編號"].values.tolist()

        if not alert_data.empty:
            for index in range(len(alert_data)):
                detail_url = self.detail_url_prefix + str(system_ids[index])
                self._mail(mail_subject, "wrong_format", detail_url, [recipient_list[index]])
        else:
            print("All esh table formats are correct.")


if __name__ == "__main__":
    a = Alarm_mail("setting.ini")
    a.send_wrong_format_mail()
