# -*- coding: utf-8 -*-

import os
from datetime import datetime

import pandas as pd

from db_handler import DBHandler

MANAGER_NAME = u'חני'
MANAGER_ANSWER_ROW_KEY = u'הערות חני'
EMPLOYEE_ANSWER_ROW_KEY = u'תשובות לחני'
CUSTOMER_NAME_KEY_IN_OUTPUT_COMMENTS = 'customerName'
CUSTOMER_NAME_KEY = u'שם לקוח'
TEAM_NAME_KEY = u'צוות'
DATE_KEY = u'תאריך לביצוע'
OLD_COMMENTS_KEY = u'הערות מגיליון הגביה'
COMMENTS_KEY = 'comments'
DATE = 'date'
TEMP_FILE_NAME = "hst.txt"


class FileManager(object):
    def __init__(self):
        self.daily_report_df = pd.DataFrame()
        self.comments = []
        self.db_manager = DBHandler()

    def extract_all_comments_and_update_cem(self, daily_report_file_path, worker_name):
        self.load_file_to_df(daily_report_file_path).extract_all_comment(
            worker_name).store_comments_to_db().remove_users_file(daily_report_file_path)

    def load_file_to_df(self, daily_report_file_path):
        self.daily_report_df = pd.read_excel(daily_report_file_path)
        return self

    def extract_all_comment(self, worker_name):
        for index, row in self.daily_report_df.iterrows():
            if type(row[EMPLOYEE_ANSWER_ROW_KEY]) == unicode:
                customer_name = row[CUSTOMER_NAME_KEY]
                employee_comment = row[EMPLOYEE_ANSWER_ROW_KEY]
                manager_comment = row[MANAGER_ANSWER_ROW_KEY]
                date = datetime.now().strftime("%d/%m/%Y")
                employee_comment_format = date + " - " + worker_name + " : " + employee_comment
                manager_comment_format = date + " - " + MANAGER_NAME + " : " + manager_comment
                hs = open(TEMP_FILE_NAME, "a")
                hs.write(self.db_manager.get_old_comment(customer_name).encode('utf-8'))
                hs.write(employee_comment_format.encode('utf-8'))
                if type(row[MANAGER_ANSWER_ROW_KEY]) == unicode:
                    hs.write(manager_comment_format.encode('utf-8'))
                hs.close()
                hs = open(TEMP_FILE_NAME, 'r')
                self.comments.append(
                    {COMMENTS_KEY: hs.read(), CUSTOMER_NAME_KEY_IN_OUTPUT_COMMENTS: customer_name, DATE: row[DATE_KEY]})
                hs.close()
                os.remove(TEMP_FILE_NAME)
                return self

    def store_comments_to_db(self):
        for comment in self.comments:
            self.db_manager.inset_comment_to_db(comment[COMMENTS_KEY],
                                                comment[CUSTOMER_NAME_KEY_IN_OUTPUT_COMMENTS],
                                                comment[DATE])
        return self

    def remove_users_file(self, file_path):
        os.remove(file_path)
        return self
