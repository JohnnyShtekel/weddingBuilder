# -*- coding: utf-8 -*-

import os
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
TEMP_FILE_NAME = "hst.txt"


class FileManager(object):
    def __init__(self):
        self.daily_report_df = pd.DataFrame()
        self.comments = []
        self.db_manager = DBHandler()

    def extract_all_comments_and_update_cem(self, daily_report_file_path, workername):
        self.load_file_to_df(daily_report_file_path).extract_all_comment(
            workername).store_comments_to_db().remove_users_file(daily_report_file_path)

    def load_file_to_df(self, daily_report_file_path):
        self.daily_report_df = pd.read_excel(daily_report_file_path)
        return self

    def extract_all_comment(self, workername):
        for index, row in self.daily_report_df.iterrows():
            if (type(row[EMPLOYEE_ANSWER_ROW_KEY]) == unicode):
                customer_name = row[CUSTOMER_NAME_KEY]
                employee_comment = row[EMPLOYEE_ANSWER_ROW_KEY]
                manager_comment = row[MANAGER_ANSWER_ROW_KEY]
                date = "{d}/{m}/{y}".format(d=str(row[DATE_KEY].day), m=str(row[DATE_KEY].month),
                                            y=str(row[DATE_KEY].year))
                # old_comments = row[OLD_COMMENTS_KEY]
                employee_comment_format = date + " - " + workername + " : " + employee_comment
                manager_comment_format = date + " - " + MANAGER_NAME + " : " + manager_comment
                hs = open(TEMP_FILE_NAME, "a")
                hs.write(self.db_manager.get_old_comment(customer_name).encode('utf-8') + "\n")
                hs.write(employee_comment_format.encode('utf-8') + "\n")
                if type(row[MANAGER_ANSWER_ROW_KEY]) == unicode:
                    hs.write(manager_comment_format.encode('utf-8') + "\n")
                hs.close()
                hs = open(TEMP_FILE_NAME, 'r')
                self.comments.append(
                    {COMMENTS_KEY: hs.read(), CUSTOMER_NAME_KEY_IN_OUTPUT_COMMENTS: customer_name})
                hs.close()
                os.remove(TEMP_FILE_NAME)
                return self

    def store_comments_to_db(self):
        for comment in self.comments:
            self.db_manager.inset_comment_to_db(comment[COMMENTS_KEY], comment[CUSTOMER_NAME_KEY_IN_OUTPUT_COMMENTS])
        return self

    def remove_users_file(self, file_path):
        os.remove(file_path)
        return self
