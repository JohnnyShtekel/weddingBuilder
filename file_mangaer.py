# -*- coding: utf-8 -*-

import os
import datetime
import math
import pandas as pd

from db_handler import DBHandler

MANAGER_NAME = u'חני'
MANAGER_ANSWER_ROW_KEY = u'הערות לחני/אורטל'
EMPLOYEE_ANSWER_ROW_KEY = u'תשובות לחני/אורטל'
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
        self.db_manager = DBHandler()
        self.date = None

    def extract_all_comments_and_update_cem(self, daily_report_file_path, worker_name, manager_name):
        self.load_file_to_df(daily_report_file_path)
        self.extract_date_from_df(worker_name, manager_name)
        self.remove_users_file(daily_report_file_path)

    def load_file_to_df(self, daily_report_file_path):
        self.daily_report_df = pd.read_excel(daily_report_file_path)

    def extract_date_from_df(self, worker_name, manager_name):
        for index in range(len(self.daily_report_df)):
            customer_name = self.daily_report_df.loc[index][CUSTOMER_NAME_KEY]
            employee_comment = self.daily_report_df.loc[index][EMPLOYEE_ANSWER_ROW_KEY]
            manager_comment = self.daily_report_df.loc[index][MANAGER_ANSWER_ROW_KEY]
            date_for_action = self.daily_report_df.loc[index][DATE_KEY]
            date_for_action = '-'.join([str(date_for_action.year), str(date_for_action.month), str(date_for_action.day)]) + ' ' + '00:00:00'
            employee_answer_not_empty = type(
                self.daily_report_df.loc[index][EMPLOYEE_ANSWER_ROW_KEY]) == unicode or not math.isnan(
                self.daily_report_df.loc[index][EMPLOYEE_ANSWER_ROW_KEY])
            manager_answer_not_empty = type(
                self.daily_report_df.loc[index][MANAGER_ANSWER_ROW_KEY]) == unicode or not math.isnan(
                self.daily_report_df.loc[index][MANAGER_ANSWER_ROW_KEY])
            self.store_comments_to_db(customer_name, employee_comment, manager_comment, datetime.datetime.now(), employee_answer_not_empty, manager_answer_not_empty, worker_name, date_for_action, manager_name)



    def convert_pandas_date_to_datetime(self, pandas_date):
        return datetime.datetime(pandas_date.year, pandas_date.month, pandas_date.day)


    def store_comments_to_db(self, customer_name, employee_comment, manager_comment, date, employee_answer_not_empty, manager_answer_not_empty, worker_name, date_for_action, manager_name):
        date_string = date.strftime("%d/%m/%Y")
        # print self.db_manager.get_old_comment(customer_name)
        # print "*********************************************"
        if employee_answer_not_empty and manager_answer_not_empty:
            employee_comment_format = date_string + " - " + worker_name + " : " + employee_comment
            manager_comment_format = date_string + " - " + manager_name + " : " + manager_comment
            comments = '\n'.join(
                [self.db_manager.get_old_comment(customer_name).encode('utf-8'),
                 employee_comment_format.encode('utf-8'), manager_comment_format.encode('utf-8')])
            self.db_manager.inset_comment_to_db(comments, customer_name, date_for_action)
        elif employee_answer_not_empty and not manager_answer_not_empty:
            employee_comment_format = date_string + " - " + worker_name + " : " + employee_comment
            comments = '\n'.join(
                [self.db_manager.get_old_comment(customer_name).encode('utf-8'),
                 employee_comment_format.encode('utf-8')])
            self.db_manager.inset_comment_to_db(comments, customer_name, date_for_action)

        elif not employee_answer_not_empty and manager_answer_not_empty:
            manager_comment_format = date_string + " - " + manager_name + " : " + manager_comment
            comments = '\n'.join(
                [self.db_manager.get_old_comment(customer_name).encode('utf-8'),
                 manager_comment_format.encode('utf-8')])
            self.db_manager.inset_comment_to_db(comments, customer_name, date_for_action)


    def remove_users_file(self, file_path):
        os.remove(file_path)
