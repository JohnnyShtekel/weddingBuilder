# coding=utf-8
import pandas as pd
from esg_services import EsgServiceManager
import numpy as np
from StyleFrame import StyleFrame, colors
from pandas import concat
import re



class TotalGviaDailyReport(object):
    def __init__(self):
        self.manager = EsgServiceManager()
        self.df = pd.DataFrame()
        self.rows_of_sum = []
        self.df_from_sql_for_gvia_megvia_and_tzefi = np.nan


    def add_col_team(self):
        df = pd.read_excel('Gvia day report new.xlsx', convert_float=True, sheetname=u'דוח גביה יומי חיובי')
        df = df[[u'צוות']]
        df = df.drop_duplicates(u'צוות')
        list_of_teams = df[u'צוות'].tolist()
        # list_of_teams = [u'ברקת', u'אלמוג', u'אודם', u'פנינה', u'קריסטל', u'שוהם-שכר', u'שנהב', u'ספיר', u'טורקיז',
        #         u'סה"כ יישום', u'מחלקת גביה']
        self.df[u'צוות'] = list_of_teams

    def get_paid_payment_for_consultation_from_excel_daily_report(self):
        df = pd.read_excel('Gvia day report new.xlsx', convert_float=True, sheetname=u'דוח גביה יומי חיובי')
        df = df[[u'צוות', u'שם לקוח', u'סוג לקוח', u'תשלום ששולם עד היום לייעוץ']]
        list_of_sum_rows = []
        for row in range(0, len(df)):
            if df[u'שם לקוח'][row].startswith(u'סיכום'):
                list_of_sum_rows.append(row)
        df.drop(df.index[list_of_sum_rows], inplace=True)
        df.index = range(len(df))
        return df

    def add_gvia_cols_for_3_kinds_of_customers(self):
        df = self.get_paid_payment_for_consultation_from_excel_daily_report()
        list_for_regular_customer_col = []
        list_for_consultation_customer_col = []
        list_for_project_customer_col = []
        team = df[u'צוות'][0]
        regular_sum = 0
        consultation_sum = 0
        project_sum = 0
        for row in range(0, len(df)):
            if df[u'צוות'][row] == team:
                if str(df[u'תשלום ששולם עד היום לייעוץ'][row]) == "nan":
                    df[u'תשלום ששולם עד היום לייעוץ'][row] = 0.0
                if df[u'סוג לקוח'][row] == u'ES - בסיס הצלחה':
                    regular_sum += df[u'תשלום ששולם עד היום לייעוץ'][row]
                elif df[u'סוג לקוח'][row] == u'ES-יעוץ':
                    consultation_sum += df[u'תשלום ששולם עד היום לייעוץ'][row]
                elif df[u'סוג לקוח'][row] == u'פרוייקטלי':
                    project_sum += df[u'תשלום ששולם עד היום לייעוץ'][row]
            if row == len(df) - 1:
                list_for_regular_customer_col.append(regular_sum)
                list_for_consultation_customer_col.append(consultation_sum)
                list_for_project_customer_col.append(project_sum)
            elif df[u'צוות'][row + 1] != team:
                team = df[u'צוות'][row + 1]
                list_for_regular_customer_col.append(regular_sum)
                list_for_consultation_customer_col.append(consultation_sum)
                list_for_project_customer_col.append(project_sum)
                regular_sum = 0
                consultation_sum = 0
                project_sum = 0
        self.df[u'גביה מרגיל'] = list_for_regular_customer_col
        self.df[u'גביה מייעוץ'] = list_for_consultation_customer_col
        self.df[u'גביה מפרויקטלי'] = list_for_project_customer_col

    def add_col_total_gvia(self):
        list_of_formulas = []
        for row in range(0, len(self.df)):
            formula = "=B{row}+C{row}+D{row}".format(row=row + 2)
            list_of_formulas.append(formula)
        self.df[u'גביה כוללת'] = list_of_formulas

    def get_df_from_sql_for_gvia_megvia_and_tzefi(self):
        query = '''SELECT PaySuccessFU, PaySuccessFUTeam, DateFU, dbo.tblCustomers.KodCustomer, NameCustomer,NameTeam
          FROM dbo.tbltaxes
          INNER JOIN dbo.tblCustomers
            ON dbo.tbltaxes.KodCustomer = dbo.tblCustomers.KodCustomer
          INNER JOIN dbo.tblTeamList
            ON dbo.tblCustomers.KodTeamCare = dbo.tblTeamList.KodTeamCare
            WHERE CustomerStatus = 2 AND YEAR(DateFU) = YEAR(getdate()) AND MONTH(DateFU) = MONTH(getdate())
             AND DAY(DateFU) <= DAY(getdate())
                ORDER BY NameTeam'''
        q = self.manager.db_service.search(query=query)
        self.df_from_sql_for_gvia_megvia_and_tzefi = pd.DataFrame.from_records(q)
        self.df_from_sql_for_gvia_megvia_and_tzefi.to_excel("gvia megvia.xlsx")

    def add_col_gvia_from_gvia(self):
        df = (self.df_from_sql_for_gvia_megvia_and_tzefi.groupby('NameTeam').PaySuccessFU.sum()).reset_index()
        df.to_excel("gvia megvia1.xlsx")
        dict_for_consultation = df.set_index("NameTeam")["PaySuccessFU"].to_dict()
        self.df[u'גביה מהגביה'] = np.nan
        for index in range(0, len(self.df)):
            if self.df.iloc[index]["NameTeam"] in dict_for_consultation.keys():
                self.df.set_value(index, u'גביה מהגביה',
                                  dict_for_consultation.get(self.df.iloc[index]["NameTeam"]))


    def add_col_tzefi_from_gvia(self):
        df = (self.df_from_sql_for_gvia_megvia_and_tzefi.groupby('NameTeam').PaySuccessFUTeam.sum()).reset_index()
        df.to_excel("gvia megvia2.xlsx")


    def order_columns(self):
        self.df = self.df[[u'צוות', u'גביה מרגיל', u'גביה מייעוץ', u'גביה מפרויקטלי', u'גביה כוללת']]



gvia_teams = TotalGviaDailyReport()
gvia_teams.add_col_team()
gvia_teams.add_gvia_cols_for_3_kinds_of_customers()
gvia_teams.add_col_total_gvia()
gvia_teams.order_columns()
gvia_teams.get_df_from_sql_for_gvia_megvia_and_tzefi()
# gvia_teams.add_col_gvia_from_gvia()
# gvia_teams.add_col_tzefi_from_gvia()
gvia_teams.df.to_excel("total gvia.xlsx")

sf = StyleFrame(gvia_teams.df)
writer = StyleFrame.ExcelWriter("total gvia.xlsx")
sf.to_excel(excel_writer=writer, sheet_name=u'דוח גביה יומי לפי צוותים', right_to_left=True, row_to_add_filters=0,
            columns_and_rows_to_freeze='A2')
writer.save()





