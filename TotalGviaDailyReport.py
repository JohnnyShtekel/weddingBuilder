# coding=utf-8
import pandas as pd
from esg_services import EsgServiceManager
import numpy as np
from StyleFrame import StyleFrame, colors
import math
from pandas import concat
import re
from openpyxl import load_workbook


class TotalGviaDailyReport(object):
    def __init__(self):
        self.manager = EsgServiceManager()
        self.df = pd.DataFrame()
        self.rows_of_sum = []
        self.df_from_sql_for_gvia_megvia_and_tzefi = np.nan


    def add_col_team(self):
        query = '''SELECT NameTeam [צוות]
            FROM dbo.tblTeamList
            WHERE NameTeam  NOT LIKE 'כללי'
            ORDER BY NameTeam'''
        q = self.manager.db_service.search(query=query)
        self.df = pd.DataFrame.from_records(q)

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
                ORDER BY NameTeam'''
        q = self.manager.db_service.search(query=query)
        self.df_from_sql_for_gvia_megvia_and_tzefi = pd.DataFrame.from_records(q)

    def add_col_gvia_from_gvia(self):
        df = (self.df_from_sql_for_gvia_megvia_and_tzefi.groupby('NameTeam').PaySuccessFU.sum()).reset_index()
        dict_for_consultation = df.set_index('NameTeam')['PaySuccessFU'].to_dict()
        self.df[u'גביה מהגביה'] = np.nan
        for row in range(0, len(self.df)):
            if self.df[u'צוות'][row] in dict_for_consultation.keys():
                self.df.set_value(row, u'גביה מהגביה',
                                  round(dict_for_consultation.get(self.df[u'צוות'][row])))


    def add_col_tzefi_from_gvia(self):
        df = (self.df_from_sql_for_gvia_megvia_and_tzefi.groupby('NameTeam').PaySuccessFUTeam.sum()).reset_index()
        dict_for_consultation = df.set_index('NameTeam')['PaySuccessFUTeam'].to_dict()
        self.df[u'צפי של הגביה / צוות'] = np.nan
        for row in range(0, len(self.df)):
            if self.df[u'צוות'][row] in dict_for_consultation.keys():
                self.df.set_value(row, u'צפי של הגביה / צוות',
                                  round(dict_for_consultation.get(self.df[u'צוות'][row])))

    def add_col_total_gvia_plus_tzefi_gvia_team(self):
        list_for_column = []
        for row in range(0, len(self.df)):
            list_for_column.append("=E{row}+G{row}".format(row=row + 2))
        self.df[u'גביה כוללת + צפי של הגבייה / צוות'] = list_for_column

    def add_col_target(self):
        list_of_targets = []
        list_of_indexes_for_teams_targets = [3, 2, 1, 10, 6, 4, 5, 7, 8]
        query = '''SELECT *
        FROM dbo.tblMonthlyTarget
        WHERE year = YEAR(getdate())AND month = MONTH(getdate())'''
        q = self.manager.db_service.search(query=query)
        df = pd.DataFrame.from_records(q)
        for index in list_of_indexes_for_teams_targets:
            list_of_targets.append(df['sumOfTarget{index}'.format(index=index)][0])
        self.df[u'יעד'] = list_of_targets

    def add_col_percentage_of_completion(self):
        self.df[u'אחוז ביצוע'] = np.nan
        for row in range(0, len(self.df)):
            self.df[u'אחוז ביצוע'][row] = "=(IF(I{row}>0,E{row}/I{row},0))".format(row=row + 2)


    def add_col_percentage_of_completion_including_tzefi(self):
        self.df[u'אחוז ביצוע כולל צפי'] = np.nan
        for row in range(0, len(self.df)):
            self.df[u'אחוז ביצוע כולל צפי'][row] = "=(IF(I{row}>0,H{row}/I{row},0))".format(row=row + 2)

    def add_row_for_yesum(self):
        formula = "=SUM(({col}2:{col}4),({col}6,{col}8,{col}10))"
        new_yesum_row = ({u'צוות': [u'סה"כ יישום'], u'גביה מרגיל': [formula.format(col='B')],
                          u'גביה מייעוץ': [formula.format(col='C')], u'גביה מפרויקטלי': [formula.format(col='D')],
                          u'גביה כוללת': [formula.format(col='E')], u'גביה מהגביה': [formula.format(col='F')],
                          u'צפי של הגביה / צוות': [formula.format(col='G')],
                          u'גביה כוללת + צפי של הגבייה / צוות': [formula.format(col='H')],
                          u'יעד': [formula.format(col='I')], u'אחוז ביצוע': [formula.format(col='J')],
                          u'אחוז ביצוע כולל צפי': [formula.format(col='K')]})
        print new_yesum_row
        self.df = concat([self.df, new_yesum_row]).reset_index(drop=True)

    def order_columns(self):
        self.df = self.df[[u'צוות', u'גביה מרגיל', u'גביה מייעוץ', u'גביה מפרויקטלי', u'גביה כוללת', u'גביה מהגביה',
                           u'צפי של הגביה / צוות', u'גביה כוללת + צפי של הגבייה / צוות', u'יעד', u'אחוז ביצוע',
                           u'אחוז ביצוע כולל צפי']]

    def change_types_of_cells(self):
        wb = load_workbook('total gvia.xlsx')
        ws = wb.active
        for i in range(2, 11):
            ws['J{i}'.format(i=i)].number_format = '0%'
            ws['K{i}'.format(i=i)].number_format = '0%'
            ws['B{i}'.format(i=i)].number_format = '#,###'
            ws['C{i}'.format(i=i)].number_format = '#,###'
            ws['D{i}'.format(i=i)].number_format = '#,###'
            ws['E{i}'.format(i=i)].number_format = '#,###'
            ws['F{i}'.format(i=i)].number_format = '#,###'
            ws['G{i}'.format(i=i)].number_format = '#,###'
            ws['H{i}'.format(i=i)].number_format = '#,###'
            ws['I{i}'.format(i=i)].number_format = '#,###'
        wb.save('total gvia.xlsx')

gvia_teams = TotalGviaDailyReport()
gvia_teams.add_col_team()
gvia_teams.add_gvia_cols_for_3_kinds_of_customers()
gvia_teams.add_col_total_gvia()
gvia_teams.get_df_from_sql_for_gvia_megvia_and_tzefi()
gvia_teams.add_col_gvia_from_gvia()
gvia_teams.add_col_tzefi_from_gvia()
gvia_teams.add_col_total_gvia_plus_tzefi_gvia_team()
gvia_teams.add_col_target()
gvia_teams.add_col_percentage_of_completion()
gvia_teams.add_col_percentage_of_completion_including_tzefi()
# gvia_teams.add_row_for_yesum()
gvia_teams.order_columns()


sf = StyleFrame(gvia_teams.df)
writer = StyleFrame.ExcelWriter("total gvia.xlsx")
sf.to_excel(excel_writer=writer, sheet_name=u'דוח גביה יומי לפי צוותים', right_to_left=True, row_to_add_filters=0,
            columns_and_rows_to_freeze='A2')
writer.save()
gvia_teams.change_types_of_cells()








