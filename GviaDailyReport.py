# coding=utf-8
import pandas as pd
from esg_services import EsgServiceManager
from StyleFrame import StyleFrame, colors
from datetime import datetime
import numbers
import time

import numpy as np
from pandas import concat



class GviaDailyReport(object):
    def __init__(self):
        self.manager = EsgServiceManager()
        self.df = pd.read_excel('gvia_month.xlsx', convert_float=True)
        self.rows_of_sum = []


    def remove_sum_rows(self):
        list_of_sum_rows = []
        for row in range(0, len(self.df)):
            if self.df[u'שם לקוח'][row].startswith(u'סיכום'):
                list_of_sum_rows.append(row)
        self.df.drop(self.df.index[list_of_sum_rows], inplace=True)
        self.df.index = range(len(self.df))


    def calc_month_tzefi(self):
        month_tzefi_list = []
        for row in range(0, len(self.df)):
            if self.df[u'לקוח משפטי'][row] == u'כן':
                month_tzefi_list.append(0)
            else:
                if self.df[u'מספר התשלומים מעבר לאשראי'][row] > 3:
                    month_tzefi_list.append(0)
                else:
                    month_tzefi_list.append(self.df[u'תשלום ייעוץ חודשי'][row])
        return month_tzefi_list

    def calc_month_extra_payments(self):
        month_extra_payments_list = []
        for row in range(0, len(self.df)):
            if self.df[u'לקוח משפטי'][row] == u'כן':
                month_extra_payments_list.append(0)
            else:
                month_extra_payments_list.append(self.df[u'תשלום על בונוס'][row] + self.df[u'תשלומים מיוחדים'][row] +
                                                 self.df[u'סכום התשלומים מעבר לאשראי'][row])
        return month_extra_payments_list

    def add_col_month_tzefi(self):
        month_tzefi_list = []
        for tzefi, extra in zip(self.calc_month_tzefi(), self.calc_month_extra_payments()):
            month_tzefi_list.append(tzefi + extra)
        self.df[u'צפי לחודש'] = month_tzefi_list


    def get_dict_for_feeTeam(self):
        dates = []
        query = '''SELECT NameCustomer, feeTeam, DatePay
                FROM dbo.tblCustomers LEFT JOIN dbo.tbltaxesPay
                  ON dbo.tbltaxesPay.KodCustomer = dbo.tblCustomers.KodCustomer'''
        q = self.manager.db_service.search(query=query)
        df = pd.DataFrame.from_records(q)
        for index, row in df.iterrows():
            if row['DatePay'] != None:
                date = datetime(row['DatePay'][0], row['DatePay'][1], row['DatePay'][2],
                                row['DatePay'][3], row['DatePay'][4], row['DatePay'][5])
                dates.append(date)
            else:
                date = datetime(1920, 1, 1, 00, 00, 00)
                dates.append(date)
        df['date_pay'] = dates
        df['date_pay'] = pd.to_datetime(df['date_pay'])
        df.drop(['DatePay'], axis=1, inplace=True)
        today = date.today()
        month = today.month
        year = today.year
        first_of_month = date.replace(year, month, 1)
        df = df[df.date_pay <= today]
        df = df[df.date_pay >= first_of_month]
        df_grouped = pd.DataFrame(df.groupby('NameCustomer').feeTeam.sum()).reset_index()
        dict_for_feeTeam = df_grouped.set_index('NameCustomer')['feeTeam'].to_dict()
        return dict_for_feeTeam

    def add_col_paid_payment_for_consulation(self):
        dict_of_feeTeam = self.get_dict_for_feeTeam()
        self.df[u'תשלום ששולם עד היום לייעוץ - חלוקת הגביה'] = np.nan
        for index in range(0, len(self.df)):
            if self.df.iloc[index][u'שם לקוח'] in dict_of_feeTeam.keys():
                self.df.set_value(index, u'תשלום ששולם עד היום לייעוץ - חלוקת הגביה',
                                  dict_of_feeTeam.get(self.df.iloc[index][u'שם לקוח']))


    def add_col_tzefi_left(self, status):
        list_for_col = []
        rows = len(self.df.index)
        if status == 1:
            for index in range(0, rows):
                if index not in self.rows_of_sum:
                    list_for_col.append('=IF(F{row}-G{row}<0,0,F{row}-G{row})'.format(row=index + 2))
                else:
                    list_for_col.append(self.df.iloc[index][u'צפי שנותר'])
        else:
            for index in range(0, rows):
                list_for_col.append(0)
        self.df[u'צפי שנותר'] = list_for_col


    def remove_needless_columns(self):
        self.df = self.df[[u'צוות', u'שם לקוח', u'סוג לקוח',  u'לקוח משפטי', u'תשלום ייעוץ חודשי',
                           u'צפי לחודש', u'תשלום ששולם עד היום לייעוץ - חלוקת הגביה', u'צפי שנותר',
                           u'הערות מגיליון הגביה', u'הערות']]



    def change_type_of_col_to_int(self, col_name):
        self.df[col_name] = self.df[col_name].apply(lambda x: int(x))


    def change_type_of_sums_col_to_int(self):
        self.change_type_of_col_to_int(u'תשלום ייעוץ חודשי')
        self.change_type_of_col_to_int(u'צפי לחודש')
        self.change_type_of_col_to_int(u'תשלום ששולם עד היום לייעוץ - חלוקת הגביה')
        self.change_type_of_col_to_int(u'צפי שנותר')

    def add_mid_sums(self):
        rows = len(self.df.index)
        team = self.df[u'צוות'][0]
        first_index = 0
        last_index = 0
        for r in range(0, rows):
            if r == rows - 1:
                last_index = len(self.df.index) - 1
                self.calc_mid_sums(first_index, last_index, team)
            if self.df[u'צוות'][r + 1] == team:
                last_index = last_index + 1

            else:
                if self.df[u'צוות'][r + 1] != team:
                    new_team = self.df[u'צוות'][r + 1]
                    self.calc_mid_sums(first_index, last_index, team)
                    last_index = last_index + 1
                    first_index = last_index + 1
                    team = new_team
        self.df.sort_index()

    def calc_mid_sums(self, first_index, last_index, team):
        index1 = first_index + 2
        index2 = last_index + 2
        sum_of_month_consultation = '=SUM(E{first_index}:E{last_index})'.format(first_index=index1, last_index=index2)
        sum_of_month_zefi = '=SUM(F{first_index}:F{last_index})'.format(first_index=index1, last_index=index2)
        sum_of_paid_payment_for_consulation = '=SUM(G{first_index}:G{last_index})'.format(first_index=index1,
                                                                                      last_index=index2)
        sum_of_zefi_left = '=SUM(H{first_index}:H{last_index})'.format(first_index=index1, last_index=index2)
        new_sum_line = pd.DataFrame({u'צוות': [team], u'שם לקוח': [u'סיכום לצוות {team}'.format(team=team)],
                                     u'סוג לקוח': [""],  u'לקוח משפטי': [""],
                                     u'תשלום ייעוץ חודשי': [sum_of_month_consultation],
                                     u'צפי לחודש': [sum_of_month_zefi],
                                     u'תשלום ששולם עד היום לייעוץ - חלוקת הגביה': [sum_of_paid_payment_for_consulation],
                                     u'צפי שנותר': [sum_of_zefi_left],
                                     u'הערות מגיליון הגביה': [""], u'הערות': [""]})
        new_df = concat([self.df.ix[:last_index], new_sum_line, self.df.ix[last_index + 1:]]).reset_index(drop=True)
        new_df.sort_index()
        self.df = new_df
        self.rows_of_sum.append(last_index + 1)

    def apply_style_on_sum_rows(self, sf):
        for row in self.rows_of_sum:
            sf.apply_style_by_indexes(indexes_to_style=[sf.index[row]], bg_color=colors.yellow, bold=True)

    def add_total_row(self):
        sum_of_month_consultation = '=E{r1}+E{r2}+E{r3}+E{r4}+E{r5}+E{r6}+E{r7}+E{r8}+E{r9}'.format(
            r1=self.rows_of_sum[0] + 2, r2=self.rows_of_sum[1] + 2, r3=self.rows_of_sum[2] + 2,
            r4=self.rows_of_sum[3] + 2, r5=self.rows_of_sum[4] + 2, r6=self.rows_of_sum[5] + 2,
            r7=self.rows_of_sum[6] + 2, r8=self.rows_of_sum[7] + 2, r9=self.rows_of_sum[8] + 2)
        sum_of_month_zefi = '=F{r1}+F{r2}+F{r3}+F{r4}+F{r5}+F{r6}+F{r7}+F{r8}+F{r9}'.format(
            r1=self.rows_of_sum[0] + 2, r2=self.rows_of_sum[1] + 2, r3=self.rows_of_sum[2] + 2,
            r4=self.rows_of_sum[3] + 2, r5=self.rows_of_sum[4] + 2, r6=self.rows_of_sum[5] + 2,
            r7=self.rows_of_sum[6] + 2, r8=self.rows_of_sum[7] + 2, r9=self.rows_of_sum[8] + 2)
        sum_of_paid_payment_for_consulation = '=G{r1}+G{r2}+G{r3}+G{r4}+G{r5}+G{r6}+G{r7}+G{r8}+G{r9}'.format(
            r1=self.rows_of_sum[0] + 2, r2=self.rows_of_sum[1] + 2, r3=self.rows_of_sum[2] + 2,
            r4=self.rows_of_sum[3] + 2, r5=self.rows_of_sum[4] + 2, r6=self.rows_of_sum[5] + 2,
            r7=self.rows_of_sum[6] + 2, r8=self.rows_of_sum[7] + 2, r9=self.rows_of_sum[8] + 2)
        sum_of_zefi_left = '=H{r1}+H{r2}+H{r3}+H{r4}+H{r5}+H{r6}+H{r7}+H{r8}+H{r9}'.format(
            r1=self.rows_of_sum[0] + 2, r2=self.rows_of_sum[1] + 2, r3=self.rows_of_sum[2] + 2,
            r4=self.rows_of_sum[3] + 2, r5=self.rows_of_sum[4] + 2, r6=self.rows_of_sum[5] + 2,
            r7=self.rows_of_sum[6] + 2, r8=self.rows_of_sum[7] + 2, r9=self.rows_of_sum[8] + 2)
        new_sum_line = pd.DataFrame({u'צוות': [""], u'שם לקוח': [u'סיכום לכל הצוותים:'],
                                     u'סוג לקוח': [""],  u'לקוח משפטי': [""],
                                     u'תשלום ייעוץ חודשי': [sum_of_month_consultation],
                                     u'צפי לחודש': [sum_of_month_zefi],
                                     u'תשלום ששולם עד היום לייעוץ - חלוקת הגביה': [sum_of_paid_payment_for_consulation],
                                     u'צפי שנותר': [sum_of_zefi_left],
                                     u'הערות מגיליון הגביה': [""], u'הערות': [""]})
        new_df = concat([self.df, new_sum_line]).reset_index(drop=True)
        self.df = new_df



gvia_day_manager = GviaDailyReport()
gvia_day_manager.remove_sum_rows()
gvia_day_manager.add_col_tzefi_left(0)
gvia_day_manager.add_col_paid_payment_for_consulation()
gvia_day_manager.add_col_month_tzefi()
gvia_day_manager.add_mid_sums()
gvia_day_manager.rows_of_sum.append(len(gvia_day_manager.df))
gvia_day_manager.add_total_row()
gvia_day_manager.add_col_tzefi_left(1)
gvia_day_manager.remove_needless_columns()
sf = StyleFrame(gvia_day_manager.df)
gvia_day_manager.apply_style_on_sum_rows(sf)
writer = StyleFrame.ExcelWriter('Gvia day report.xlsx')
sf.to_excel(excel_writer=writer, right_to_left=True, row_to_add_filters=0, columns_and_rows_to_freeze='A2')
writer.save()
