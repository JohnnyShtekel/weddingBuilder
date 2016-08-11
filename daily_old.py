# coding=utf-8
import pandas as pd
from esg_services import EsgServiceManager
from StyleFrame import StyleFrame, colors
from datetime import datetime
import math
import re
import numpy as np
from pandas import concat



class GviaDailyReport(object):
    def __init__(self):
        self.manager = EsgServiceManager()
        self.df = pd.read_excel('yki.xlsx', convert_float=True, sheetname=u'צפי לחודש אוגוסט 2016')
        # self.df = pd.read_excel('gvia_month.xlsx', convert_float=True)
        self.rows_of_sum_positive = []
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

    def add_col_month_tzefi_from_original_month_report(self):
        month_tzefi_list = []
        for tzefi, extra in zip(self.calc_month_tzefi(), self.calc_month_extra_payments()):
            month_tzefi_list.append(tzefi + extra)
        self.df[u'צפי לחודש'] = month_tzefi_list

    def add_col_month_tzefi(self):
        if str(self.df[u'צפי לחודש'][0]) == "nan":
            self.add_col_month_tzefi_from_original_month_report()
        else:
            month_tzefi_list = []
            for row in range(0, len(self.df)):
                month_tzefi_list.append(math.ceil((self.df[u'צפי מאוחד'][row])))
                # month_tzefi_list.append(self.df[u'צפי מאוחד'][row])
            self.df[u'צפי לחודש'] = month_tzefi_list

    def get_df_of_fees(self):
        dates = []
        query = '''SELECT NameCustomer, feeTeam, DatePay, feeTax
                FROM dbo.tblCustomers LEFT JOIN dbo.tbltaxesPay
                  ON dbo.tbltaxesPay.KodCustomer = dbo.tblCustomers.KodCustomer'''
        q = self.manager.db_service.search(query=query)
        df = pd.DataFrame.from_records(q)
        today = datetime.today()
        for index, row in df.iterrows():
            if row['DatePay'] is not None and row['DatePay'][0] == today.year and row['DatePay'][1] == today.month:
                date = datetime(row['DatePay'][0], row['DatePay'][1], row['DatePay'][2],
                                row['DatePay'][3], row['DatePay'][4], row['DatePay'][5])
                dates.append(date)
            else:
                date = datetime(1920, 1, 1, 00, 00, 00)
                dates.append(date)
        df['date_pay'] = dates
        df.drop(['DatePay'], axis=1, inplace=True)
        old_date_for_compare = datetime(1920, 1, 1, 00, 00, 00)
        df = df[df.date_pay > old_date_for_compare]
        df.drop(['date_pay'], axis=1, inplace=True)
        return df

    def add_cols_for_fees(self):
        df = self.get_df_of_fees()
        self.add_col_paid_payment_for_consulation(df)
        self.add_col_paid_payment_for_gvia(df)

    def get_dict_for_fee_team(self, df):
        df_grouped = pd.DataFrame(df.groupby('NameCustomer').feeTeam.sum()).reset_index()
        dict_for_feeTeam = df_grouped.set_index('NameCustomer')['feeTeam'].to_dict()
        return dict_for_feeTeam

    def add_col_paid_payment_for_consulation(self, df):
        dict_of_feeTeam = self.get_dict_for_fee_team(df)
        self.df[u'תשלום ששולם עד היום לייעוץ'] = np.nan
        for index in range(0, len(self.df)):
            if self.df.iloc[index][u'שם לקוח'] in dict_of_feeTeam.keys():
                self.df.set_value(index, u'תשלום ששולם עד היום לייעוץ',
                                  math.ceil(dict_of_feeTeam.get(self.df.iloc[index][u'שם לקוח'])))

    def get_dict_for_fee_tax(self, df):
        df_grouped = pd.DataFrame(df.groupby('NameCustomer').feeTax.sum()).reset_index()
        dict_for_feeTax = df_grouped.set_index('NameCustomer')['feeTax'].to_dict()
        return dict_for_feeTax

    def add_col_paid_payment_for_gvia(self, df):
        dict_of_feeTax = self.get_dict_for_fee_tax(df)
        self.df[u'תשלום ששולם עד היום לגביה'] = np.nan
        for index in range(0, len(self.df)):
            if self.df.iloc[index][u'שם לקוח'] in dict_of_feeTax.keys():
                self.df.set_value(index, u'תשלום ששולם עד היום לגביה',
                                  math.ceil(dict_of_feeTax.get(self.df.iloc[index][u'שם לקוח'])))

    def add_col_execution_date(self):
        dates = []
        query = '''SELECT NameCustomer, DateFU
                FROM dbo.tblCustomers LEFT JOIN dbo.tbltaxes
                  ON dbo.tbltaxes.KodCustomer = dbo.tblCustomers.KodCustomer'''
        q = self.manager.db_service.search(query=query)
        df = pd.DataFrame.from_records(q)
        for index, row in df.iterrows():
            if row['DateFU'] is not None:
                date = datetime(row['DateFU'][0], row['DateFU'][1], row['DateFU'][2],
                                row['DateFU'][3], row['DateFU'][4], row['DateFU'][5])
                dates.append(date)
            else:
                date = np.nan
                dates.append(date)
        df['date for execution'] = dates
        dict_for_dates = df.set_index('NameCustomer')['date for execution'].to_dict()
        self.df[u'תאריך לביצוע'] = np.nan
        for index in range(0, len(self.df)):
            if self.df.iloc[index][u'שם לקוח'] in dict_for_dates.keys():
                self.df.set_value(index, u'תאריך לביצוע',
                                  dict_for_dates.get(self.df.iloc[index][u'שם לקוח']))

    def add_col_for_ans_for_hani(self):
        self.df[u'תשובות לחני'] = np.nan


    def add_col_notes_from_gvia_sheet(self):
        query = '''SELECT NameCustomer, ISNULL(Text, '') AS Text
                      FROM dbo.tbltaxes LEFT JOIN dbo.tblCustomers
                        ON dbo.tblCustomers.KodCustomer = dbo.tbltaxes.KodCustomer
                          WHERE dbo.tblCustomers.CustomerStatus = 2'''
        query_to_db_manager = self.manager.db_service.search(query=query)
        df = pd.DataFrame.from_records(query_to_db_manager)
        rows = len(df.index)
        for row in range(0, rows):
            last_3_notes = u""
            notes = df['Text'][row]
            patern = re.compile(r'\d{2}[/-]\d{2}[/-]\d{2,4}')
            list_of_notes = patern.split(notes)
            list_of_dates = re.findall(r'\d{2}[/-]\d{2}[/-]\d{2,4}', notes)
            list_of_notes = filter(lambda x: x != u'', list_of_notes[-3:])
            list_of_dates = filter(lambda x: x != u'', list_of_dates[-3:])
            for note, date in zip(list_of_notes, list_of_dates):
                last_3_notes = last_3_notes + date + note
            for row1 in range(0, len(self.df)):
                if self.df[u'שם לקוח'][row1] == df['NameCustomer'][row]:
                    self.df[u'הערות מגיליון הגביה'][row1] = last_3_notes


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
                           u'צפי לחודש', u'תשלום ששולם עד היום לייעוץ', u'צפי שנותר',
                           u'תשלום ששולם עד היום לגביה', u'הערות מגיליון הגביה', u'הערות', u'תאריך לביצוע',
                           u'תשובות לחני']]



    # def change_type_of_col_to_int(self, col_name):
    #     self.df[col_name] = self.df[col_name].apply(lambda x: int(x))
    #
    #
    # def change_type_of_sums_col_to_int(self):
    #     self.change_type_of_col_to_int(u'תשלום ייעוץ חודשי')
    #     self.change_type_of_col_to_int(u'צפי לחודש')
    #     self.change_type_of_col_to_int(u'תשלום ששולם עד היום לייעוץ')
    #     self.change_type_of_col_to_int(u'צפי שנותר')
    #     self.change_type_of_col_to_int(u'תשלום ששולם עד היום לגביה')

    def add_mid_sums(self, df):
        rows = len(df.index)
        team = df.iloc[0][u'צוות']
        first_index = 0
        last_index = 0
        for r in range(0, rows):
            if r == rows - 1:
                last_index = len(df.index) - 1
                df = self.calc_mid_sums(df, first_index, last_index, team)
            elif df[u'צוות'][r + 1] == team:
                last_index += 1
            else:
                if df[u'צוות'][r + 1] != team:
                    new_team = df[u'צוות'][r + 1]
                    df = self.calc_mid_sums(df, first_index, last_index, team)
                    last_index += 1
                    first_index = last_index + 1
                    team = new_team
        df = df.sort_index()
        return df

    def calc_mid_sums(self, df, first_index, last_index, team):
        index1 = first_index + 2
        index2 = last_index + 2
        sum_of_month_consultation = '=SUM(E{first_index}:E{last_index})'.format(first_index=index1, last_index=index2)
        sum_of_month_zefi = '=SUM(F{first_index}:F{last_index})'.format(first_index=index1, last_index=index2)
        sum_of_paid_payment_for_consulation = '=SUM(G{first_index}:G{last_index})'.format(first_index=index1,
                                                                                          last_index=index2)
        sum_of_zefi_left = '=SUM(H{first_index}:H{last_index})'.format(first_index=index1, last_index=index2)
        sum_of_paid_payment_for_gvia = '=SUM(I{first_index}:I{last_index})'.format(first_index=index1,
                                                                                   last_index=index2)
        new_sum_line = pd.DataFrame({u'צוות': [team], u'שם לקוח': [u'סיכום לצוות {team}'.format(team=team)],
                                     u'סוג לקוח': [""],  u'לקוח משפטי': [""],
                                     u'תשלום ייעוץ חודשי': [sum_of_month_consultation],
                                     u'צפי לחודש': [sum_of_month_zefi],
                                     u'תשלום ששולם עד היום לייעוץ': [sum_of_paid_payment_for_consulation],
                                     u'צפי שנותר': [sum_of_zefi_left],
                                     u'תשלום ששולם עד היום לגביה': [sum_of_paid_payment_for_gvia],
                                     u'הערות מגיליון הגביה': [""], u'הערות': [""],
                                     u'תאריך לביצוע': [""]})
        new_df = concat([df.ix[:last_index], new_sum_line, df.ix[last_index + 1:]]).reset_index(drop=True)
        new_df.sort_index()
        self.rows_of_sum.append(last_index + 1)
        return new_df

    def apply_style_on_sum_rows(self, sf):
        for row in self.rows_of_sum_positive:
            sf.apply_style_by_indexes(indexes_to_style=[sf.index[row]], bg_color=colors.yellow, bold=True)
        for row in self.rows_of_sum:
            sf.apply_style_by_indexes(indexes_to_style=[sf.index[row]], bg_color=colors.yellow, bold=True)


    def add_total_row(self, df):
        sum_of_month_consultation = '=E{r1}+E{r2}+E{r3}+E{r4}+E{r5}+E{r6}+E{r7}+E{r8}+E{r9}'.format(
            r1=self.rows_of_sum[0] + 2, r2=self.rows_of_sum[1] + 2, r3=self.rows_of_sum[2] + 2,
            r4=self.rows_of_sum[3] + 2, r5=self.rows_of_sum[4] + 2, r6=self.rows_of_sum[5] + 2,
            r7=self.rows_of_sum[6] + 2, r8=self.rows_of_sum[7] + 2, r9=self.rows_of_sum[8] + 2)
        sum_of_month_zefi = '=F{r1}+F{r2}+F{r3}+F{r4}+F{r5}+F{r6}+F{r7}+F{r8}+F{r9}'.format(
            r1=self.rows_of_sum[0] + 2, r2=self.rows_of_sum[1] + 2, r3=self.rows_of_sum[2] + 2,
            r4=self.rows_of_sum[3] + 2, r5=self.rows_of_sum[4] + 2, r6=self.rows_of_sum[5] + 2,
            r7=self.rows_of_sum[6] + 2, r8=self.rows_of_sum[7] + 2, r9=self.rows_of_sum[8] + 2)
        sum_of_paid_payment_for_consultation = '=G{r1}+G{r2}+G{r3}+G{r4}+G{r5}+G{r6}+G{r7}+G{r8}+G{r9}'.format(
            r1=self.rows_of_sum[0] + 2, r2=self.rows_of_sum[1] + 2, r3=self.rows_of_sum[2] + 2,
            r4=self.rows_of_sum[3] + 2, r5=self.rows_of_sum[4] + 2, r6=self.rows_of_sum[5] + 2,
            r7=self.rows_of_sum[6] + 2, r8=self.rows_of_sum[7] + 2, r9=self.rows_of_sum[8] + 2)
        sum_of_zefi_left = '=H{r1}+H{r2}+H{r3}+H{r4}+H{r5}+H{r6}+H{r7}+H{r8}+H{r9}'.format(
            r1=self.rows_of_sum[0] + 2, r2=self.rows_of_sum[1] + 2, r3=self.rows_of_sum[2] + 2,
            r4=self.rows_of_sum[3] + 2, r5=self.rows_of_sum[4] + 2, r6=self.rows_of_sum[5] + 2,
            r7=self.rows_of_sum[6] + 2, r8=self.rows_of_sum[7] + 2, r9=self.rows_of_sum[8] + 2)
        sum_of_paid_payment_for_gvia = '=I{r1}+I{r2}+I{r3}+I{r4}+I{r5}+I{r6}+I{r7}+I{r8}+I{r9}'.format(
            r1=self.rows_of_sum[0] + 2, r2=self.rows_of_sum[1] + 2, r3=self.rows_of_sum[2] + 2,
            r4=self.rows_of_sum[3] + 2, r5=self.rows_of_sum[4] + 2, r6=self.rows_of_sum[5] + 2,
            r7=self.rows_of_sum[6] + 2, r8=self.rows_of_sum[7] + 2, r9=self.rows_of_sum[8] + 2)
        new_sum_line = pd.DataFrame({u'צוות': [""], u'שם לקוח': [u'סיכום לכל הצוותים:'],
                                     u'סוג לקוח': [""],  u'לקוח משפטי': [""],
                                     u'תשלום ייעוץ חודשי': [sum_of_month_consultation],
                                     u'צפי לחודש': [sum_of_month_zefi],
                                     u'תשלום ששולם עד היום לייעוץ': [sum_of_paid_payment_for_consultation],
                                     u'צפי שנותר': [sum_of_zefi_left],
                                     u'תשלום ששולם עד היום לגביה': [sum_of_paid_payment_for_gvia],
                                     u'הערות מגיליון הגביה': [""], u'הערות': [""],
                                     u'תאריך לביצוע': [""]})
        new_df = concat([df, new_sum_line]).reset_index(drop=True)
        return new_df


    def separate_positive_and_negative (self):
        cond_for_negative = self.df[u'תשלום ששולם עד היום לייעוץ'] < 0
        positive_df = self.df[~cond_for_negative]
        positive_df = positive_df.reset_index(drop=True)
        positive_df = self.add_mid_sums(positive_df)
        positive_df = self.add_total_row(positive_df)
        self.rows_of_sum_positive = list(self.rows_of_sum)
        self.rows_of_sum_positive.append(len(positive_df) - 1)
        self.df = pd.DataFrame(positive_df)
        # self.rows_of_sum = []
        # negative_df = self.df[cond_for_negative]
        # negative_df = negative_df.reset_index(drop=True)
        # negative_df.to_excel("df negative.xlsx")
        # negative_df = self.add_mid_sums(negative_df)
        # negative_df = self.add_total_row(negative_df)
        # self.rows_of_sum_negative.append(len(negative_df) - 1)
        # self.df = pd.DataFrame(negative_df)
        # self.df = concat([positive_df, negative_df]).reset_index(drop=True)




    def add_row_for_test(self):
        self.df.loc[len(self.df)] = [u'שנהב', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13',
                                     '14', '15', '16', '17', '18', '19', '20', '21', -24523, '23', '24', '25']
        self.df.loc[len(self.df)] = [u'שנהב', '1', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13',
                                     '14', '15', '16', '17', '18', '19', '20', '21', -86, '23', '24', '25']


    def create_df(self):
        self.remove_sum_rows()
        self.add_col_tzefi_left(0)
        self.add_cols_for_fees()
        self.add_col_execution_date()
        self.add_col_for_ans_for_hani()
        # self.add_col_notes_from_gvia_sheet()     ///////////PUT BACK!!!!!!///////////
        self.add_col_month_tzefi()
        self.add_row_for_test()
        self.separate_positive_and_negative()
        self.add_col_tzefi_left(1)
        self.remove_needless_columns()


gvia_day_manager = GviaDailyReport()
gvia_day_manager.create_df()

sf = StyleFrame(gvia_day_manager.df)
gvia_day_manager.apply_style_on_sum_rows(sf)
writer = StyleFrame.ExcelWriter('Gvia day report.xlsx')
sf.to_excel(excel_writer=writer, right_to_left=True, row_to_add_filters=0, columns_and_rows_to_freeze='A2')
writer.save()
