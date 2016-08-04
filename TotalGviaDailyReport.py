# coding=utf-8
import pandas as pd
from esg_services import EsgServiceManager
from StyleFrame import StyleFrame, colors
from pandas import concat
import re



class TotalGviaDailyReport(object):
    def __init__(self):
        self.manager = EsgServiceManager()
        self.df = pd.DataFrame()
        self.rows_of_sum = []


    def add_col_team(self):
        list_of_teams = [u'ברקת', u'אלמוג', u'אודם', u'פנינה', u'קריסטל', u'שוהם-שכר', u'שנהב', u'ספיר', u'טורקיז',
                u'סה"כ יישום', u'מחלקת גביה']
        self.df[u'צוות'] = list_of_teams

    def get_paid_payment_for_consulation_from_excel_daily_report(self):
        df = pd.read_excel('Gvia day report.xlsx', convert_float=True)
        df = df[[u'צוות', u'שם לקוח', u'סוג לקוח', u'תשלום ששולם עד היום לייעוץ - חלוקת הגביה']]
        list_of_sum_rows = []
        for row in range(0, len(df)):
            if df[u'שם לקוח'][row].startswith(u'סיכום'):
                list_of_sum_rows.append(row)
        df.drop(df.index[list_of_sum_rows], inplace=True)
        df.index = range(len(df))
        df.to_excel("cuted daily.xlsx")
        return df

    def add_gvia_cols_for_3_kinds_of_customers(self):
        df = self.get_paid_payment_for_consulation_from_excel_daily_report()
        list_for_regular_customer_col = []
        list_for_consoltation_customer_col = []
        list_for_project_customer_col = []
        team = df[u'צוות'][0]
        regular_sum = 0
        consultation_sum = 0
        project_sum = 0
        for row in range(0, len(df)):
            if df[u'צוות'][row] == team:
                if df[u'סוג לקוח'][row] == u'ES - בסיס הצלחה':
                    regular_sum += df[u'תשלום ששולם עד היום לייעוץ - חלוקת הגביה']
                elif df[u'סוג לקוח'][row] == u'ES-יעוץ':
                    consultation_sum += df[u'תשלום ששולם עד היום לייעוץ - חלוקת הגביה']
                elif df[u'סוג לקוח'][row] == u'פרוייקטלי':
                    project_sum += df[u'תשלום ששולם עד היום לייעוץ - חלוקת הגביה']
            if df[u'צוות'][row + 1] != team:
                list_for_regular_customer_col.append(regular_sum)
                list_for_consoltation_customer_col.append(consultation_sum)
                list_for_project_customer_col.append(project_sum)
                team = df[u'צוות'][row + 1]
        df[u'גביה מרגיל']






gvia_teams = TotalGviaDailyReport()
gvia_teams.add_col_team()
gvia_teams.get_paid_payment_for_consulation_from_excel_daily_report()
gvia_teams.add_gvia_cols_for_3_kinds_of_customers()
print "***********************************************************************************"
print gvia_teams.df


#     self.df = pd.DataFrame.from_records(self.manager.db_service.search(query=query))
#
#     def add_col_tzefi_for_month(self, status):
#         tzefi_for_month = []
#         rows = len(self.df.index)
#         if status == 1:
#             for r in range(2, rows + 2):
#                 if r-2 not in self.rows_of_sum:
#                     tzefi_for_month.append('=IF(E{r}={yes},0,IF(I{r}>3,0,F{r}))'.format(r=r, yes='"כן"'))
#                 else:
#                     tzefi_for_month.append(self.df.iloc[r - 2][u'צפי לחודש'])
#         else:
#             for r in range(2, rows + 2):
#                 tzefi_for_month.append(0)
#         self.df[u'צפי לחודש'] = tzefi_for_month
#
#
#     def add_col_additional_payments_for_month(self, status):
#         additional_payments = []
#         rows = len(self.df.index)
#         if status == 1:
#             for r in range(2, rows + 2):
#                 if r-2 not in self.rows_of_sum:
#                     additional_payments.append("=IF(E{r}={yes}, 0, SUM(G{r}+H{r}+J{r}))".format(r=r, yes='"כן"'))
#                 else:
#                     additional_payments.append(self.df.iloc[r - 2][u'תשלומים נוספים בחודש'])
#         else:
#             for r in range(2, rows + 2):
#                 additional_payments.append(0)
#         self.df[u'תשלומים נוספים בחודש'] = additional_payments
#
#     def add_col_tzefi_meuhad_for_month(self, status):
#         tzefi_meuhad = []
#         rows = len(self.df.index)
#         if status == 1:
#             for r in range(2, rows + 2):
#                 if r-2 not in self.rows_of_sum:
#                     tzefi_meuhad.append('=SUM(M{r}+L{r})'.format(r=r))
#                 else:
#                     tzefi_meuhad.append(self.df.iloc[r - 2][u'צפי מאוחד'])
#         else:
#             for r in range(2, rows + 2):
#                 tzefi_meuhad.append(0)
#         self.df[u'צפי מאוחד'] = tzefi_meuhad
#
#
#     def set_3_last_notes(self):
#         rows = len(self.df.index)
#         for row in range(0, rows):
#             last_3_notes = u""
#             notes = self.df[u'הערות מגיליון הגביה'][row]
#             patern = re.compile(r'\d{2}[/-]\d{2}[/-]\d{2,4}')
#             list_of_notes = patern.split(notes)
#             list_of_dates = re.findall(r'\d{2}[/-]\d{2}[/-]\d{2,4}', notes)
#             list_of_notes = filter(lambda x : x!= u'',list_of_notes[-3:])
#             list_of_dates = filter(lambda x : x!= u'',list_of_dates[-3:])
#             for note, date in zip(list_of_notes, list_of_dates):
#                 last_3_notes = last_3_notes + date + note
#                 # last_3_notes = u'{a}{b}{c}'.format(a=last_3_notes, b=date, c=note)
#             self.df[u'הערות מגיליון הגביה'][row] = last_3_notes
#
#
#     def arrange_col_order(self):
#         self.df = self.df[[u'צוות', u'שם לקוח', u'אמצעי תשלום', u'סוג לקוח', u'לקוח משפטי', u'תשלום ייעוץ חודשי',
#                            u'תשלום על בונוס', u'תשלומים מיוחדים', u'מספר התשלומים מעבר לאשראי',
#                            u'סכום התשלומים מעבר לאשראי', u'הערות', u'צפי לחודש', u'תשלומים נוספים בחודש', u'צפי מאוחד',
#                            u'הערות מגיליון הגביה']]
#
#     def change_num_format(self):
#         columns_to_add_comma = [u'תשלום ייעוץ חודשי', u'תשלום על בונוס', u'תשלומים מיוחדים', u'מספר התשלומים מעבר לאשראי',
#                                 u'סכום התשלומים מעבר לאשראי']
#         for column in columns_to_add_comma:
#             self.df[column] = self.df[column].apply(lambda x: format(int((round(x))), ',d'))
#
#     def change_type_of_col_to_int(self, col_name):
#         self.df[col_name] = self.df[col_name].apply(lambda x: int(x))
#
#     def change_type_of_sums_col_to_int(self):
#         self.change_type_of_col_to_int(u'תשלום ייעוץ חודשי')
#         self.change_type_of_col_to_int(u'תשלום על בונוס')
#         self.change_type_of_col_to_int(u'תשלומים מיוחדים')
#         self.change_type_of_col_to_int(u'סכום התשלומים מעבר לאשראי')
#
#
#
#     def add_mid_sums(self):
#         rows = len(self.df.index)
#         team = self.df[u'צוות'][0]
#         first_index = 0
#         last_index = 0
#         for r in range(0, rows):
#             if r == rows - 1:
#                 last_index = len(self.df.index) - 1
#                 self.calc_mid_sums(first_index, last_index, team)
#             if self.df[u'צוות'][r + 1] == team:
#                 last_index = last_index + 1
#
#             else:
#                 if self.df[u'צוות'][r + 1] != team:
#                     new_team = self.df[u'צוות'][r + 1]
#                     self.calc_mid_sums(first_index, last_index, team)
#                     last_index = last_index + 1
#                     first_index = last_index + 1
#                     team = new_team
#         self.df.sort_index()
#
#     def calc_mid_sums(self, first_index, last_index, team):
#         index1 = first_index + 2
#         index2 = last_index + 2
#         sum_of_month_consultation = '=SUM(F{first_index}:F{last_index})'.format(first_index=index1, last_index=index2)
#         sum_of_pay_for_bonus = '=SUM(G{first_index}:G{last_index})'.format(first_index=index1, last_index=index2)
#         sum_of_special_payments = '=SUM(H{first_index}:H{last_index})'.format(first_index=index1, last_index=index2)
#         sum_of_month_over_credit = '=SUM(J{first_index}:J{last_index})'.format(first_index=index1, last_index=index2)
#         sum_of_month_zefi = '=SUM(L{first_index}:L{last_index})'.format(first_index=index1, last_index=index2)
#         sum_of_extra_payments_for_month = '=SUM(M{first_index}:M{last_index})'.format(first_index=index1, last_index=index2)
#         sum_of_zefi_meuhad = '=SUM(N{first_index}:N{last_index})'.format(first_index=index1, last_index=index2)
#         new_sum_line = pd.DataFrame({u'צוות': [team], u'שם לקוח': [u'סיכום לצוות {team}'.format(team=team)],
#                                      u'אמצעי תשלום': [""], u'סוג לקוח': [""], u'לקוח משפטי': [""],
#                                      u'תשלום ייעוץ חודשי': [sum_of_month_consultation],
#                                      u'תשלום על בונוס': [sum_of_pay_for_bonus],
#                                      u'תשלומים מיוחדים': [sum_of_special_payments],
#                                      u'מספר התשלומים מעבר לאשראי': [""],
#                                      u'סכום התשלומים מעבר לאשראי': [sum_of_month_over_credit],
#                                      u'הערות': [""], u'צפי לחודש': [sum_of_month_zefi],
#                                      u'תשלומים נוספים בחודש': [sum_of_extra_payments_for_month],
#                                      u'צפי מאוחד': [sum_of_zefi_meuhad], u'הערות מגיליון הגביה': [""]})
#         new_df = concat([self.df.ix[:last_index], new_sum_line, self.df.ix[last_index + 1:]]).reset_index(drop=True)
#         new_df.sort_index()
#         self.df = new_df
#         self.rows_of_sum.append(last_index + 1)
#
#     def apply_style_on_sum_rows(self, sf):
#         for row in self.rows_of_sum:
#             sf.apply_style_by_indexes(indexes_to_style=[sf.index[row]], bg_color=colors.yellow, bold=True)
#
#     def add_total_row(self):
#         sum_of_month_consultation = '=F{r1}+F{r2}+F{r3}+F{r4}+F{r5}+F{r6}+F{r7}+F{r8}+F{r9}'.format(
#             r1=self.rows_of_sum[0]+2, r2=self.rows_of_sum[1]+2, r3=self.rows_of_sum[2]+2, r4=self.rows_of_sum[3]+2,
#             r5=self.rows_of_sum[4]+2, r6=self.rows_of_sum[5]+2, r7=self.rows_of_sum[6]+2, r8=self.rows_of_sum[7]+2,
#             r9=self.rows_of_sum[8]+2)
#         sum_of_pay_for_bonus = '=G{r1}+G{r2}+G{r3}+G{r4}+G{r5}+G{r6}+G{r7}+G{r8}+G{r9}'.format(
#             r1=self.rows_of_sum[0]+2, r2=self.rows_of_sum[1]+2, r3=self.rows_of_sum[2]+2, r4=self.rows_of_sum[3]+2,
#             r5=self.rows_of_sum[4]+2, r6=self.rows_of_sum[5]+2, r7=self.rows_of_sum[6]+2, r8=self.rows_of_sum[7]+2,
#             r9=self.rows_of_sum[8]+2)
#         sum_of_special_payments = '=H{r1}+H{r2}+H{r3}+H{r4}+H{r5}+H{r6}+H{r7}+H{r8}+H{r9}'.format(
#             r1=self.rows_of_sum[0]+2, r2=self.rows_of_sum[1]+2, r3=self.rows_of_sum[2]+2, r4=self.rows_of_sum[3]+2,
#             r5=self.rows_of_sum[4]+2, r6=self.rows_of_sum[5]+2, r7=self.rows_of_sum[6]+2, r8=self.rows_of_sum[7]+2,
#             r9=self.rows_of_sum[8]+2)
#         sum_of_month_over_credit = '=J{r1}+J{r2}+J{r3}+J{r4}+J{r5}+J{r6}+J{r7}+J{r8}+J{r9}'.format(
#             r1=self.rows_of_sum[0]+2, r2=self.rows_of_sum[1]+2, r3=self.rows_of_sum[2]+2, r4=self.rows_of_sum[3]+2,
#             r5=self.rows_of_sum[4]+2, r6=self.rows_of_sum[5]+2, r7=self.rows_of_sum[6]+2, r8=self.rows_of_sum[7]+2,
#             r9=self.rows_of_sum[8]+2)
#         sum_of_month_zefi = '=L{r1}+L{r2}+L{r3}+L{r4}+L{r5}+L{r6}+L{r7}+L{r8}+L{r9}'.format(
#             r1=self.rows_of_sum[0]+2, r2=self.rows_of_sum[1]+2, r3=self.rows_of_sum[2]+2, r4=self.rows_of_sum[3]+2,
#             r5=self.rows_of_sum[4]+2, r6=self.rows_of_sum[5]+2, r7=self.rows_of_sum[6]+2, r8=self.rows_of_sum[7]+2,
#             r9=self.rows_of_sum[8]+2)
#         sum_of_extra_payments_for_month = '=M{r1}+M{r2}+M{r3}+M{r4}+M{r5}+M{r6}+M{r7}+M{r8}+M{r9}'.format(
#             r1=self.rows_of_sum[0]+2, r2=self.rows_of_sum[1]+2, r3=self.rows_of_sum[2]+2, r4=self.rows_of_sum[3]+2,
#             r5=self.rows_of_sum[4]+2, r6=self.rows_of_sum[5]+2, r7=self.rows_of_sum[6]+2, r8=self.rows_of_sum[7]+2,
#             r9=self.rows_of_sum[8]+2)
#         sum_of_zefi_meuhad = '=N{r1}+N{r2}+N{r3}+N{r4}+N{r5}+N{r6}+N{r7}+N{r8}+N{r9}'.format(
#             r1=self.rows_of_sum[0]+2, r2=self.rows_of_sum[1]+2, r3=self.rows_of_sum[2]+2, r4=self.rows_of_sum[3]+2,
#             r5=self.rows_of_sum[4]+2, r6=self.rows_of_sum[5]+2, r7=self.rows_of_sum[6]+2, r8=self.rows_of_sum[7]+2,
#             r9=self.rows_of_sum[8]+2)
#         new_sum_line = pd.DataFrame({u'צוות': [""], u'שם לקוח': [u'סיכום לכל הצוותים:'],
#                                      u'אמצעי תשלום': [""], u'סוג לקוח': [""], u'לקוח משפטי': [""],
#                                      u'תשלום ייעוץ חודשי': [sum_of_month_consultation],
#                                      u'תשלום על בונוס': [sum_of_pay_for_bonus],
#                                      u'תשלומים מיוחדים': [sum_of_special_payments],
#                                      u'מספר התשלומים מעבר לאשראי': [""],
#                                      u'סכום התשלומים מעבר לאשראי': [sum_of_month_over_credit],
#                                      u'הערות': [""], u'צפי לחודש': [sum_of_month_zefi],
#                                      u'תשלומים נוספים בחודש': [sum_of_extra_payments_for_month],
#                                      u'צפי מאוחד': [sum_of_zefi_meuhad], u'הערות מגיליון הגביה': [""]})
#         new_df = concat([self.df, new_sum_line]).reset_index(drop=True)
#         self.df = new_df
#
#
#
# query_for_gvia_month_report = '''SELECT
#   NameTeam AS [צוות],
#   NameCustomer AS [שם לקוח],
#   MiddlePay AS [אמצעי תשלום],
#   AgreementKind AS [סוג לקוח],
#   CASE WHEN proceFinished=0 THEN 'כן' ELSE '' END AS [לקוח משפטי],
#   CASE WHEN Conditions='דולר' THEN ISNULL(PayDollar, 0)*4 ELSE ISNULL(PayDollar, 0) END AS [תשלום ייעוץ חודשי],
#   ISNULL(SumBonus,0) as [תשלום על בונוס],
#   ISNULL(SumSpecial,0) as [תשלומים מיוחדים],
#   ISNULL(NumOfDebt,0) as [מספר התשלומים מעבר לאשראי],
#   ISNULL(Debt,0) as [סכום התשלומים מעבר לאשראי],
#   ISNULL(Remark,'') as [הערות],
#   ISNULL(Text, '') AS [הערות מגיליון הגביה]
# FROM
#   dbo.tblCustomers
#   LEFT JOIN
#   dbo.tblTeamList
#     ON dbo.tblCustomers.KodTeamCare = dbo.tblTeamList.KodTeamCare
#   LEFT JOIN
#   (SELECT * FROM
#     (
#       SELECT
#         IDagreement,
#         KodAgreementKind,
#         KodCustomer,
#         PayProcess,
#         PayDollar,
#         row_number() OVER(PARTITION BY KodCustomer ORDER BY DateNew DESC) AS orderAgreementByDateDesc
#       FROM
#         dbo.tblAgreementConditionAdvice
#     ) AS temp
#       WHERE temp.orderAgreementByDateDesc = 1) AS tblLastAgreements
#     ON dbo.tblCustomers.KodCustomer = tblLastAgreements.KodCustomer
#   LEFT JOIN tblJudicialProc
#     ON tblCustomers.KodCustomer = tblJudicialProc.kodCustomer
#   LEFT JOIN dbo.tblMiddlePay
#     ON tblLastAgreements.PayProcess = dbo.tblMiddlePay.KodMiddlePay
#   LEFT JOIN dbo.tblAgreementKind
#     ON dbo.tblAgreementKind.KodAgreementKind = tblLastAgreements.KodAgreementKind
#   LEFT JOIN (
#     SELECT DISTINCT NumR, Conditions,
#       SUM(CASE WHEN NumPay = 0 AND SumP-Pay>0 AND PayRemark LIKE '%בונוס%' THEN SumP-pay ELSE 0 END) OVER(PARTITION BY NumR) AS SumBonus,
#       SUM(CASE WHEN NumPay = 0 AND SumP-Pay>0  AND NOT PayRemark LIKE '%בונוס%' THEN SumP ELSE 0 END) OVER(PARTITION BY NumR) AS SumSpecial,
#       SUM(CASE WHEN NumPay > 0 AND SumP-Pay>0  AND DatePay IS NULL AND (GETDATE() > (CASE
#      WHEN tblAgreementConditionAdvice.shotef = 1 THEN
#        DateAdd(DAY, tblAgreementConditionAdvice.ashray+1,
#                DateAdd(DAY,-1, Cast(CONVERT(VARCHAR(10),DATEPART(MM, DateAdd(month,1,tblAgreementConditionAdvicePay.DateP)),103)
#                                     + '/' + '01/' + CONVERT(VARCHAR(10),DATEPART(YYYY, DateAdd(month,1,tblAgreementConditionAdvicePay.DateP)),103) AS DATETIME)))
#      ELSE
#        DateAdd(DAY, tblAgreementConditionAdvice.ashray+1,tblAgreementConditionAdvicePay.DateP)
#      END)) THEN 1 ELSE 0 END) OVER(PARTITION BY NumR) AS NumOfDebt,
#       SUM(CASE WHEN NumPay > 0 AND SumP-Pay>0  AND DatePay IS NULL AND (GETDATE() > (CASE
#      WHEN tblAgreementConditionAdvice.shotef = 1 THEN
#        DateAdd(DAY, tblAgreementConditionAdvice.ashray+1,
#                DateAdd(DAY,-1, Cast(CONVERT(VARCHAR(10),DATEPART(MM, DateAdd(month,1,tblAgreementConditionAdvicePay.DateP)),103)
#                                     + '/' + '01/' + CONVERT(VARCHAR(10),DATEPART(YYYY, DateAdd(month,1,tblAgreementConditionAdvicePay.DateP)),103) AS DATETIME)))
#      ELSE
#        DateAdd(DAY, tblAgreementConditionAdvice.ashray+1,tblAgreementConditionAdvicePay.DateP)
#      END) ) THEN SumP ELSE 0 END) OVER(PARTITION BY NumR) AS Debt,
#       STUFF((SELECT ', ' + tblRemarks.PayRemark AS [text()]FROM dbo.tblAgreementConditionAdvicePay tblRemarks WHERE tblRemarks.NumR = dbo.tblAgreementConditionAdvicePay.NumR AND tblRemarks.NumPay = 0 AND SumP-Pay>0 FOR XML PATH('')), 1, 1, '' ) AS Remark
#     FROM dbo.tblAgreementConditionAdvicePay
#     LEFT JOIN dbo.tblAgreementConditionAdvice
#       ON dbo.tblAgreementConditionAdvice.IDagreement = dbo.tblAgreementConditionAdvicePay.NumR
#
#
#      )AS tblSumAllPay
#     ON tblLastAgreements.IDagreement = tblSumAllPay.NumR
#   LEFT JOIN (
#     SELECT KodCustomer, Text
#     FROM tbltaxes
#     ) tbltaxes
#     ON
#       dbo.tblCustomers.KodCustomer = tbltaxes.KodCustomer
# WHERE CustomerStatus = 2
# ORDER BY NameTeam'''
#
# gvia_manager = GviaMonthReport(query_for_gvia_month_report)
# gvia_manager.add_col_tzefi_for_month(0)
# gvia_manager.add_col_additional_payments_for_month(0)
# gvia_manager.add_col_tzefi_meuhad_for_month(0)
# gvia_manager.set_3_last_notes()
# gvia_manager.change_type_of_sums_col_to_int()
# # gvia_manager.change_num_format()   ////מוסיף פסיק למספרים ארוכים. נכון לעכשיו דופק את התאים ששונו מטקסט למספר
# gvia_manager.change_type_of_col_to_int(u'מספר התשלומים מעבר לאשראי')
# gvia_manager.add_mid_sums()
# gvia_manager.rows_of_sum.append(len(gvia_manager.df))
# gvia_manager.add_total_row()
# gvia_manager.add_col_tzefi_for_month(1)
# gvia_manager.add_col_additional_payments_for_month(1)
# gvia_manager.add_col_tzefi_meuhad_for_month(1)
# gvia_manager.arrange_col_order()
# sf = StyleFrame(gvia_manager.df)
# gvia_manager.apply_style_on_sum_rows(sf)
# writer = StyleFrame.ExcelWriter('gvia_month.xlsx')
# sf.to_excel(excel_writer=writer, right_to_left=True, row_to_add_filters=0, columns_and_rows_to_freeze='A2')
# writer.save()
#
#
#








