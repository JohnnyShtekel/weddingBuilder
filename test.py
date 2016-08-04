#
#
#
#
# # coding=utf-8
import pandas as pd
# import xlsxwriter
# from xlrd import open_workbook
# from xlutils.copy import copy
# from xlsxwriter.workbook import Workbook
# from esg_services import EsgServiceManager
# from esg_dal import EsgDal()
#
# dal = EsgDal(tables_list=['tblCustomers','tbltaxesPay'])
#
# import pandas as pd
# manager = EsgServiceManager()
# df = pd.DataFrame.from_records(manager.db_service.search(query='''SELECT NameCustomer
# FROM tbltaxesPay  as a LEFT JOIN tblCustomers as b
#   ON a.KodCustomer = b.KodCustomer'''))
# df.to_excel('output.xlsx')
# tblCustomers = dal()
# dal.query()
#
# # df = pd.DataFrame({'a': [1, 2, 3], 'b': [4, 5, 6], 'c': [7, 8, 9], 'd': ["=B2+C2+D2", 11, 12]})
# # rb = open_workbook("gvia_month.xlsx")
# # wb = copy(rb)
# # locked = wb.add_format({'locked': False})
#
# # workbook = Workbook('protection1.xlsx')
# # worksheet = workbook.add_worksheet()
# #
# # # Create a cell format with protection properties.
# # unlocked = workbook.add_format({'locked': False})
# # worksheet.set_column('A:A', 40)
# # worksheet.protect()
# # worksheet.write('A1', 'Cell B1 is locked. It cannot be edited.')
# # worksheet.write('A2', 'Cell B2 is unlocked. It can be edited.')
# # worksheet.write_formula('B1', '=3+2')  # Locked by default.
# # worksheet.write_formula('B2', '=1+2', unlocked)
# # workbook.close()
#
# # wb.save('names.xls')
#
# # df.to_excel('test_excel.xlsx')
# # rb = open_workbook("test_excel.xlsx")
# # wb = copy(rb)
# # ws = wb.sheet_by_index(0)
# # # worksheet = wb.get_worksheet_by_name('Sheet1')
# # ws.protect = True
# #
# # locked = wb.add_format({'locked': True})
# # # s = wb.get_sheet(0)
# # wb.save('names1.xls')
# #
# # # line = pd.DataFrame({'a': [13], 'b': [14], 'c': [15], 'd': [16]})
# #
# #
# # # writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')
# # # locked = writer.add_format({'locked': 1})
# # # # writer.write_cells(sheet_name=None,startrow=2,startcol=11, locked)
# # # writer.protect()
# # #
# # # df.to_excel(writer, sheet_name='Sheet1')
# # # writer.save()
# #
# #
# # # workbook = xlsxwriter.Workbook('protection.xlsx')
# # # worksheet = workbook.add_worksheet()
# # #
# # # # Create some cell formats with protection properties.
# # # locked = workbook.add_format({'locked': 1})
# # # # hidden = workbook.add_format({'hiddendden': 1})
# # #
# # # # Format the columns to make the text more visible.
# # # # worksheet.set_column('A:A', 40)
# # #
# # # # Turn worksheet protection on.
# # # worksheet.protect()
# # #
# # # # Write a locked, unlocked and hidden cell.
# # # worksheet.write('A1', 'Cell B1 is locked. It cannot be edited.')
# # # worksheet.write('A2', 'Cell B2 is unlocked. It can be edited.')
# # # worksheet.write('A3', "Cell B3 is hidden. The formula isn't visible.")
# # #
# # # worksheet.write_formula('B1', '=1+2')  # Locked by default.
# # # worksheet.write_formula('B2', '=1+2', locked)
# # #
# # # workbook.close()

# from datetime import datetime
#
# date = datetime(2016, 12, 24)
# # date.year = 2016
# # date.month = 10
# # date.day = 12
# print date

#
# df = pd.DataFrame({'a': [1,1,1,1,2,2,2], 'b': [3,4,5,6,7,8,9], 'c': [8,9,10,11,12,13,4]})
# print df
#
# print "**********************************************************"
#
# tmp = df.groupby(['a'])['c'].transform(max) == df['c']
# df1 = df[tmp]
# print df1

# df1 = df.groupby(['a']).agg({'b': lambda x: x.iloc[-1]})
# df2 = df.groupby(['a']).agg({'c': lambda x: x.iloc[-1]})
# df1['a'] = df1.index
# df2['a'] = df2.index
# df1 = df1.reindex_axis(sorted(df1.columns), axis=1)
# df2 = df2.reindex_axis(sorted(df2.columns), axis=1)
#
# print df1
# print df2
# print df1.merge(key='a', how='inner')



# *************************************************************************
# df = df.sort(['NameCustomer', 'report_date'], axis=0, ascending=[1, 0])
# df = df.groupby(['NameCustomer'], sort=False)['report_date'].max()
# ******************************************************************
import pandas as pd
from esg_services import EsgServiceManager

class GviaMonthReport(object):
    def __init__(self, query):
        self.manager = EsgServiceManager()
        self.df = pd.DataFrame.from_records(self.manager.db_service.search(query=query))
        self.rows_of_sum = []


q = "SELECT * FROM dbo.tblMonthlyTarget"

gvia_manager = GviaMonthReport(q)

gvia_manager.df.to_excel("tblMonthlyTarget.xlsx")
