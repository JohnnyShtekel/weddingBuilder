# -*- coding: utf-8 -*-
import re
from esg_services import EsgServiceManager


class DBHandler(object):
    def __init__(self):
        self.manager = EsgServiceManager()

    def inset_comment_to_db(self, total_comment, customer_name, date):
        total = total_comment.decode('utf-8')
        customer = customer_name.encode('utf-8')
        date_fu = date
        query = """
            UPDATE tbltaxes
            SET Text = '{total}', DateFU = '{date_fu}'
            From tbltaxes
            INNER JOIN tblCustomers
                ON tblCustomers.kodCustomer = tbltaxes.kodCustomer
            WHERE tblCustomers.NameCustomer LIKE '%{customer}%'
        """.format(total=total.encode('utf-8'), date_fu=date_fu, customer=customer)
        response = self.manager.db_service.edit(query=query)

    def get_old_comment(self, customer_name):
        customer = customer_name.encode('utf-8')
        query = """select Text from tblTaxes, tblCustomers
                    where NameCustomer like '%{}%' and tblTaxes.KodCustomer=tblCustomers.KodCustomer""".format(customer)
        try:
            return self.manager.db_service.search(query=query)[0]['Text']
        except:
            print customer_name
            print self.manager.db_service.search(query=query)

# customer = "נאנאוניקס אימג''ינג בעמ"
# print result

# customer = 'אדירם'
# query = """select Text from tblTaxes, tblCustomers
#             where NameCustomer like '%{}%' and tblTaxes.KodCustomer=tblCustomers.KodCustomer""".format(customer)
# manager = EsgServiceManager()
# comments = manager.db_service.search(query=query)[0]['Text']
# pattern = re.compile(r'\d{2}[/-]\d{2}[/-]\d{2,4}\s+\-\s+(.*)]')
# list_of_notes = pattern.split(comments)
# list_of_dates = re.findall(r'\d{2}[/-]\d{2}[/-]\d{2,4}', comments)
# for note in list_of_notes:
#     print note