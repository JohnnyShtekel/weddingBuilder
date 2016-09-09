# -*- coding: utf-8 -*-
import datetime
from esg_services import EsgServiceManager


class DBHandler(object):
    def __init__(self):
        self.manager = EsgServiceManager()

    def inset_comment_to_db(self, total_comment, customer_name, date):
        total = total_comment.decode('utf-8')
        customer = customer_name.encode('utf-8')
        date_fu = date.strftime("%Y-%m-%d 00:00:00")
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

        # query = """SELECT a.Text
        #           FROM tbltaxes as a ,tblCustomers as b
        #           WHERE  b.kodCustomer = a.kodCustomer AND b.NameCustomer  LIKE '%{}%'""".format(customer)
        return self.manager.db_service.search(query=query)[0]['Text']
