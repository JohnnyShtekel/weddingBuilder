# -*- coding: utf-8 -*-

from esg_services import  EsgServiceManager

class DBHandler(object):
    def __init__(self):
        self.manager = EsgServiceManager()

    def inset_comment_to_db(self,total_comment,customerName):
       total = total_comment.decode('utf-8')
       customer = customerName.encode('utf-8')
       query = """
            UPDATE taxs
            SET Text = '{total}'
            From tbltaxes taxs
            INNER JOIN tblCustomers
                ON tblCustomers.kodCustomer = taxs.kodCustomer
            WHERE NameCustomer LIKE '%{customer}%'
       """.format(total=total.encode('utf-8'),customer=customer)
       self.manager.db_service.edit(query=query)


    def get_old_comment(self,customerName):
       customer = customerName.encode('utf-8')
       query = "SELECT a.Text FROM tbltaxes as a ,tblCustomers as b WHERE  b.kodCustomer = a.kodCustomer AND b.NameCustomer  LIKE '%{}%'".format(customer)
       return self.manager.db_service.search(query=query)[0]['Text']

