from gvia_daily_report import GviaDaily
import datetime
if __name__ == '__main__':
    current_date = datetime.datetime.now()
    gvia_daily_report_handler = GviaDaily('report_for_hani', current_date, True)