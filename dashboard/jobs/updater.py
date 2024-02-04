from datetime import datetime, timedelta
from apscheduler.schedulers.background import BackgroundScheduler, BlockingScheduler
from .jobs import schedule, holidays, generate_noce_timesheet
from ..models import (User)
from ..constants import constants

def generate_ts_ocg(param):
    ocg_users = User.objects.filter(company=constants.OCG)
    generate_noce_timesheet(users=ocg_users, companyName=constants.COMPANY_CHOICES[constants.OCG][1])

def generate_ts_nk(param):
    nk_users = User.objects.filter(company=constants.NK)
    generate_noce_timesheet(users=nk_users, companyName=constants.COMPANY_CHOICES[constants.NK][1])

def generate_ts_EHAF(param):
    EHAF_users = User.objects.filter(company=constants.EHAF)
    generate_noce_timesheet(users=EHAF_users, companyName=constants.COMPANY_CHOICES[constants.EHAF][1])

def generate_ts_ACE(param):
    ACE_users = User.objects.filter(company=constants.ACE)
    generate_noce_timesheet(users=ACE_users, companyName=constants.COMPANY_CHOICES[constants.ACE][1])

def start():
    print('start')
    scheduler = BackgroundScheduler()
    scheduler.add_job(schedule, 'cron', month='*', day='1', hour=0, minute=0, second=0)
    # Calculate the last day of the current month
    last_day_of_current_month = datetime.now().replace(day=1, hour=23, minute=59, second=59) - timedelta(days=1)
    # Schedule the 'holidays' job to run at 11:59 pm on the last day of the current month
    scheduler.add_job(holidays, 'date', run_date=last_day_of_current_month)

    scheduler.add_job(generate_noce_timesheet, 'cron', hour=13, minute=2)
    scheduler.add_job(generate_ts_ocg, 'cron', hour=12, minute=52, args=('param_ocg',), id='ts_ocg')
    scheduler.add_job(generate_ts_nk, 'cron', hour=12, minute=54, args=('param_nk',), id='ts_nk')
    scheduler.add_job(generate_ts_EHAF, 'cron', hour=13, minute=57, args=('param_EHAF',), id='ts_EHAF')
    scheduler.add_job(generate_ts_ACE, 'cron', hour=13, minute=59, args=('param_ACE',), id='ts_ACE')

    scheduler.start()