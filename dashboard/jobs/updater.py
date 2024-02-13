from datetime import datetime, timedelta
from apscheduler.schedulers.background import BackgroundScheduler, BlockingScheduler
from .jobs import schedule, holidays, generate_noce_timesheet, clear_activity_files
from ..models import (User)
from ..constants import constants
import logging

# Dictionary to store the last execution time of each task
last_execution_times = {}

def generate_ts_ocg(param):
    ocg_users = User.objects.filter(company=constants.OCG)
    generate_noce_timesheet(users=ocg_users, companyName=constants.COMPANY_CHOICES[constants.OCG][1])
    last_execution_times['generate_ts_ocg'] = datetime.now()

def generate_ts_nk(param):
    nk_users = User.objects.filter(company=constants.NK)
    generate_noce_timesheet(users=nk_users, companyName=constants.COMPANY_CHOICES[constants.NK][1])
    last_execution_times['generate_ts_nk'] = datetime.now()

def generate_ts_EHAF(param):
    EHAF_users = User.objects.filter(company=constants.EHAF)
    generate_noce_timesheet(users=EHAF_users, companyName=constants.COMPANY_CHOICES[constants.EHAF][1])
    last_execution_times['generate_ts_EHAF'] = datetime.now()

def generate_ts_ACE(param):
    ACE_users = User.objects.filter(company=constants.ACE)
    generate_noce_timesheet(users=ACE_users, companyName=constants.COMPANY_CHOICES[constants.ACE][1])
    last_execution_times['generate_ts_ACE'] = datetime.now()

def generate_all():
    generate_noce_timesheet()
    last_execution_times['generate_all'] = datetime.now()

def start():
    try:
        for job in [generate_ts_ocg, generate_ts_nk, generate_ts_EHAF, generate_ts_ACE, generate_all]:
            job_name = job.__name__
            if job_name not in last_execution_times:
                last_execution_times[job_name] = datetime.now() - timedelta(days=1)

        scheduler = BackgroundScheduler()
        scheduler.add_job(schedule, 'cron', month='*', day='1', hour=0, minute=0, second=0)
        
        scheduler.add_job(generate_all, 'cron', hour=21, minute=32, id='ts_all')
        scheduler.add_job(generate_ts_ocg, 'cron', hour=21, minute=32, args=('param_ocg',), id='ts_ocg')
        scheduler.add_job(generate_ts_nk, 'cron', hour=21, minute=32, args=('param_nk',), id='ts_nk')
        scheduler.add_job(generate_ts_EHAF, 'cron', hour=21, minute=32, args=('param_EHAF',), id='ts_EHAF')
        scheduler.add_job(generate_ts_ACE, 'cron', hour=21, minute=32, args=('param_ACE',), id='ts_ACE')

        scheduler.start()
    except Exception as e:
        logging.exception("An error occurred in the scheduler: %s", e)