import os
import django
import datetime
import time
import sys
from pathlib import Path
from datetime import datetime

# Get the parent directory of the script
parent_dir = Path(__file__).resolve().parent.parent.parent

# Add the parent directory to sys.path
sys.path.append(parent_dir.__str__())

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'dashboardAPI.settings')
django.setup()

from dashboard.models import User
from dashboard.constants import constants
from dashboard.jobs.jobs import generate_noce_timesheet


last_day_checked_1 = 0
last_day_checked_2 = 0
last_day_checked_3 = 0
last_day_checked_4 = 0
last_day_checked_5 = 0


def generate_ts_ocg():
    ocg_users = User.objects.filter(company=constants.OCG)
    generate_noce_timesheet(users=ocg_users, companyName=constants.COMPANY_CHOICES[constants.OCG][1])

def generate_ts_nk():
    nk_users = User.objects.filter(company=constants.NK)
    generate_noce_timesheet(users=nk_users, companyName=constants.COMPANY_CHOICES[constants.NK][1])

def generate_ts_EHAF():
    EHAF_users = User.objects.filter(company=constants.EHAF)
    generate_noce_timesheet(users=EHAF_users, companyName=constants.COMPANY_CHOICES[constants.EHAF][1])

def generate_ts_ACE():
    ACE_users = User.objects.filter(company=constants.ACE)
    generate_noce_timesheet(users=ACE_users, companyName=constants.COMPANY_CHOICES[constants.ACE][1])

def generate_all():
    generate_noce_timesheet()

while True:
    time_now = datetime.now()
    if time_now.hour == 13 and time_now.minute > 0 and last_day_checked_1 != time_now.day:
        generate_ts_ocg()
        last_day_checked_1 = time_now.day
    if time_now.hour == 13 and time_now.minute > 0 and last_day_checked_2 != time_now.day:
        generate_ts_nk()
        last_day_checked_2 = time_now.day
    if time_now.hour == 13 and time_now.minute > 0 and last_day_checked_3 != time_now.day:
        generate_ts_EHAF()
        last_day_checked_3 = time_now.day
    if time_now.hour == 13 and time_now.minute > 0 and last_day_checked_4 != time_now.day:
        generate_ts_ACE()
        last_day_checked_4 = time_now.day
    if time_now.hour == 13 and time_now.minute > 0 and last_day_checked_5 != time_now.day:
        generate_all()
        last_day_checked_5 = time_now.day
    time.sleep(60)