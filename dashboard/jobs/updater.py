from datetime import datetime
from apscheduler.schedulers.background import BackgroundScheduler
from .jobs import schedule

def start():
    scheduler = BackgroundScheduler()
    scheduler.add_job(schedule, 'cron', month='*', day='1', hour=0, minute=0, second=0)
    scheduler.start()