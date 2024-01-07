from datetime import datetime, timedelta
from apscheduler.schedulers.background import BackgroundScheduler
from .jobs import schedule, holidays

def start():
    scheduler = BackgroundScheduler()
    scheduler.add_job(schedule, 'cron', month='*', day='1', hour=0, minute=0, second=0)
    # Calculate the last day of the current month
    last_day_of_current_month = datetime.now().replace(day=1, hour=23, minute=59, second=59) - timedelta(days=1)
    # Schedule the 'holidays' job to run at 11:59 pm on the last day of the current month
    scheduler.add_job(holidays, 'date', run_date=last_day_of_current_month)
    scheduler.start()