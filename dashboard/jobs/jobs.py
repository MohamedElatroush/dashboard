from django.utils import timezone
from ..models import Activity, User, ActivityFile
from datetime import datetime
from ..constants import constants
from datetime import date
from ..utilities.utilities import create_activity_excel_report
from django.db.models import Q


def schedule():
    today = timezone.now().date()
    eight_months_ago = today - timezone.timedelta(days=8 * 30)
     # Filter activities created more than 8 months ago
    old_activities = Activity.objects.filter(created__lt=eight_months_ago)
    # Delete the selected activities
    old_activities.delete()

def holidays():


    current_date = timezone.now()

    users = User.objects.all()

    # Get the first day and last day of the month
    last_day_of_current_month = current_date.replace(day=1) + timezone.timedelta(days=32)
    last_day = last_day_of_current_month.replace(day=1) - timezone.timedelta(days=1)

    # LOCAL USERS
    for user in users:
        current_day = current_date.replace(day=1)
        while current_day <= last_day:
            existing_activity = Activity.objects.filter(
                user=user,
                activityDate=current_day.date()
            ).first()

            if existing_activity:
                # If the user has an existing activity, do nothing
                pass
            else:
                # If no activity exists, create a new activity with type X (OFFDAY)
                Activity.objects.create(
                    user=user,
                    activityDate=current_day.date(),
                    activityType=constants.OFFDAY
                )

            # Move to the next day
            current_day += timezone.timedelta(days=1)


def generate_noce_timesheet(users=None, companyName=None):
    if not users:
        users = User.objects.all()
    activities = Activity.objects.filter(user__in=users)
    date_param = date.today().strftime('%Y-%m-%d')
    datetime.strptime(date_param, '%Y-%m-%d').date()
    current_date = date.today()
    create_activity_excel_report(users, activities, current_date, companyName)

def clear_activity_files():
    activity_files = ActivityFile.objects.all()
    print(activity_files)
    # Check if there are any objects
    if activity_files.exists():
    # If there are objects, delete them
        activity_files.delete()
    return