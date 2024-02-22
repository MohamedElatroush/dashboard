from django.utils import timezone
from ..models import Activity, User, ActivityFile
from datetime import datetime
from ..constants import constants
from datetime import date
from ..utilities.utilities import create_activity_excel_report
from django.db.models import Q
import boto3
from django.conf import settings
from botocore.exceptions import NoCredentialsError


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

def generate_noce_timesheet(users=None, companyName=None, date=None):
    if not users:
        users = User.objects.all().exclude(isAdmin=True)

    activities = Activity.objects.filter(user__in=users,\
                                          activityDate__year=date.year, \
                                            activityDate__month=date.month)

    date_param = date.today().strftime('%Y-%m-%d')
    datetime.strptime(date_param, '%Y-%m-%d').date()
    current_date = date_param

    return create_activity_excel_report(users, activities, current_date, companyName, date)

def clear_activity_files():
    try:
        s3 = boto3.client(
            's3',
            aws_access_key_id=settings.AWS_ACCESS_KEY_ID,
            aws_secret_access_key=settings.AWS_SECRET_ACCESS_KEY,
        )
    except NoCredentialsError:
        # Handle the case where credentials are not found
        print("Unable to locate AWS credentials. Please ensure they are configured.")
        return

    activity_files = ActivityFile.objects.all()
    if activity_files:
        for file in activity_files:
            s3.delete_object(Bucket=str(settings.AWS_STORAGE_BUCKET_NAME), Key=str(file.file))
            file.delete()
    return