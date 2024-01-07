from django.utils import timezone
from ..models import Activity, User
from datetime import datetime
from ..constants import constants

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