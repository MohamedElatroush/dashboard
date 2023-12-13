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
    users = User.objects.all()

    current_date = timezone.now()
    year = current_date.year
    month = current_date.month

    # Get the first day and last day of the month
    last_day_of_current_month = timezone.datetime(year, month, 1) + timezone.timedelta(days=32)
    last_day = timezone.datetime(last_day_of_current_month.year, last_day_of_current_month.month, 1) - timezone.timedelta(days=1)

    for user in users:
        current_day = timezone.datetime(year, month, 1)
        while current_day <= last_day:
            # Check if the current day is a Friday for local users or a Saturday/Sunday for experts
            if (
                (user.expert == constants.EXPERT_USER and current_day.weekday() in [5, 6]) or
                (user.expert == constants.LOCAL_USER and current_day.weekday() == 4)
            ):
                # Create the activity for the user
                activity_date = current_day.date()
                # Activity.objects.all().delete()
                existing_activity = Activity.objects.filter(
                    user=user,
                    activityDate=activity_date
                ).first()

                if existing_activity:
                    # Update existing activity type
                    existing_activity.activityType = constants.HOLIDAY
                    existing_activity.save()
                else:
                    # Create a new activity
                    Activity.objects.create(
                        user=user,
                        activityDate=activity_date,
                        activityType=constants.HOLIDAY
                    )

            # Move to the next day
            current_day += timezone.timedelta(days=1)
    # Bulk create the activities
    # Activity.objects.bulk_create(activities_to_create)