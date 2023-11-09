from django.utils import timezone
from ..models import Activity
from datetime import datetime

def schedule():
    today = timezone.now().date()
    eight_months_ago = today - timezone.timedelta(days=8 * 30)
     # Filter activities created more than 8 months ago
    old_activities = Activity.objects.filter(created__lt=eight_months_ago)
    # Delete the selected activities
    old_activities.delete()