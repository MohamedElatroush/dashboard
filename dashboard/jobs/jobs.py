from django.utils import timezone
from ..models import Activity
from datetime import datetime

def schedule():
    today = timezone.now().date()
    last_month = today - timezone.timedelta(days=today.day)  # Calculate the first day of the last month
    last_month_start = datetime.combine(last_month, datetime.min.time())
    last_month_end = datetime.combine(today, datetime.min.time())
    Activity.objects.filter(created__gte=last_month_start, created__lt=last_month_end).exclude(created=today).delete()