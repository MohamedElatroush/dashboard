# Generated by Django 4.2.6 on 2024-01-04 22:23

import datetime
from django.db import migrations, models


class Migration(migrations.Migration):
    dependencies = [
        ("dashboard", "0040_activity_user_details"),
    ]

    operations = [
        migrations.AlterField(
            model_name="activity",
            name="activityDate",
            field=models.DateField(default=datetime.date(2024, 1, 4)),
        ),
    ]
