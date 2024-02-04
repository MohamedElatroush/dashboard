# Generated by Django 4.2.6 on 2024-02-04 20:51

import datetime
from django.db import migrations, models
import storages.backends.s3


class Migration(migrations.Migration):
    dependencies = [
        ("dashboard", "0044_alter_activity_activitydate"),
    ]

    operations = [
        migrations.AlterField(
            model_name="activity",
            name="activityDate",
            field=models.DateField(default=datetime.date(2024, 2, 4)),
        ),
        migrations.AlterField(
            model_name="activityfile",
            name="file",
            field=models.FileField(
                storage=storages.backends.s3.S3Storage(), upload_to="reports/"
            ),
        ),
    ]
