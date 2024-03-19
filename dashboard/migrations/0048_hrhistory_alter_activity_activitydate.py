# Generated by Django 4.2.6 on 2024-02-22 09:52

import datetime
from django.db import migrations, models
import django.utils.timezone
import model_utils.fields


class Migration(migrations.Migration):
    dependencies = [
        ("dashboard", "0047_activityfile_file_name_alter_activity_activitydate"),
    ]

    operations = [
        migrations.CreateModel(
            name="hrHistory",
            fields=[
                (
                    "id",
                    models.BigAutoField(
                        auto_created=True,
                        primary_key=True,
                        serialize=False,
                        verbose_name="ID",
                    ),
                ),
                (
                    "created",
                    model_utils.fields.AutoCreatedField(
                        default=django.utils.timezone.now,
                        editable=False,
                        verbose_name="created",
                    ),
                ),
                (
                    "modified",
                    model_utils.fields.AutoLastModifiedField(
                        default=django.utils.timezone.now,
                        editable=False,
                        verbose_name="modified",
                    ),
                ),
                ("hrCode", models.CharField(blank=True, max_length=256, null=True)),
            ],
            options={
                "abstract": False,
            },
        ),
        migrations.AlterField(
            model_name="activity",
            name="activityDate",
            field=models.DateField(default=datetime.date(2024, 2, 22)),
        ),
    ]