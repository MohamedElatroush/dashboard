# Generated by Django 4.2.6 on 2024-06-04 19:02

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("dashboard", "0051_department_alter_activity_activitydate_user_dep"),
    ]

    operations = [
        migrations.AlterField(
            model_name="department",
            name="name",
            field=models.CharField(blank=True, max_length=256, null=True),
        ),
    ]
