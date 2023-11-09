# Generated by Django 4.2.6 on 2023-11-04 09:50

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("dashboard", "0018_rename_usergrade_user_grade"),
    ]

    operations = [
        migrations.AddField(
            model_name="user",
            name="department",
            field=models.CharField(blank=True, max_length=256, null=True),
        ),
        migrations.AddField(
            model_name="user",
            name="expert",
            field=models.IntegerField(
                blank=True, choices=[(0, "EXP"), (1, "LOC")], null=True
            ),
        ),
        migrations.AddField(
            model_name="user",
            name="isMobilized",
            field=models.BooleanField(
                blank=True, choices=[(0, "Not Mobilized"), (1, "Mobilized")], null=True
            ),
        ),
        migrations.AddField(
            model_name="user",
            name="natGroup",
            field=models.CharField(blank=True, max_length=5, null=True),
        ),
        migrations.AddField(
            model_name="user",
            name="workingLocation",
            field=models.CharField(blank=True, max_length=256, null=True),
        ),
    ]