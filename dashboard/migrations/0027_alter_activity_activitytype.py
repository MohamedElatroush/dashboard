# Generated by Django 4.2.6 on 2023-11-08 23:43

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("dashboard", "0026_alter_activity_activitytype"),
    ]

    operations = [
        migrations.AlterField(
            model_name="activity",
            name="activityType",
            field=models.IntegerField(
                choices=[
                    (0, "H: holiday"),
                    (1, "C: in office"),
                    (2, "X: day off"),
                    (3, "J: home assignment"),
                ],
                default=1,
            ),
        ),
    ]
