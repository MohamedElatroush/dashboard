# Generated by Django 4.2.6 on 2023-11-03 10:22

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("dashboard", "0012_user_hrcode"),
    ]

    operations = [
        migrations.AlterField(
            model_name="user",
            name="grade",
            field=models.IntegerField(
                blank=True,
                choices=[
                    (0, "A1"),
                    (1, "A2"),
                    (2, "A3"),
                    (3, "A4"),
                    (4, "A5"),
                    (5, "B1"),
                    (6, "B2"),
                    (7, "B3"),
                    (8, "B4"),
                    (9, "B5"),
                ],
                null=True,
            ),
        ),
        migrations.AlterField(
            model_name="user",
            name="hrCode",
            field=models.CharField(blank=True, max_length=10, null=True, unique=True),
        ),
    ]