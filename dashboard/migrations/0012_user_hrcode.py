# Generated by Django 4.2.6 on 2023-11-03 09:53

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("dashboard", "0011_user_position"),
    ]

    operations = [
        migrations.AddField(
            model_name="user",
            name="hrCode",
            field=models.CharField(blank=True, max_length=5, null=True, unique=True),
        ),
    ]