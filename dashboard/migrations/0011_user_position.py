# Generated by Django 4.2.6 on 2023-11-03 09:52

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("dashboard", "0010_alter_user_grade"),
    ]

    operations = [
        migrations.AddField(
            model_name="user",
            name="position",
            field=models.CharField(blank=True, max_length=256, null=True),
        ),
    ]
