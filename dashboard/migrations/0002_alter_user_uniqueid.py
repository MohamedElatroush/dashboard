# Generated by Django 4.2.6 on 2023-10-21 16:52

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("dashboard", "0001_initial"),
    ]

    operations = [
        migrations.AlterField(
            model_name="user",
            name="uniqueId",
            field=models.AutoField(primary_key=True, serialize=False),
        ),
    ]
