# Generated by Django 4.2.6 on 2023-10-22 20:05

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ("dashboard", "0002_alter_user_uniqueid"),
    ]

    operations = [
        migrations.RenameField(model_name="user", old_name="uniqueId", new_name="id",),
    ]
