# Generated by Django 4.2.6 on 2023-11-04 09:43

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ("dashboard", "0016_alter_user_grade"),
    ]

    operations = [
        migrations.RenameField(
            model_name="user", old_name="grade", new_name="userGrade",
        ),
    ]
