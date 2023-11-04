# Generated by Django 4.2.6 on 2023-11-03 09:48

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("dashboard", "0006_alter_user_email_alter_user_usertype"),
    ]

    operations = [
        migrations.RemoveField(model_name="user", name="userType",),
        migrations.AddField(
            model_name="user",
            name="department",
            field=models.CharField(default="23", max_length=256),
            preserve_default=False,
        ),
        migrations.AddField(
            model_name="user",
            name="expert",
            field=models.IntegerField(choices=[(0, "EXP"), (1, "LOC")], default="1"),
            preserve_default=False,
        ),
        migrations.AddField(
            model_name="user",
            name="grade",
            field=models.IntegerField(
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
                default="1",
            ),
            preserve_default=False,
        ),
        migrations.AddField(
            model_name="user",
            name="hrCode",
            field=models.CharField(blank=True, max_length=5, null=True, unique=True),
        ),
        migrations.AddField(
            model_name="user",
            name="isMobilized",
            field=models.BooleanField(
                choices=[(0, "Not Mobilized"), (1, "Mobilized")], default="1"
            ),
            preserve_default=False,
        ),
        migrations.AddField(
            model_name="user",
            name="natGroup",
            field=models.CharField(default="dskad", max_length=5),
            preserve_default=False,
        ),
        migrations.AddField(
            model_name="user",
            name="position",
            field=models.CharField(default="adsas", max_length=256),
            preserve_default=False,
        ),
        migrations.AddField(
            model_name="user",
            name="workingLocation",
            field=models.CharField(default="adsad", max_length=256),
            preserve_default=False,
        ),
        migrations.AlterField(
            model_name="user",
            name="email",
            field=models.EmailField(
                max_length=254, unique=True, verbose_name="email address"
            ),
        ),
    ]
