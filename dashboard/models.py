from django.db import models
from model_utils.models import TimeStampedModel
from phonenumber_field.modelfields import PhoneNumberField
from django.contrib.auth.models import AbstractUser
from .constants import constants
from django.db.models.signals import post_delete
from django.dispatch import receiver
from django.utils import timezone
from datetime import datetime, timedelta
from django.db.models.signals import pre_delete
from storages.backends.s3boto3 import S3Boto3Storage

# Create your models here.
class User(AbstractUser, TimeStampedModel):
    id = models.AutoField(primary_key=True)
    phoneNumber = PhoneNumberField(blank=True, null=True)
    email = models.EmailField(unique=True, blank=True, null=True)
    isAdmin = models.BooleanField(default=False)

    # work related fields
    grade = models.IntegerField(choices=constants.USER_GRADE_CHOICES, blank=True, null=True)
    hrCode = models.CharField(max_length=10, unique=True, blank=True, null=True)
    organizationCode = models.CharField(max_length=256, blank=True, null=True)
    position = models.CharField(max_length=256, null=True, blank=True)
    department = models.CharField(max_length=256, null=True, blank=True)
    natGroup = models.IntegerField(choices=constants.NAT_GROUP_CHOICES, null=True, blank=True)
    workingLocation = models.CharField(max_length=256, null=True, blank=True)
    expert = models.IntegerField(choices=constants.EXPERT_LOCAL_CHOICES, null=True, blank=True)
    mobilization = models.CharField(max_length=256, null=True, blank=True)
    company = models.IntegerField(choices=constants.COMPANY_CHOICES, null=True, blank=True)
    needsPasswordReset = models.BooleanField(default=True)
    calendarType = models.IntegerField(choices=constants.CALENDAR_CHOICES, null=True, blank=True)
    isDeleted = models.BooleanField(default=False)

    def get_grade(self):
        # Create a dict from USER_GRADE_CHOICES for reverse lookup
        grade_dict = dict(constants.USER_GRADE_CHOICES)
        # Return the display value for the grade
        return grade_dict.get(self.grade, None)
    def get_expert(self):
        # Create a dict from USER_GRADE_CHOICES for reverse lookup
        expert_dict = dict(constants.EXPERT_LOCAL_CHOICES)
        # Return the display value for the grade
        return expert_dict.get(self.expert, None)
    def get_natGroup(self):
        nat_dict = dict(constants.NAT_GROUP_CHOICES)
        return nat_dict.get(self.natGroup, None)
    def get_company(self):
        company_dict = dict(constants.COMPANY_CHOICES)
        return company_dict.get(self.company, None)

    def generate_hr_code(self):
        grade_value = constants.USER_GRADE_CHOICES[self.grade][1]
        grade_prefix = grade_value[:2]
        max_hr_code = User.objects.filter(hrCode__startswith=grade_prefix).order_by('-hrCode').first()
        if max_hr_code:
            numerical_part = int(max_hr_code.hrCode[len(grade_prefix):]) + 1
        else:
            numerical_part = 1
        while True:
            new_hr_code = f'{grade_prefix}{numerical_part:03d}'
            if new_hr_code not in hrHistory.objects.values_list('hrCode', flat=True):
                self.hrCode = new_hr_code
                break
            numerical_part += 1
    def get_full_name(self):
        return self.first_name + " " + self.last_name


    def save(self, *args, **kwargs):
        # if not self.pk: # New user
        if self.grade is None:
            self.hrCode = None
            self.expert = None
        else:
            if self.grade in [constants.GRADE_A_1, constants.GRADE_A_2, constants.GRADE_A_3]:
                self.expert = constants.EXPERT_USER
            elif self.grade in [constants.GRADE_B_1, constants.GRADE_B_2, constants.GRADE_B_3, constants.GRADE_B_4, constants.GRADE_B_5]:
                self.expert = constants.LOCAL_USER
            self.generate_hr_code()
        super().save(*args, **kwargs)
        # Add hrCode to hrHistory model
        hrHistory.objects.create(hrCode=self.hrCode)

    def __str__(self):
        return f'username: {self.username}'

class Activity(TimeStampedModel):
    userActivity = models.TextField(null=True, blank=True)
    activityType = models.IntegerField(choices=constants.ACTIVITY_TYPES_CHOICES, blank=False, null=False, default=constants.INCAIRO)
    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    activityDate = models.DateField(default=timezone.now().date())
    user_details = models.JSONField(null=True, blank=True)

    def get_activity_type(self):
        return constants.ACTIVITY_TYPES_CHOICES[self.activityType][1]

    def save_user_details(self):
        if self.user:
            user_details = {
                'id': self.user.id,
                'username': self.user.username,
                'fullName': self.user.get_full_name(),
                'email': self.user.email,
                'phoneNumber': str(self.user.phoneNumber) if self.user.phoneNumber else None,
                'isAdmin': self.user.isAdmin,
                'grade': str(self.user.get_grade()) if self.user.get_grade() else None,
                'hrCode': self.user.hrCode,
                'organizationCode': self.user.organizationCode,
                'position': self.user.position,
                'department': self.user.department,
                'natGroup': str(self.user.get_natGroup()) if self.user.get_natGroup() else None,
                'workingLocation': self.user.workingLocation,
                'expert': str(self.user.get_expert()) if self.user.get_expert() else None,
                'mobilization': self.user.mobilization,
                'company': str(self.user.get_company()) if self.user.get_company() else None,
                'needsPasswordReset': self.user.needsPasswordReset,
                'calendarType': self.user.calendarType,
            }
            self.user_details = user_details

    def get_working_days(self, year, month):
        # Check if the user is of type LOC
        if self.user and self.user.expert == constants.EXPERT.LOC:
            # Calculate the first and last day of the month
            first_day = datetime(year, month, 1)
            last_day = datetime(year, month + 1, 1) - timedelta(days=1)

            # Get all activities for the specified month
            activities = Activity.objects.filter(
                user=self.user,
                activityDate__range=(first_day, last_day)
            )

            # Define a function to check if a day is a working day
            def is_working_day(activity_type):
                return activity_type == constants.WORKING_DAY

            working_days_count = 0

            # Iterate over each day in the month
            current_day = first_day
            while current_day <= last_day:
                # Check if the current day is a Friday
                if current_day.weekday() == 4:  # 0 = Monday, 1 = Tuesday, ..., 6 = Sunday
                    # Check the activity type for Friday
                    friday_activity = activities.filter(activityDate=current_day).first()

                    if friday_activity:
                        # Check the conditions for counting Friday as a working day
                        if (
                            is_working_day(friday_activity.activityType) or
                            (current_day - timedelta(days=1)).weekday() == 3 and is_working_day(activities.filter(activityDate=current_day - timedelta(days=1)).first().activityType) and
                            (current_day + timedelta(days=1)).weekday() == 5 and is_working_day(activities.filter(activityDate=current_day + timedelta(days=1)).first().activityType)
                        ):
                            working_days_count += 1

                # Move to the next day
                current_day += timedelta(days=1)

            return working_days_count
        else:
            return 0

    def save(self, *args, **kwargs):
        self.save_user_details()
        super().save(*args, **kwargs)

    def __str__(self):
        if self.user:
            return f'User Activity by: {self.user.username} -- Date: {self.activityDate}'
        else:
            return f'User Activity (Deleted User) -- Date: {self.activityDate}'

class ActivityFile(TimeStampedModel):
    file = models.FileField(upload_to='reports/', storage=S3Boto3Storage())
    file_name = models.CharField(max_length=256, null=True, blank=True)
    company = models.IntegerField(choices=constants.COMPANY_CHOICES, null=True, blank=True)
    department = models.CharField(max_length=256, null=True, blank=True)

class hrHistory(TimeStampedModel):
    hrCode = models.CharField(max_length=256, unique=True)