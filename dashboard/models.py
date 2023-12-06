from django.db import models
from model_utils.models import TimeStampedModel
from phonenumber_field.modelfields import PhoneNumberField
from django.contrib.auth.models import AbstractUser
from .constants import constants
from django.db.models.signals import post_delete
from django.dispatch import receiver
from django.utils import timezone
from datetime import datetime

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
        self.hrCode = f'{grade_prefix}{numerical_part:03d}'
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

    def __str__(self):
        return f'username: {self.username}'

class Activity(TimeStampedModel):
    userActivity = models.TextField(null=True, blank=True)
    activityType = models.IntegerField(choices=constants.ACTIVITY_TYPES_CHOICES, blank=False, null=False, default=constants.INOFFICE)
    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    activityDate = models.DateField(default=timezone.now().date())

    def get_activity_type(self):
        return constants.ACTIVITY_TYPES_CHOICES[self.activityType][1]

    def __str__(self):
        return f'User Activity by: {self.user.username} -- Date: {self.activityDate}'
