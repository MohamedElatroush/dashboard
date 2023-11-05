from django.db import models
from model_utils.models import TimeStampedModel
from phonenumber_field.modelfields import PhoneNumberField
from django.contrib.auth.models import AbstractUser
from .constants import constants
from django.db.models.signals import post_delete
from django.dispatch import receiver
from .utilities.utilities import format_hr_codes

# Create your models here.
class User(AbstractUser, TimeStampedModel):
    id = models.AutoField(primary_key=True)
    phoneNumber = PhoneNumberField(blank=True, null=True)
    email = models.EmailField(unique=True, blank=True, null=True)

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
        print(self.grade)
        print('\n\n\n')

    # def generate_hr_code(self):
    #     if self.grade is not None and (self._state.adding or not self.hrCode):
    #         # Find the maximum count for the given grade
    #         max_count = User.objects.filter(grade=self.grade).aggregate(max_count=models.Max(models.F('hrCode')))
    #         max_count = max_count['max_count']
    #         if max_count:
    #             # Extract the numerical part and increment it
    #             numerical_part = int(max_count[2:])  # Assuming the format is "A2xxx" or similar
    #             next_numerical_part = numerical_part + 1
    #         else:
    #             # If no existing user with this grade, start with 1
    #             next_numerical_part = 1
    #         # Construct the new HR code
    #         grade_prefix = constants.USER_GRADE_CHOICES[self.grade][1]
    #         self.hrCode = f'{grade_prefix}{next_numerical_part:03d}'

    def save(self, *args, **kwargs):
        # if not self.pk: # New user
        if self.grade is None:
            self.hrCode = None
        else:
            self.generate_hr_code()
        super().save(*args, **kwargs)

    def __str__(self):
        return f'username: {self.username}'

class Activity(TimeStampedModel):
    userActivity = models.TextField()
    user = models.ForeignKey(User, on_delete=models.CASCADE, null=True)

    def __str__(self):
        return f'User Activity by: {self.user.username}'