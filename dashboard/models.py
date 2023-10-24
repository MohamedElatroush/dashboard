from django.db import models
from model_utils.models import TimeStampedModel
from phonenumber_field.modelfields import PhoneNumberField
from django.contrib.auth.models import AbstractUser
from .constants import constants

# Create your models here.
class User(AbstractUser, TimeStampedModel):
    id = models.AutoField(primary_key=True)
    phoneNumber = PhoneNumberField(blank=True, null=True)
    userType = models.IntegerField(choices=constants.USER_TYPE_CHOICES, default=constants.USER_TYPE_A)

    def __str__(self):
        return f'username: {self.username}'

# class Profile(TimeStampedModel):
#     firstName = models.CharField(max_length=200)
#     lastName = models

class Activity(TimeStampedModel):
    userActivity = models.TextField()
    user = models.ForeignKey(User, on_delete=models.CASCADE, null=True)
    
    def __str__(self):
        return f'User Activity by: {self.user.username}'
