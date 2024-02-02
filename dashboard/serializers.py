from rest_framework import serializers
from .models import User, Activity, ActivityFile
from .constants import constants
from gettext import gettext as _
from datetime import date

class CreateUserSerializer(serializers.ModelSerializer):
    class Meta:
        model = User
        fields = ['username', 'password', 'email', 'phoneNumber', 'first_name', 'last_name', 'grade', 'organizationCode',\
                   'position', 'department', 'natGroup', 'workingLocation', 'mobilization', 'company','isDeleted']

class EditUserSerializer(serializers.ModelSerializer):
    class Meta:
        model = User
        fields = ['grade', 'organizationCode','position', 'department', 'natGroup', 'workingLocation', 'mobilization', 'company']

class ListUsersSerializer(serializers.ModelSerializer):
    grade = serializers.SerializerMethodField()
    expert = serializers.SerializerMethodField()
    natGroup = serializers.SerializerMethodField()
    company = serializers.SerializerMethodField()
    class Meta:
        model = User
        exclude =['password']

    def get_grade(self, obj):
        return obj.get_grade()
    def get_expert(self,obj):
        return obj.get_expert()
    def get_natGroup(self,obj):
        return obj.get_natGroup()
    def get_company(self,obj):
        return obj.get_company()

class UserDeleteSerializer(serializers.Serializer):
    userId = serializers.IntegerField()

class ActivitySerializer(serializers.ModelSerializer):
    username = serializers.SerializerMethodField()
    hrCode = serializers.SerializerMethodField()
    firstName = serializers.SerializerMethodField()
    lastName = serializers.SerializerMethodField()
    activityType = serializers.SerializerMethodField()
    class Meta:
        model=Activity
        fields="__all__"
    def get_username(self, obj):
        # Access the related User object and retrieve the username
        return obj.user.username
    # def get_userType(self, obj):
    #     return constants.USER_TYPE_CHOICES[obj.user.userType][1]
    def get_firstName(self, obj):
        return obj.user.first_name
    def get_lastName(self, obj):
        return obj.user.last_name
    def get_hrCode(self, obj):
        return obj.user.hrCode
    def get_activityType(self, obj):
        activity_type = obj.activityType
        if activity_type is not None:
            return constants.ACTIVITY_TYPES_CHOICES[activity_type][1]
        return None

class CreateActivitySerializer(serializers.ModelSerializer):
    activityDate = serializers.DateField(default=date.today)
    class Meta:
        model=Activity
        fields = ['userActivity', 'activityType', 'activityDate']

class MakeUserAdminSerializer(serializers.ModelSerializer):
    userId = serializers.IntegerField()
    class Meta:
        model=User
        fields = ['isAdmin', 'userId']

class ChangePasswordSerializer(serializers.Serializer):
    password = serializers.CharField()

class UserTimeSheetSerializer(serializers.Serializer):
    date = serializers.DateField(default=date.today)

class CalculateActivitySerializer(serializers.Serializer):
    date = serializers.DateField(default=date.today)

class EditActivitySerializer(serializers.ModelSerializer):
    class Meta:
        model=Activity
        fields=['userActivity', 'activityType', 'activityDate']

class ActivityFileSerializer(serializers.ModelSerializer):
    class Meta:
        model = ActivityFile
        fields = ('file', 'created')