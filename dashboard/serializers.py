from rest_framework import serializers
from .models import User, Activity, Department
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
    dep = serializers.SerializerMethodField()  # Add this field for dep
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

    def get_dep(self, obj):
        if obj.dep:  # Check if dep exists
            return obj.dep.name  # Return the name of the department
        return None  # Return None if dep is None
    

class ListDepSerializer(serializers.ModelSerializer):
    class Meta:
        model = Department
        fields = '__all__'


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
        if obj.user:
        # Access the related User object and retrieve the username
            return obj.user.username
        elif obj.user_details and 'username' in obj.user_details:
            # Extract the username from user_details JSON field
            return obj.user_details['username']
        else:
            # Return a default value or handle the case as needed
            return "User Deleted"
    def get_firstName(self, obj):
        if obj.user:
            return obj.user.first_name
        elif obj.user_details and 'fullName' in obj.user_details:
            # Extract the first part of the fullName as first_name
            return obj.user_details['fullName'].split()[0]
        else:
            # Return a default value or handle the case as needed
            return ""

    def get_lastName(self, obj):
        if obj.user:
            return obj.user.last_name
        elif obj.user_details and 'fullName' in obj.user_details:
            # Extract the last part of the fullName as last_name
            return obj.user_details['fullName'].split()[-1]
        else:
            # Return a default value or handle the case as needed
            return ""
    def get_hrCode(self, obj):
        if obj.user:
            return obj.user.hrCode
        elif obj.user_details and 'hrCode' in obj.user_details:
            # Extract the username from user_details JSON field
            return obj.user_details['hrCode']
        else:
            # Return a default value or handle the case as needed
            return "User Deleted"
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