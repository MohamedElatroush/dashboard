from rest_framework import serializers
from .models import User, Activity
from .constants import constants

class CreateUserSerializer(serializers.ModelSerializer):
    class Meta:
        model = User
        fields = ['username', 'password', 'email', 'phoneNumber', 'first_name', 'last_name', 'userType']

class ListUsersSerializer(serializers.ModelSerializer):
    user_type = serializers.SerializerMethodField()
    class Meta:
        model = User
        exclude =['password']
    def get_user_type(self, obj):
        for user_type, user_type_text in constants.USER_TYPE_CHOICES:

            if obj.userType == user_type:
                return user_type_text
        return _("Unknown User Type")

class UserDeleteSerializer(serializers.Serializer):
    userId = serializers.IntegerField()

class ActivitySerializer(serializers.ModelSerializer):
    username = serializers.SerializerMethodField()
    userType = serializers.SerializerMethodField()
    firstName = serializers.SerializerMethodField()
    lastName = serializers.SerializerMethodField()
    class Meta:
        model=Activity
        fields="__all__"
    def get_username(self, obj):
        # Access the related User object and retrieve the username
        return obj.user.username
    def get_userType(self, obj):
        return constants.USER_TYPE_CHOICES[obj.user.userType][1]
    def get_firstName(self, obj):
        return obj.user.first_name
    def get_lastName(self, obj):
        return obj.user.last_name

class ModifyActivitySerializer(serializers.ModelSerializer):
    class Meta:
        model=Activity
        fields = ['userActivity']

class MakeUserAdminSerializer(serializers.ModelSerializer):
    userId = serializers.IntegerField()
    class Meta:
        model=User
        fields = ['is_superuser', 'userId']