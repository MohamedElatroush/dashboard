from rest_framework import serializers
from .models import User, Activity
from .constants import constants

class CreateUserSerializer(serializers.ModelSerializer):
    class Meta:
        model = User
        fields = ['username', 'password', 'email', 'phoneNumber', 'first_name', 'last_name', 'userType']

class ListUsersSerializer(serializers.ModelSerializer):
     class Meta:
        model = User
        fields = "__all__"

class UserDeleteSerializer(serializers.Serializer):
    username = serializers.CharField()

class ActivitySerializer(serializers.ModelSerializer):
    username = serializers.SerializerMethodField()
    userType = serializers.SerializerMethodField()
    class Meta:
        model=Activity
        fields="__all__"
    def get_username(self, obj):
        # Access the related User object and retrieve the username
        return obj.user.username
    def get_userType(self, obj):
        return constants.USER_TYPE_CHOICES[obj.user.userType][1]

class ModifyActivitySerializer(serializers.ModelSerializer):
    class Meta:
        model=Activity
        fields = ['userActivity']