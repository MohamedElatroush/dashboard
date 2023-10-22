from rest_framework import serializers
from .models import User

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