from .models import User, Activity
from rest_framework import mixins
from rest_framework.response import Response
from rest_framework import permissions, status
from .serializers import CreateUserSerializer, ListUsersSerializer, UserDeleteSerializer, ActivitySerializer,ModifyActivitySerializer
from rest_framework import viewsets
from rest_framework.decorators import action
from rest_framework.permissions import IsAuthenticated
from .constants import constants
from rest_framework_simplejwt.serializers import TokenObtainPairSerializer
from rest_framework_simplejwt.views import TokenObtainPairView
from rest_framework.authtoken.models import Token
from django.shortcuts import get_object_or_404
from django.contrib.auth.hashers import make_password
from django.contrib.auth.tokens import default_token_generator

class MyTokenObtainPairSerializer(TokenObtainPairSerializer):
    @classmethod
    def get_token(cls, user):
        token = super().get_token(user)
        token['username'] = user.username
        token['isAdmin'] = user.is_superuser
        return token
class MyTokenObtainPairView(TokenObtainPairView):
    serializer_class = MyTokenObtainPairSerializer

class UserRegistrationViewSet(viewsets.ViewSet):
    def create_user(self, request):
        data = request.data
        data['password'] = make_password(data['password'])  # Hash the password

        serializer = CreateUserSerializer(data=data)

        if serializer.is_valid():
            user = serializer.save()

            # Create a token for the user
            token = default_token_generator.make_token(user)

            return Response({
                'user': serializer.data,
                'token': token  # Include the token in the response
            }, status=status.HTTP_201_CREATED)

        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

class UserViewSet(viewsets.ModelViewSet):
    permission_classes=(IsAuthenticated,)
    queryset = User.objects.all()

    def get_serializer_class(self):
        if self.action == 'create' or self.action == 'update':
            return CreateUserSerializer  # Replace with the appropriate serializer
        else:
            return ListUsersSerializer  # Replace with the serializer for other actions

    @action(detail=False, methods=['delete'], url_path=r'delete_user/(?P<username>\w+(?:-\w+)*)')
    def delete_user(self, request, *args, **kwargs):
        username = self.kwargs['username']
        userObject = User.objects.filter(username=username).first()
        if not userObject:
            return Response(constants.CANT_DELETE_USER_ERROR, status=status.HTTP_400_BAD_REQUEST)
        userObject.delete()
        return Response(constants.SUCCESSFULLY_DELETED_USER, status=status.HTTP_200_OK)

    @action(detail=False, methods=['post'], serializer_class=CreateUserSerializer)
    def excel_sign_up(self, request, *args, **kwargs):
        """
        This endpoint creates bulk users by taking the first_name column and last_name column
        concatenate them to generate a username, and make a default password of 1234 that's required to be changed
        by the user
        """
        return Response(status=status.HTTP_200_OK)

class ActivityViewSet(viewsets.ModelViewSet):
    permission_classes=(IsAuthenticated,)
    queryset = Activity.objects.all()
    def get_serializer_class(self):
        return ActivitySerializer

    @action(detail=False, methods=['GET'])
    def get_activities(self, request, *args, **kwargs):
        activities = Activity.objects.all()
        serializer = ActivitySerializer(activities, many=True)
        return Response(serializer.data, status=status.HTTP_200_OK)

    @action(detail=False, methods=['PATCH', 'PUT'], serializer_class=ModifyActivitySerializer, url_path=r'update_activity/(?P<activityId>\w+(?:-\w+)*)')
    def update_activity(self, request, *args, **kwargs):
        """
            This endpoint allows the user owner of an activity to edit it
        """
        userId = request.user.id
        activityId = kwargs['activityId']
        activityObj = get_object_or_404(Activity, id=activityId)
        # Check if the user accessing the endpoint is the same one that created the activity or not
        if userId != activityObj.user.id:
            return Response(data=constants.NOT_ALLOWED_TO_ACCESS, status=status.HTTP_400_BAD_REQUEST)

        serializer = ModifyActivitySerializer(activityObj, data=request.data, partial=True)
        if serializer.is_valid():
            serializer.save()
            return Response(data=constants.SUCCESSFULLY_UPDATED_ACTIVITY, status=status.HTTP_200_OK)

        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

    @action(detail=False, methods=['POST'])
    def create_activity(self, request, *args, **kwargs):
        """
            This endpoint allows the user to create an activity
        """
        userId = request.user.id
        serializer = ModifyActivitySerializer(data=request.data)
        serializer.is_valid()
        if not serializer.is_valid():
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

        userActivity = serializer.validated_data['userActivity']
        new_activity = Activity.objects.create(userActivity=userActivity, user_id=userId)
        new_activity.save()
        return Response(constants.SUCCESSFULLY_CREATED_ACTIVITY, status=status.HTTP_201_CREATED)

