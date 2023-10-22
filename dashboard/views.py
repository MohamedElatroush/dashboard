from .models import User
from rest_framework import mixins
from rest_framework.response import Response
from rest_framework import permissions, status
from .serializers import CreateUserSerializer, ListUsersSerializer, UserDeleteSerializer
from rest_framework import viewsets
from rest_framework.decorators import action
from rest_framework.permissions import IsAuthenticated
from .constants import constants
# Create your views here.

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

