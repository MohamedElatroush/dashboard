from .models import User, Activity
from rest_framework.response import Response
from rest_framework import permissions, status
from .serializers import CreateUserSerializer, ListUsersSerializer, UserDeleteSerializer,\
      ActivitySerializer,ModifyActivitySerializer, MakeUserAdminSerializer
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
from datetime import date, datetime 
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from io import BytesIO
from django.http import HttpResponse
import pytz
import pandas as pd
from .utilities import utilities

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
                # 'user': serializer.data,
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

    @action(detail=False, methods=['GET'])
    def get_users(self, request, *args, **kwargs):
        userId = request.user.id
        user = User.objects.filter(id=userId).first()

        # if not super user dont show users
        if not user.is_superuser:
            return Response(status=status.HTTP_403_FORBIDDEN)
        users = User.objects.all().order_by('hrCode')
        serializer = self.get_serializer(users, many=True)
        return Response(serializer.data, status=status.HTTP_200_OK)

    @action(detail=False, methods=['GET'])
    def get_user(self, request, *args, **kwargs):
        userId = request.user.id
        user = User.objects.filter(id=userId).first()
        serializer = self.get_serializer(user)
        return Response(serializer.data, status=status.HTTP_200_OK)

    @action(detail=False, methods=['PATCH'])
    def make_admin(self, request, *args, **kwargs):
        adminId = request.user.id
        user = User.objects.filter(id=adminId).first()
        # if not super user dont show users
        if not user.is_superuser:
            return Response(status=status.HTTP_403_FORBIDDEN)

        serializer = MakeUserAdminSerializer(data=request.data)
        serializer.is_valid()
        if not serializer.is_valid():
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)
        user_id = serializer.validated_data['userId']
        is_user_super = serializer.validated_data['is_superuser']

        modified_user = User.objects.filter(id=user_id).first()
        modified_user.is_staff = is_user_super
        modified_user.is_superuser = is_user_super
        modified_user.save()
        return Response(status=status.HTTP_200_OK)

    @action(detail=False, methods=['delete'])
    def delete_user(self, request, *args, **kwargs):
        adminId = request.user.id
        user = User.objects.filter(id=adminId).first()
        if not user.is_superuser:
            return Response(status=status.HTTP_403_FORBIDDEN)

        serializer = UserDeleteSerializer(data=request.data)  # Pass request data to the serializer

        if serializer.is_valid():  # Validate the data
            user_id = serializer.validated_data['userId']
            userObject = User.objects.filter(id=user_id).first()

            if not userObject:
                return Response(status=status.HTTP_400_BAD_REQUEST)
             # Check if the user to be deleted is a superuser
            if userObject.is_superuser:
                superusers_count = User.objects.filter(is_superuser=True).count()
                if superusers_count <= 1:
                    return Response(constants.LAST_SUPERUSER_DELETION_ERROR, status=status.HTTP_400_BAD_REQUEST)

            userObject.delete()
            return Response(constants.SUCCESSFULLY_DELETED_USER, status=status.HTTP_200_OK)
        else:
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

    @action(detail=False, methods=['POST'])
    def reset_user_password(self, request, *args, **kwargs):
        adminId = request.user.id
        user = User.objects.filter(id=adminId).first()

         # if not super user dont show users
        if not user.is_superuser:
            return Response(status=status.HTTP_403_FORBIDDEN)
        serializer = UserDeleteSerializer(data=request.data)
        serializer.is_valid()
        if not serializer.is_valid():
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

        user_id = serializer.validated_data['userId']
        userObject = User.objects.filter(id=user_id).first()
        if not userObject:
            return Response(constants.CANT_DELETE_USER_ERROR, status=status.HTTP_400_BAD_REQUEST)
        userObject.password = make_password("1234")
        userObject.save()
        return Response(constants.SUCCESSFULLY_DELETED_USER, status=status.HTTP_200_OK)


    @action(detail=False, methods=['post'], serializer_class=CreateUserSerializer)
    def excel_sign_up(self, request, *args, **kwargs):
        """
        This endpoint creates bulk users by taking the first_name column and last_name column
        concatenate them to generate a username, and make a default password of 1234 that's required to be changed
        by the user
        """
        # Check if user is not an admin
        if not utilities.check_is_admin(request.user.id):
            return Response(status=status.HTTP_403_FORBIDDEN)
        # Check if an Excel file was uploaded
        if 'file' not in request.FILES:
            return Response({'error': 'No file uploaded'}, status=status.HTTP_400_BAD_REQUEST)
        excel_file = request.FILES['file']
        try:
            all_usernames = set(User.objects.all().values_list('username', flat=True))
            # Read the Excel file again, this time setting the second row as headers
            dataframe_with_headers = pd.read_excel(excel_file, header=2)

            # Find the column index for "HR Code" and "Remarks"
            hr_code_index = dataframe_with_headers.columns.get_loc("HR Code")
            remarks_index = dataframe_with_headers.columns.get_loc("Remarks") if "Remarks" in dataframe_with_headers.columns else None

            # If "Remarks" column is not found, we will display all columns up to the last one.
            if remarks_index is None:
                remarks_index = len(dataframe_with_headers.columns) - 1

            # Display the new column headers and a sample of the data to confirm
            dataframe_hr_to_remarks = dataframe_with_headers.iloc[:, hr_code_index:remarks_index+1]
            for index, row in dataframe_hr_to_remarks.iterrows():
                if pd.isna(row['Name']):
                    continue
                # Extract information from the row
                hr_code = row['HR Code']
                name = row['Name']
                email = row['Email Address']
                grade = row['Grade']
                organization_code = row['Organization Code']
                position = row['Position']
                department = row['Department']
                nat_group = row['NAT Group']
                working_location = row['Working Location']
                expert = row['Expert']
                mobilization_status = row['Mobilization status']
                company = row['Company']

                # Assuming you have a function to convert the grade and expert to the corresponding integer choices
                grade_choice = utilities.convert_grade_to_choice(grade)
                expert_choice = utilities.convert_expert_to_choice(expert)
                nat_group_choice = utilities.convert_nat_group_to_choice(nat_group)
                company_choice = utilities.convert_company_to_choice(company)

                email = row['Email Address'] if pd.notnull(row['Email Address']) else None
                # Create a new User object
                user = User(
                    username = utilities.generate_username_from_name(name, all_usernames),  # Assuming the username is the name, you might want to format this.
                    email=email,
                    first_name=utilities.generate_first_name(name),
                    last_name=utilities.generate_last_name(name),
                    grade=grade_choice,
                    hrCode=hr_code,  # This will be overridden if grade is set and hrCode is not.
                    organizationCode=organization_code,
                    position=position,
                    department=department,
                    natGroup=nat_group_choice,
                    workingLocation=working_location,
                    expert=expert_choice,
                    company=company_choice,
                    mobilization=mobilization_status,
                    password=make_password("1234")
                )

                # Save the new user, the save method will call generate_hr_code if needed
                user.save()


            return Response({'message': 'Data extracted and printed successfully'}, status=status.HTTP_200_OK)

        except Exception as e:
            return Response({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

class ActivityViewSet(viewsets.ModelViewSet):
    permission_classes=(IsAuthenticated,)
    queryset = Activity.objects.all()
    def get_serializer_class(self):
        return ActivitySerializer

    @action(detail=False, methods=['GET'])
    def get_activities(self, request, *args, **kwargs):
        date_param = request.query_params.get('date', None)

        if date_param:
        # If a date is provided, parse it and fetch activities for that date
            try:
                selected_date = datetime.strptime(date_param, '%Y-%m-%d').date()
            except ValueError:
                return Response({'error': 'Invalid date format. Please use YYYY-MM-DD format.'}, status=status.HTTP_400_BAD_REQUEST)
        else:
            # If no date is provided, default to today's date
            selected_date = date.today()
        activities = Activity.objects.filter(created__date=selected_date).order_by('-created')
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

    def _create_activity_excel_report(self, activities, selected_date, cairo_timezone):
        # Create a new Excel workbook and add a worksheet
        wb = Workbook()
        ws = wb.active
        ws.title = "Activity Report"

        # Write column headers from the constants
        for col_num, column_title in enumerate(constants.EXPORT_ACTIVITY_COLUMNS, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.value = column_title
            cell.alignment = Alignment(horizontal="center")

        # Create a mapping dictionary for user types
        user_type_mapping = {user_type[0]: user_type[1] for user_type in constants.USER_TYPE_CHOICES}

        # Populate the worksheet with activity data
        for row_num, activity in enumerate(activities, 2):
            # Convert the datetime to Cairo timezone and then format it
            created_cairo = activity.created.astimezone(cairo_timezone)
            ws.cell(row=row_num, column=1, value=activity.user.username)
            # ws.cell(row=row_num, column=2, value=str(constants.USER_TYPE_CHOICES[activity.user.userType][1]))
            ws.cell(row=row_num, column=3, value=created_cairo.strftime("%Y-%m-%d"))
            ws.cell(row=row_num, column=4, value=created_cairo.strftime("%H:%M:%S"))
            ws.cell(row=row_num, column=5, value=activity.userActivity)

        excel_data = BytesIO()
        wb.save(excel_data)
        excel_data.seek(0)
        response = HttpResponse(excel_data, content_type="application/ms-excel")
        response["Content-Disposition"] = f'attachment; filename=activity_report_{selected_date}.xlsx'
        return response

    @action(detail=False, methods=['GET'])
    def export_all(self, request, *args, **kwargs):
        adminId = request.user.id
        userObj = User.objects.filter(id=adminId).first()
        if not userObj.is_superuser:
            return Response(constants.NOT_ALLOWED_TO_ACCESS, status=status.HTTP_400_BAD_REQUEST)
        selected_date = date.today()
        cairo_timezone = pytz.timezone('Africa/Cairo')
        activities = Activity.objects.filter(created__date=selected_date).order_by('-created')

        response = self._create_activity_excel_report(activities, selected_date, cairo_timezone)
        return response

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

    @action(detail=False, methods=['DELETE'], url_path=r'delete_activity/(?P<activityId>\w+(?:-\w+)*)')
    def delete_activity(self, request, *args, **kwargs):
        """
            This endpoint allows the user/admin to delete an activity
        """
        userId = request.user.id

        activity = Activity.objects.filter(id=kwargs['activityId']).first()
        user = User.objects.filter(id=userId).first()

        if not activity:
            return Response(status=status.HTTP_404_NOT_FOUND)

        if userId == activity.user.id or user.is_superuser:
            activity.delete()
            return Response(constants.SUCCESSFULLY_DELETED_ACTIVITY, status=status.HTTP_200_OK)

        return Response(status=status.HTTP_400_BAD_REQUEST)

