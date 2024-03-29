from .models import User, Activity, hrHistory
from rest_framework.response import Response
from rest_framework import status
from .serializers import CreateUserSerializer,\
      ListUsersSerializer, UserDeleteSerializer,\
      ActivitySerializer, CreateActivitySerializer,\
    MakeUserAdminSerializer, ChangePasswordSerializer,\
    UserTimeSheetSerializer, EditUserSerializer,\
        CalculateActivitySerializer, EditActivitySerializer
from rest_framework import viewsets
from rest_framework.decorators import action
from rest_framework.permissions import IsAuthenticated
from .constants import constants
from rest_framework_simplejwt.serializers import TokenObtainPairSerializer
from rest_framework_simplejwt.views import TokenObtainPairView
from django.shortcuts import get_object_or_404
from django.contrib.auth.hashers import make_password
from django.contrib.auth.tokens import default_token_generator
from datetime import date, datetime, timedelta
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill
from io import BytesIO
from django.http import HttpResponse
import pandas as pd
from .utilities import utilities
from openpyxl.styles import Font, Border, Side
from django.db import transaction
from openpyxl.drawing.image import Image
import os
import calendar
import numpy as np
from django.utils import timezone
import calendar
from dashboard.jobs.jobs import generate_noce_timesheet
from dateutil.relativedelta import relativedelta

class MyTokenObtainPairSerializer(TokenObtainPairSerializer):
    @classmethod
    def get_token(cls, user):
        token = super().get_token(user)
        token['username'] = user.username
        token['isAdmin'] = user.isAdmin
        token['is_superuser'] = user.is_superuser
        token['calendarType'] = user.calendarType
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
        if not (user.is_superuser or user.isAdmin):
            return Response(status=status.HTTP_403_FORBIDDEN)
        users = User.objects.filter(isDeleted=False).order_by('created')
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
        is_admin = serializer.validated_data['isAdmin']

        modified_user = User.objects.filter(id=user_id).first()
        modified_user.isAdmin = is_admin
        modified_user.save()
        return Response(status=status.HTTP_200_OK)

    @action(detail=False, methods=['PATCH'])
    def revoke_admin(self, request, *args, **kwargs):
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

        modified_user = User.objects.filter(id=user_id).first()
        modified_user.isAdmin = False
        modified_user.save()
        return Response(status=status.HTTP_200_OK)

    # An endpoint for superusers only to edit any other user (grade, organizationcode..etc)
    @action(detail=False, methods=['patch'], url_path=r'edit_details/(?P<userId>\w+(?:-\w+)*)')
    def edit_details(self, request, *args, **kwargs):
        user_id = kwargs['userId']
        request_user = request.user

        if not request_user.is_superuser:
            return Response(status=status.HTTP_403_FORBIDDEN)

        user = User.objects.filter(id=user_id).first()
        if not user:
            return Response(status=status.HTTP_404_NOT_FOUND)

        serializer = EditUserSerializer(user, data=request.data, partial=True)
        if serializer.is_valid():
            serializer.save()
            return Response(status=status.HTTP_200_OK)
        else:
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

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

            # Instead of deleting, mark the user as deleted
            userObject.isDeleted = True
            userObject.save()
            # Clear specific fields for the deleted user
            with transaction.atomic():
                userObject.grade = None
                userObject.hrCode = None
                userObject.organizationCode = None
                userObject.save()
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
            return Response(constants.CANT_RESET_USER_PASSWORD_ERROR, status=status.HTTP_400_BAD_REQUEST)
        userObject.password = make_password("1234")
        userObject.save()
        return Response(constants.SUCCESSFULLY_DELETED_USER, status=status.HTTP_200_OK)

    @transaction.atomic
    @action(detail=False, methods=['post'])
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
            return Response({'detail': 'No file uploaded'}, status=status.HTTP_400_BAD_REQUEST)
        excel_file = request.FILES['file']

        unavailable_codes = hrHistory.objects.all().values_list('hrCode', flat=True)

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

                if hr_code in unavailable_codes:
                    return Response(constants.ERR_HR_CODE_USED, status=status.HTTP_400_BAD_REQUEST)

                # Assuming you have a function to convert the grade and expert to the corresponding integer choices
                grade_choice = utilities.convert_grade_to_choice(grade)
                expert_choice = utilities.convert_expert_to_choice(expert)
                nat_group_choice = utilities.convert_nat_group_to_choice(nat_group)
                company_choice = utilities.convert_company_to_choice(company)

                if expert_choice == constants.EXPERT_USER:
                    calendar_type = constants.CALENDAR_TYPE_EXPERT
                elif expert_choice == constants.LOCAL_USER:
                    calendar_type = constants.CALENDAR_TYPE_LOCAL
                else:
                    calendar_type = None

                email = row['Email Address'] if pd.notnull(row['Email Address']) else None
                new_username = utilities.generate_username_from_name(name, all_usernames)
                if User.objects.filter(username=new_username).exists() or User.objects.filter(email=email).exists():
                    # Skip this row if the username already exists
                    continue

                # Create a new User object
                user = User(
                    username = new_username,  # Assuming the username is the name, you might want to format this.
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
                    calendarType=calendar_type,
                    password=make_password("1234")
                )

                # Save the new user, the save method will call generate_hr_code if needed
                user.save()
            return Response({'message': 'Data extracted and printed successfully'}, status=status.HTTP_200_OK)
        except Exception as e:
            return Response({'detail': str(e)}, status=status.HTTP_400_BAD_REQUEST)

    @action(detail=False, methods=['PATCH'], url_path=r'change_password/(?P<userId>\w+(?:-\w+)*)')
    def change_password(self, request, *args, **kwargs):
        request_user_id = request.user.id
        target_user_id = kwargs['userId']
        # Check if the user accessing the endpoint is the same one that created the activity or not
        if str(request_user_id) != str(target_user_id):
            return Response(data=constants.NOT_ALLOWED_TO_ACCESS, status=status.HTTP_400_BAD_REQUEST)

        serializer = ChangePasswordSerializer(data=request.data)
        if serializer.is_valid():
            User.objects.filter(id=target_user_id).update(needsPasswordReset=False,\
                                                          password=make_password(serializer.validated_data['password']))
            return Response(data=constants.SUCCESSFULLY_CHANGED_PASSWORD, status=status.HTTP_200_OK)
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)
    
    def __add_local_working_days_individual__(self, current_date, user, cover_ws):
        start_date = current_date.replace(day=1)
        last_day_of_month = calendar.monthrange(current_date.year, current_date.month)[1]
        end_date = (current_date + timedelta(days=last_day_of_month)).replace(day=1)

        # Adjust end_date to the next day
        end_date = end_date + timedelta(days=1)

        total_working_days_expert = np.busday_count(start_date, end_date, weekmask='0011111')
        cover_ws['F38'].value = total_working_days_expert
        cover_ws['F38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        start_date = current_date.replace(day=1)
        last_day_of_month = calendar.monthrange(current_date.year, current_date.month)[1]
        end_date = current_date.replace(day=last_day_of_month)

        last_day_of_month = calendar.monthrange(current_date.year, current_date.month)[1]
        end_date = current_date.replace(day=last_day_of_month) + timedelta(days=1)
        total_working_days = np.busday_count(start_date, end_date, weekmask='1111111')

        # Iterate through each day in the range
        for day in range(1, last_day_of_month + 1):
            day_off = Activity.objects.filter(
                user=user,
                activityDate__day=current_date.day,
                activityDate__month=current_date.month,
                activityDate__year=current_date.year,
                activityType=constants.OFFDAY,
            ).exists()

            if day_off and (day == 5 and day - 1 in (4, 6) and day + 1 in (4, 6)):
                total_working_days -= 1

            # Filter activities for the current user and month excluding 'H' type activities
            activities = Activity.objects.filter(
                user=user,
                activityDate__month=current_date.month,
                activityDate__year=current_date.year,
            ).filter(activityType__in=[constants.INCAIRO, constants.HOLIDAY])

            # Count the number of activities
            working_days = activities.count()
            # Iterate over weeks in the month
            for week in calendar.monthcalendar(current_date.year, current_date.month):
                for day in week:
                    # Check if Thursday and Saturday are off-days in the current week
                    if day != 0:  # Ignore days that belong to the previous or next month (represented as 0)
                        thursday_offday = any(activity.activityType == constants.OFFDAY for activity in activities.filter(activityDate__day=day + 3))
                        saturday_offday = any(activity.activityType == constants.OFFDAY for activity in activities.filter(activityDate__day=day + 5))

                        # If either Thursday or Saturday is off-day, decrement working_days
                        if thursday_offday or saturday_offday:
                            working_days -= 1

            # Add "0" under "Cairo" for local users
            cover_ws['D38'].value = working_days
            cover_ws['D38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text
            cover_ws['G38'].value = total_working_days
            cover_ws['G38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

            start_date = current_date.replace(day=1)
            end_date = current_date.replace(day=calendar.monthrange(current_date.year, current_date.month)[1])
            end_date = end_date + relativedelta(days=1)

            activities_japan = Activity.objects.filter(
                user=user,
                activityDate__month=current_date.month,
                activityDate__year=current_date.year,
            ).filter(activityType__in=[constants.HOMEASSIGN]).count()

            cover_ws['C38'].value = activities_japan
            cover_ws['C38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

            cover_ws['J38'].value = round(cover_ws['D38'].value / cover_ws['G38'].value, 3)
            cover_ws['J38'].font = Font(size=11)
            cover_ws['J38'].alignment = Alignment(horizontal='center', vertical='center')

            # Japan NOD/TCD
            cover_ws['I38'].value = round(cover_ws['C38'].value / cover_ws['F38'].value, 3)
            cover_ws['I38'].font = Font(size=11)
            cover_ws['I38'].alignment = Alignment(horizontal='center', vertical='center')

    def __add_expert_working_days_individual__(self, current_date, user, cover_ws):
            ###### LOCAL ######
            start_date = current_date.replace(day=1)
            last_day_of_month = calendar.monthrange(current_date.year, current_date.month)[1]
            end_date = current_date.replace(day=last_day_of_month)

            last_day_of_month = calendar.monthrange(current_date.year, current_date.month)[1]
            end_date = current_date.replace(day=last_day_of_month) + timedelta(days=1)
            total_working_days_cairo = np.busday_count(start_date, end_date, weekmask='1111111')
            ###### LOCAL ######
            start_date = current_date.replace(day=1)
            last_day_of_month = calendar.monthrange(current_date.year, current_date.month)[1]
            end_date = (current_date + timedelta(days=last_day_of_month)).replace(day=1)

            # Adjust end_date to the next day
            end_date = end_date + timedelta(days=1)
            total_working_days = np.busday_count(start_date, end_date, weekmask='0011111')

            # Filter activities for the current user and month excluding 'H' type activities
            activities = Activity.objects.filter(
                user=user,
                activityDate__month=current_date.month,
                activityDate__year=current_date.year,
            ).filter(activityType__in=[constants.HOMEASSIGN])

            # Count the number of activities
            working_days = activities.count()

            cover_ws['C38'].value = working_days
            cover_ws['C38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

            cover_ws['F38'].value = total_working_days
            cover_ws['F38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

            start_date = current_date.replace(day=1)
            end_date = current_date.replace(day=calendar.monthrange(current_date.year, current_date.month)[1])
            end_date = end_date + relativedelta(days=1)

            activities_cairo_count = Activity.objects.filter(
                user=user,
                activityDate__month=current_date.month,
                activityDate__year=current_date.year,
                activityType__in=[constants.INCAIRO, constants.HOLIDAY]
            ).count()

            # Cairo NOD
            cover_ws['D38'].value = activities_cairo_count
            cover_ws['D38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

            # Cairo TCD
            cover_ws['G38'].value = total_working_days_cairo
            cover_ws['G38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

            # Japan NOD/TCD
            cover_ws['I38'].value = round(cover_ws['C38'].value / cover_ws['F38'].value, 3)
            cover_ws['I38'].font = Font(size=11)
            cover_ws['I38'].alignment = Alignment(horizontal='center', vertical='center')

            # Cairo NOD/TCD
            cover_ws['J38'].value = round(cover_ws['D38'].value / cover_ws['G38'].value, 3)
            cover_ws['J38'].font = Font(size=11)
            cover_ws['J38'].alignment = Alignment(horizontal='center', vertical='center')

    def __add_individual_timesheet__(self, wb, current_month_name, current_year, user, current_date, current_month):
        # Create cover page
        start_date = current_date.replace(day=1)
        last_day_of_month = calendar.monthrange(current_date.year, current_date.month)[1]
        end_date = current_date.replace(day=last_day_of_month)

        cover_ws = wb.create_sheet(title=str(user.first_name))

        cover_ws.merge_cells('A3:G3')
        cover_ws['A3'].value = constants.COVER_TS_TEXT
        cover_ws['A6'].font = Font(size=11)
        cover_ws['A3'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        cover_ws.merge_cells('A6:G6')  # Adjust the cell range as needed
        cover_ws['A6'].value = 'Monthly Time Sheet'
        cover_ws['A6'].font = Font(size=16, italic=True)  # Set font size and make it italic
        cover_ws['A6'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        cover_ws['A7'].value = 'Month:'
        cover_ws.merge_cells('B7:C7')
        cover_ws['B7'].value = current_month_name
        cover_ws['B7'].font = Font(size=12, italic=True, bold=True)  # Set font size and make it italic
        cover_ws['B7'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        # add the year
        cover_ws['F7'].value = 'Year:'
        cover_ws['G7'].value = current_year
        cover_ws['G7'].font = Font(size=12, italic=True, bold=True)  # Set font size and make it italic
        cover_ws['G7'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text
        cover_ws['F7'].font = Font(size=12, italic=True, bold=True)  # Set font size and make it italic
        cover_ws['F7'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        # add the logo
        script_directory = os.path.dirname(os.path.abspath(__file__))
        parent_parent_directory = os.path.dirname(script_directory)
        logo_path = os.path.join(parent_parent_directory, 'static', 'images', 'logo.png')
        img = Image(logo_path)
        cover_ws.add_image(img, 'I2')

        # add the name
        cover_ws.merge_cells('B9:C9')
        cover_ws['A9'].value = 'Name:'
        cover_ws['B9'].value = str(user.get_full_name())

        user_grade = user.get_grade()

        cover_ws['A11'].value = 'Category:'
        cover_ws['D11'].value = 'A1'
        cover_ws['F11'].value = 'A2'
        cover_ws['H11'].value = 'A3'
        cover_ws['D13'].value = 'B1'
        cover_ws['F13'].value = 'B2'
        cover_ws['H13'].value = 'B3'
        cover_ws['D15'].value = 'B4'
        cover_ws['F15'].value = 'B5'

        # Mapping of grades to corresponding columns
        grade_mapping = {
            'A1': 'C',
            'A2': 'E',
            'A3': 'G',
            'B1': 'C',
            'B2': 'E',
            'B3': 'G',
            'B4': 'C',
            'B5': 'E',
        }

        # Set the letter '/' in the corresponding cell based on the user's grade
        if user_grade in grade_mapping:
            grade_column = grade_mapping[user_grade]
            cell_address = f'{grade_column}11'
            cover_ws[cell_address].value = '/'
            cover_ws[cell_address].alignment = Alignment(horizontal='center', vertical='center')

        # Set borders for various cells
        utilities.set_borders(cover_ws, 11, ['C', 'E', 'G'])
        utilities.set_borders(cover_ws, 13, ['C', 'E', 'G'])
        utilities.set_borders(cover_ws, 15, ['C', 'E'])

        for col in ['C', 'E', 'G']:
            cell_address = f'{col}11'
            cover_ws[cell_address].border = Border(top=Side(style='thin', color='000000'), left=Side(style='thin', color='000000'),
                                                right=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))
        for col in ['C', 'E', 'G']:
            cell_address = f'{col}13'
            cover_ws[cell_address].border = Border(top=Side(style='thin', color='000000'), left=Side(style='thin', color='000000'),
                                                right=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))
        for col in ['C', 'E']:
            cell_address = f'{col}15'
            cover_ws[cell_address].border = Border(top=Side(style='thin', color='000000'), left=Side(style='thin', color='000000'),
                                                right=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))

        cover_ws['A18'].value = 'Nationality:'
        cover_ws['A18'].font = Font(size=11, italic=True, bold=True)

        cover_ws['D18'].value = 'Expatriate'
        cover_ws['D18'].font = Font(size=10, bold=True)

        cover_ws['F18'].value = 'Local'
        cover_ws['F18'].font = Font(size=10, bold=True)

        # Get the user's nationality
        user_nationality = user.get_expert()

        # Mapping of nationalities to corresponding columns
        nationality_mapping = {
            'EXP': 'C',  # Column for Expatriate
            'LOC': 'E',  # Column for Local
        }

        # Set the letter '/' in the corresponding cell and center it
        if user_nationality in nationality_mapping:
            nationality_column = nationality_mapping[user_nationality]
            cell_address = f'{nationality_column}18'
            cover_ws[cell_address].value = '/'
            cover_ws[cell_address].alignment = Alignment(horizontal='center', vertical='center')

            # Apply border
            cover_ws[cell_address].border = Border(top=Side(style='thin', color='000000'),
                                                left=Side(style='thin', color='000000'),
                                                right=Side(style='thin', color='000000'),
                                                bottom=Side(style='thin', color='000000'))

        for col in ['C', 'E']:
            cell_address = f'{col}18'
            cover_ws[cell_address].border = Border(top=Side(style='thin', color='000000'), left=Side(style='thin', color='000000'),
                                                right=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))


        cover_ws['A21'].value = 'Group Field:'
        cover_ws['A21'].font = Font(size=11, italic=True, bold=True)

        cover_ws.merge_cells('D21:E21')
        cover_ws.merge_cells('H21:I21')

        cover_ws['D21'].value = 'Management and SHQE'
        cover_ws['D21'].font = Font(size=10)

        cover_ws['H21'].value = 'Tender Evaluation and Contract Negotiation'
        cover_ws['H21'].font = Font(size=10)

        cover_ws.merge_cells('D23:E23')
        cover_ws.merge_cells('H23:I23')

        cover_ws['D23'].value = 'Construction Supervision'
        cover_ws['D23'].font = Font(size=10)

        cover_ws['H23'].value = 'O&M'
        cover_ws['H23'].font = Font(size=10)

        for col in ['C', 'G']:
            cell_address = f'{col}21'
            cover_ws[cell_address].border = Border(top=Side(style='thin', color='000000'), left=Side(style='thin', color='000000'),
                                                right=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))
        for col in ['C', 'G']:
            cell_address = f'{col}23'
            cover_ws[cell_address].border = Border(top=Side(style='thin', color='000000'), left=Side(style='thin', color='000000'),
                                                right=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))

        first_day_of_month = datetime(current_year, current_month, 1)
        for i in range(1, 17):
            col_address = chr(ord('C') + (i - 1))  # Alternating columns C and G
            cell_address = f'{col_address}27'
            day_of_week = (first_day_of_month + timedelta(days=i - 1)).strftime("%a")  # Get the abbreviated day name
            cover_ws[cell_address].value = day_of_week

            col_address = chr(ord('C') + (i - 1))
            cell_address = f'{col_address}28'
            cover_ws[str(cell_address)].value = str((first_day_of_month + timedelta(days=i - 1)).day)

            col_address = chr(ord('C') + (i - 1))
            cell_address_activity_type = f'{col_address}29'

            # Replace 'activity_date' with the actual date for which you want to retrieve the activityType
            activity_date = first_day_of_month + timedelta(days=i - 1)
            activity_instance = Activity.objects.filter(user=user, activityDate=activity_date).first()

            if activity_instance:
                cover_ws[cell_address_activity_type].value = str(activity_instance.get_activity_type())
            else:
                cover_ws[cell_address_activity_type].value = None

        first_day_of_second_half = datetime(current_year, current_month, 17).date()

        # Write abbreviated weekdays in row 30 for the second half of the month
        for i in range(0, calendar.monthrange(current_year, current_month)[1] - 16):
            col_address = chr(ord('C') + i) + '30'
            day_of_week = (first_day_of_second_half + timedelta(days=i)).strftime("%a")  # Get the abbreviated day name
            cover_ws[col_address].value = day_of_week


        # Continue writing the numeric values for the remaining days in the row below (row 30)
        for i in range(17, calendar.monthrange(current_year, current_month)[1] + 1):
            col_address = chr(ord('C') + (i - 17))
            cell_address = f'{col_address}31'
            cover_ws[cell_address].value = (first_day_of_month + timedelta(days=i - 1)).day

        # Add the corresponding user activities in row 32
        for i in range(17, calendar.monthrange(current_year, current_month)[1] + 1):
            col_address_activity_type = chr(ord('C') + (i - 17)) + '32'

            activity_date = datetime(current_year, current_month, 17) + timedelta(days=i - 17)
            activity_instance = Activity.objects.filter(user=user, activityDate=activity_date).first()

            if activity_instance:
                cover_ws[col_address_activity_type].value = str(activity_instance.get_activity_type())
            else:
                cover_ws[col_address_activity_type].value = None

        for row in range(36, 38):
            for col_letter in ['C', 'D', 'F', 'G', 'I', 'J']:
                cell = cover_ws[col_letter + str(row)]
                cell.border = Border(
                    left=Side(style='thin', color='000000' if col_letter != 'C' else '000000'),
                    right=Side(style='thin', color='000000' if col_letter != 'J' else '000000'),
                    top=Side(style='thin', color='000000'),
                    bottom=Side(style='thin', color='000000')
                )
        # add summary for working days
        cover_ws.merge_cells('C36:D36')
        cover_ws['C36'].value = "No. of Days (NOD)*"
        cover_ws['C36'].font = Font(size=11, bold=True)
        cover_ws['C36'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text
        cover_ws['C36'].fill = PatternFill(start_color='C0C0C0', end_color='C0C0C0', fill_type='solid')

        start_date = current_date.replace(day=1)
        last_day_of_month = calendar.monthrange(current_date.year, current_date.month)[1]
        end_date = current_date.replace(day=last_day_of_month)
        total_working_days = np.busday_count(start_date, end_date, weekmask='0011111')
        cover_ws['F38'].value = total_working_days
        cover_ws['F38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        # Assuming user is an instance of the User model
        if user.expert == constants.LOCAL_USER:
            self.__add_local_working_days_individual__(current_date, user, cover_ws)
        elif user.expert == constants.EXPERT_USER:
            self.__add_expert_working_days_individual__(current_date, user, cover_ws)

        cover_ws.merge_cells('F36:G36')
        cover_ws['F36'].value = "Total Calendar Days (TCD)"
        cover_ws['F36'].font = Font(size=10, bold=True)
        cover_ws['F36'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text
        cover_ws['F36'].fill = PatternFill(start_color='C0C0C0', end_color='C0C0C0', fill_type='solid')

        cover_ws['C37'].value = "Japan"
        cover_ws['C37'].font = Font(size=11, bold=True)
        cover_ws['C37'].alignment = Alignment(horizontal='center', vertical='center')

        cover_ws['D37'].value = "Cairo"
        cover_ws['D37'].font = Font(size=11, bold=True)
        cover_ws['D37'].alignment = Alignment(horizontal='center', vertical='center')


        cover_ws['F37'].value = "Japan"
        cover_ws['F37'].font = Font(size=11, bold=True)
        cover_ws['F37'].alignment = Alignment(horizontal='center', vertical='center')

        cover_ws['G37'].value = "Cairo"
        cover_ws['G37'].font = Font(size=11, bold=True)
        cover_ws['G37'].alignment = Alignment(horizontal='center', vertical='center')

        cover_ws.merge_cells('I36:J36')
        cover_ws['I36'].value = "Consumption NOD/TCD"
        cover_ws['I36'].font = Font(size=11, bold=True)
        cover_ws['I36'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text
        cover_ws['I36'].fill = PatternFill(start_color='C0C0C0', end_color='C0C0C0', fill_type='solid')

        cover_ws['I37'].value = "Japan"
        cover_ws['I37'].font = Font(size=11, bold=True)
        cover_ws['I37'].alignment = Alignment(horizontal='center', vertical='center')

        cover_ws['J37'].value = "Cairo"
        cover_ws['J37'].font = Font(size=11, bold=True)
        cover_ws['J37'].alignment = Alignment(horizontal='center', vertical='center')

        # Project Director cell (B42)
        utilities.__format_cell__(cover_ws['B42'], "Project Director")

        # NAT Approval cell (L42)
        utilities.__format_cell__(cover_ws['L42'], "NAT Approval")

        for col_letter in range(ord('A'), ord('S')):
            col_letter = chr(col_letter)
            cell = cover_ws[col_letter + '45']
            cell.border = Border(bottom=Side(style='thick'))

        cover_ws.merge_cells('B47:D47')
        cover_ws['B47'].value = "J = Working day In Japan"
        cover_ws['B47'].font = Font(size=11, bold=True)

        cover_ws.merge_cells('M47:O47')
        cover_ws['M47'].value = "C = Working day In Cairo"
        cover_ws['M47'].font = Font(size=11, bold=True)

        cover_ws.merge_cells('B49:D49')
        cover_ws['B49'].value = "H = Official Holiday In Cairo"
        cover_ws['B49'].font = Font(size=11, bold=True)

        cover_ws.merge_cells('M49:O49')
        cover_ws['M49'].value = "X = Day off"
        cover_ws['M49'].font = Font(size=11, bold=True)

        cover_ws.merge_cells('B51:P51')
        cover_ws['B51'].value = "Note: According to the contract 81/M the total days are working days in Cairo plus to official holiday in Egypt *NOD=C (Working day in Cairo)+H (Official Holiday in Egypt)"


    @action(detail=False, methods=['GET'], url_path=r'extract_timesheet/(?P<userId>\w+(?:-\w+)*)', serializer_class=UserTimeSheetSerializer)
    def extract_timesheet(self, request, *args, **kwargs):
        if not utilities.check_is_admin(request.user.id):
            return Response(status=status.HTTP_403_FORBIDDEN)

        serializer = UserTimeSheetSerializer(data=request.query_params)
        serializer.is_valid()

        if not serializer.is_valid():
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

        # get the month and user id
        date = serializer.validated_data['date']
        user_id = kwargs['userId']

        # get all the activities for that month by that user
        activities = Activity.objects.filter(created__month=date.month, created__year=date.year, user__id=user_id).order_by('-created')
        user = User.objects.get(id=user_id)

        dateFont = Font(size=16)
        wb = Workbook()

        ws = wb.active
        ws.title = "User Timesheet"

        name_year_month_border = Border(
        left=Side(border_style='thin'),
        right=Side(border_style='thin'),
        top=Side(border_style='thin'),
        bottom=Side(border_style='thin')
    )

        # Write the month and year headers at the top
        ws.cell(row=1, column=1, value="Year")
        ws.cell(row=1, column=1).border = name_year_month_border

        ws.cell(row=1, column=2, value=date.year)
        ws.cell(row=1, column=2).border = name_year_month_border

        ws.cell(row=2, column=1, value="Month")
        ws.cell(row=2, column=1).border = name_year_month_border
        ws.cell(row=2, column=2, value=calendar.month_name[date.month])
        ws.cell(row=2, column=2).border = name_year_month_border
        ws.cell(row=1, column=1).font = dateFont  # Apply the font settings
        ws.cell(row=1, column=2).font = dateFont  # Apply the font settings
        ws.cell(row=2, column=1).font = dateFont  # Apply the font settings
        ws.cell(row=2, column=2).font = dateFont  # Apply the font settings

        ws.cell(row=1, column=7, value=user.get_full_name())
        ws.cell(row=1, column=7).font = dateFont
        ws.cell(row=1, column=7).border = name_year_month_border

        ws.cell(row=2, column=7, value=user.position)
        ws.cell(row=2, column=7).font = dateFont
        ws.cell(row=2, column=7).border = name_year_month_border

        # Create headers for columns
        ws.cell(row=4, column=1, value="Day").font = dateFont
        ws.cell(row=4, column=1).border = name_year_month_border
        ws.cell(row=4, column=2, value="Cairo").font = dateFont
        ws.cell(row=4, column=2).border = name_year_month_border
        ws.cell(row=4, column=3, value="Japan").font = dateFont
        ws.cell(row=4, column=3).border = name_year_month_border
        ws.cell(row=4, column=4, value="Daily Activities").font = dateFont
        ws.cell(row=4, column=4).border = name_year_month_border

        script_directory = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(script_directory, 'static', 'images', 'logo.png')
        img = Image(logo_path)
        ws.add_image(img, 'P1')

        for day in range(1, calendar.monthrange(date.year, date.month)[1] + 1):
            row_index = day + 4  # Offset by 4 to account for existing rows
            ws.cell(row=row_index, column=1, value=f"{day:02d}").font = dateFont
            ws.cell(row=row_index, column=1).alignment = Alignment(horizontal='center')
            activities_for_day = activities.filter(activityDate__day=day)
            # Display activities for the day in the second column
            activities_text = "\n".join([activity.userActivity or '' for activity in activities_for_day])

             # Merge cells for the "Activities" column
            start_column_letter = get_column_letter(4)  # Column D
            end_column_letter = get_column_letter(11)  # Adjust the last column letter as needed
            activities_column_range = f"{start_column_letter}{row_index}:{end_column_letter}{row_index}"

            # Set the width of the columns
            for col in ws.iter_cols(min_col=4, max_col=11, min_row=row_index, max_row=row_index):
                for cell in col:
                    ws.column_dimensions[cell.column_letter].width = 15  # Adjust the width as needed

            ws.merge_cells(activities_column_range)

            ws[start_column_letter + str(row_index)].alignment = Alignment(wrap_text=True)

            ws.cell(row=row_index, column=4, value=activities_text)

            activities_type = "\n".join([str(activity.get_activity_type()) for activity in activities_for_day])

            start_column_letter = get_column_letter(2)  # Column B
            end_column_letter = get_column_letter(2)

            if user.expert in [constants.LOCAL_USER, constants.EXPERT_USER]:
                ws.cell(row=row_index, column=3 if activities_type == "J" else 2, value=activities_type)

        # Get the current month and year
        current_month_name = date.strftime("%B")

        self.__add_individual_timesheet__(wb=wb, current_month_name=current_month_name,\
                                           current_year=date.year, user=user, current_date=date,\
                                            current_month=date.month)

        excel_data = BytesIO()
        wb.save(excel_data)
        excel_data.seek(0)
        response = HttpResponse(excel_data, content_type="application/ms-excel")
        response["Content-Disposition"] = f'attachment; filename=timesheet{date.year}_{date.month}_{user_id}.xlsx'
        return response

class ActivityViewSet(viewsets.ModelViewSet):
    permission_classes=(IsAuthenticated,)
    queryset = Activity.objects.all()
    def get_serializer_class(self):
        return ActivitySerializer

    @action(detail=False, methods=['GET'])
    def get_activities(self, request, *args, **kwargs):
        requestId = request.user.id
        user = User.objects.filter(id=requestId).first()
        date_param = request.query_params.get('date', None)

        if date_param:
        # If a date is provided, parse it and fetch activities for that date
            try:
                selected_date = datetime.strptime(date_param, '%Y-%m-%d').date()
            except ValueError:
                return Response({'detail': 'Invalid date format. Please use YYYY-MM-DD format.'}, status=status.HTTP_400_BAD_REQUEST)
        else:
            # If no date is provided, default to today's date
            selected_date = date.today()
        if user.is_superuser or user.isAdmin:
            activities = Activity.objects.all().order_by('-created')
        else:
            activities = Activity.objects.filter(user_id=requestId).order_by('-created')
        serializer = ActivitySerializer(activities, many=True)
        return Response(serializer.data, status=status.HTTP_200_OK)

    @action(detail=False, methods=['GET'])
    def my_activities(self, request, *args, **kwargs):
        """
            This returns the logged in user activities for the current month,
            however if the logged in user is a superuser/admin then it shows
            all users' activities
        """
        requestId = request.user.id
        user = User.objects.filter(id=requestId).first()

        serializer = UserTimeSheetSerializer(data=request.query_params)
        serializer.is_valid()
        if not serializer.is_valid():
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)
        # Get the first day and last day of the current month
        current_date = serializer.validated_data['date']

        first_day_of_month = current_date.replace(day=1)
        last_day_of_month = (first_day_of_month + timezone.timedelta(days=32)).replace(day=1) - timezone.timedelta(days=1)

        if (user.is_superuser or user.isAdmin):
            activities = Activity.objects.filter(activityDate__range=[first_day_of_month, last_day_of_month]).order_by('activityDate')
        else:
            # Filter activities for the current month and order by activityDate
            activities = Activity.objects.filter(user=user, activityDate__range=[first_day_of_month, last_day_of_month]).order_by('activityDate')
        serializer = ActivitySerializer(activities, many=True)
        return Response(serializer.data, status=status.HTTP_200_OK)

    def __format_cell__(self, cell, value, size=11, italic=False, center=True):
        cell.value = value
        cell.font = Font(size=size, italic=True)
        if center:
            cell.alignment = Alignment(horizontal='center', vertical='center')

    def __add_local_working_days__(self, current_date, user, cover_ws):
        ##### EXPERT #####
        start_date = current_date.date().replace(day=1)
        last_day_of_month = calendar.monthrange(current_date.date().year, current_date.date().month)[1]
        end_date = (current_date.date() + timedelta(days=last_day_of_month)).replace(day=1)

        # Adjust end_date to the next day
        end_date = end_date + timedelta(days=1)

        total_working_days_expert = np.busday_count(start_date, end_date, weekmask='0011111')
        cover_ws['F38'].value = total_working_days_expert
        cover_ws['F38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text
        ##### EXPERT #####

        start_date = current_date.date().replace(day=1)
        last_day_of_month = calendar.monthrange(current_date.date().year, current_date.date().month)[1]
        end_date = current_date.date().replace(day=last_day_of_month)

        last_day_of_month = calendar.monthrange(current_date.date().year, current_date.date().month)[1]
        end_date = current_date.date().replace(day=last_day_of_month) + timedelta(days=1)
        total_working_days = np.busday_count(start_date, end_date, weekmask='1111111')
        # Iterate through each day in the range
        for day in range(1, last_day_of_month + 1):
            day_off = Activity.objects.filter(
                user=user,
                activityDate__day=current_date.date().day,
                activityDate__month=current_date.date().month,
                activityDate__year=current_date.date().year,
                activityType=constants.OFFDAY,
            ).exists()

            if day_off and (day == 5 and day - 1 in (4, 6) and day + 1 in (4, 6)):
                total_working_days -= 1

            # Filter activities for the current user and month excluding 'H' type activities
            activities = Activity.objects.filter(
                user=user,
                activityDate__month=current_date.date().month,
                activityDate__year=current_date.date().year,
            ).exclude(activityType=constants.OFFDAY)

            # Count the number of activities
            working_days = activities.count()
            # Iterate over weeks in the month
            for week in calendar.monthcalendar(current_date.year, current_date.month):
                for day in week:
                    # Check if Thursday and Saturday are off-days in the current week
                    if day != 0:  # Ignore days that belong to the previous or next month (represented as 0)
                        thursday_offday = any(activity.activityType == constants.OFFDAY for activity in activities.filter(activityDate__day=day + 3))
                        saturday_offday = any(activity.activityType == constants.OFFDAY for activity in activities.filter(activityDate__day=day + 5))

                        # If either Thursday or Saturday is off-day, decrement working_days
                        if thursday_offday or saturday_offday:
                            working_days -= 1

        # Add "0" under "Cairo" for local users
        cover_ws['D38'].value = working_days
        cover_ws['D38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text
        cover_ws['G38'].value = total_working_days
        cover_ws['G38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        cover_ws['C38'].value = 0
        cover_ws['C38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        cover_ws['J38'].value = round(cover_ws['D38'].value / cover_ws['G38'].value, 3)
        cover_ws['J38'].font = Font(size=11)
        cover_ws['J38'].alignment = Alignment(horizontal='center', vertical='center')

         # Japan NOD/TCD
        cover_ws['I38'].value = round(cover_ws['C38'].value / cover_ws['F38'].value, 3)
        cover_ws['I38'].font = Font(size=11)
        cover_ws['I38'].alignment = Alignment(horizontal='center', vertical='center')

    def __add_expert_working_days__(self, current_date, user, cover_ws):
        ###### LOCAL ######
        start_date = current_date.date().replace(day=1)
        last_day_of_month = calendar.monthrange(current_date.date().year, current_date.date().month)[1]
        end_date = current_date.date().replace(day=last_day_of_month)

        last_day_of_month = calendar.monthrange(current_date.date().year, current_date.date().month)[1]
        end_date = current_date.date().replace(day=last_day_of_month) + timedelta(days=1)
        total_working_days_cairo = np.busday_count(start_date, end_date, weekmask='1111111')
         ###### LOCAL ######
        start_date = current_date.date().replace(day=1)
        last_day_of_month = calendar.monthrange(current_date.date().year, current_date.date().month)[1]
        end_date = (current_date.date() + timedelta(days=last_day_of_month)).replace(day=1)

        # Adjust end_date to the next day
        end_date = end_date + timedelta(days=1)
        total_working_days = np.busday_count(start_date, end_date, weekmask='0011111')

        # Filter activities for the current user and month excluding 'H' type activities
        activities = Activity.objects.filter(
            user=user,
            activityDate__month=current_date.date().month,
            activityDate__year=current_date.date().year,
        ).exclude(activityType=constants.OFFDAY)

        # Count the number of activities
        working_days = activities.count()

        cover_ws['C38'].value = working_days
        cover_ws['C38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        cover_ws['F38'].value = total_working_days
        cover_ws['F38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        # Cairo NOD
        cover_ws['D38'].value = 0
        cover_ws['D38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        # Cairo TCD
        cover_ws['G38'].value = total_working_days_cairo
        cover_ws['G38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        # Japan NOD/TCD
        cover_ws['I38'].value = round(cover_ws['C38'].value / cover_ws['F38'].value, 3)
        cover_ws['I38'].font = Font(size=11)
        cover_ws['I38'].alignment = Alignment(horizontal='center', vertical='center')

        # Cairo NOD/TCD
        cover_ws['J38'].value = 0
        cover_ws['J38'].font = Font(size=11)
        cover_ws['J38'].alignment = Alignment(horizontal='center', vertical='center')

    def set_borders(self, ws, row, columns):
        for col in columns:
            cell_address = f'{col}{row}'
            ws[cell_address].border = Border(top=Side(style='thin', color='000000'),
                                            left=Side(style='thin', color='000000'),
                                            right=Side(style='thin', color='000000'),
                                            bottom=Side(style='thin', color='000000'))

    @action(detail=False, methods=['GET'])
    def export_all(self, request, *args, **kwargs):
        user = request.user
        if not (user.is_superuser or user.isAdmin):
            return Response(constants.NOT_ALLOWED_TO_ACCESS, status=status.HTTP_400_BAD_REQUEST)
        
        serializer = UserTimeSheetSerializer(data=request.query_params)
        serializer.is_valid()

        if not serializer.is_valid():
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

        date = serializer.validated_data.get('date')

        company_param = request.query_params.get('company', None)
        if company_param:
            company_users = User.objects.filter(company=company_param)
            return generate_noce_timesheet(users=company_users, companyName=constants.COMPANY_CHOICES[int(company_param)][1], date=date)
        else:
            return generate_noce_timesheet(date=date)

        # return Response(status=status.HTTP_200_OK)

    @action(detail=False, methods=['POST'])
    def create_activity(self, request, *args, **kwargs):
        """
            This endpoint allows the user to create an activity
        """
        userId = request.user.id

        if User.objects.filter(id=userId).first().needsPasswordReset:
            return Response(constants.ERR_PASSWORD_RESET_NEEDED, status=status.HTTP_400_BAD_REQUEST)

        serializer = CreateActivitySerializer(data=request.data)
        serializer.is_valid()
        if not serializer.is_valid():
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

        userActivity = serializer.validated_data.get('userActivity', None)
        activityType = serializer.validated_data.get('activityType', None)
        activityDate = serializer.validated_data.get('activityDate', None)
        # Check if the user has already logged an activity today
        existing_activity = Activity.objects.filter(user__id=userId, activityDate=activityDate).first()

        if existing_activity:
            return Response({"detail": "You have already logged an activity for the selected date."}, status=status.HTTP_400_BAD_REQUEST)

        new_activity = Activity.objects.create(userActivity=userActivity,\
                                                activityType=activityType,\
                                                user_id=userId,
                                                activityDate=activityDate)
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

    @action(detail=False, methods=['GET'])
    def calculate_activity(self, request, *args, **kwargs):
        """
        This endpoint calculates how many activities a user does out of the possible working days
        EXPERT -> Saturday, Sunday off
        LOCAL -> Friday off
        """
        userId = request.user.id
        if not utilities.check_is_admin(userId):
            return Response(status=status.HTTP_403_FORBIDDEN)

        serializer = UserTimeSheetSerializer(data=request.query_params)
        serializer.is_valid()

        if not serializer.is_valid():
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

        date = serializer.validated_data['date']

        # Get all users
        users = User.objects.all()

        data = []

        for user in users:
            start_date = date.replace(day=1)
            last_day_of_month = calendar.monthrange(date.year, date.month)[1]
            end_date = date.replace(day=last_day_of_month)

            if user.expert == constants.EXPERT_USER:
                start_date = date.replace(day=1)
                end_date = date.replace(day=calendar.monthrange(date.year, date.month)[1])
                end_date = end_date + relativedelta(days=1)

                total_working_days_japan = np.busday_count(start_date, end_date, weekmask='0011111')

                # Filter activities for the current user and month excluding 'H' type activities
                activities = Activity.objects.filter(
                    user=user,
                    activityDate__month=date.month,
                    activityDate__year=date.year,
                    activityType__in=[constants.HOMEASSIGN]
                )

                # Count the number of activities
                working_days_japan = activities.count()
                last_day_of_month = calendar.monthrange(date.year, date.month)[1]
                end_date = date.replace(day=last_day_of_month) + timedelta(days=1)

                total_working_days_cairo = np.busday_count(start_date, end_date, weekmask='1111111')
                working_days_cairo = Activity.objects.filter(
                    user=user,
                    activityDate__month=date.month,
                    activityDate__year=date.year,
                ).filter(activityType__in=[constants.HOLIDAY, constants.INCAIRO]).count()
                data.append({
                    'user__id': user.id,
                    'user__first_name': user.first_name,
                    'user__last_name': user.last_name,
                    'user__email': user.email,
                    'working_days_cairo': working_days_cairo,
                    'total_days_cairo': total_working_days_cairo,
                    'working_days_japan': working_days_japan,
                    'total_days_japan': total_working_days_japan
                })

            if user.expert == constants.LOCAL_USER:
                # Get the last day of the month
                last_day_of_month = calendar.monthrange(date.year, date.month)[1]
                end_date = date.replace(day=last_day_of_month) + timedelta(days=1)
                total_working_days_cairo = np.busday_count(start_date, end_date, weekmask='1111111')
                total_working_days_japan = np.busday_count(start_date, end_date, weekmask='0011111')
                # Iterate through each day in the range

                # Filter activities for the current user and month excluding 'H' type activities
                activities_cairo = Activity.objects.filter(
                    user=user,
                    activityDate__month=date.month,
                    activityDate__year=date.year,
                    activityType__in=[constants.HOLIDAY, constants.INCAIRO]
                )

                # Count the number of activities
                working_days_cairo = activities_cairo.count()

                all_cairo_activities = Activity.objects.filter(
                    user=user,
                    activityDate__month=date.month,
                    activityDate__year=date.year,
                    activityType__in=[constants.HOLIDAY, constants.INCAIRO]
                )

                for week in calendar.monthcalendar(date.year, date.month):
                    for day in week:
                        # Check if Thursday and Saturday are off-days in the current week
                        if day != 0:  # Ignore days that belong to the previous or next month (represented as 0)
                            thursday_offday = any(activity.activityType == constants.OFFDAY for activity in all_cairo_activities.filter(activityDate__day=day + 3))
                            friday_offday = any(activity.activityType == constants.OFFDAY for activity in all_cairo_activities.filter(activityDate__day=day + 4))
                            saturday_offday = any(activity.activityType == constants.OFFDAY for activity in all_cairo_activities.filter(activityDate__day=day + 5))

                            # If either Thursday or Saturday is off-day, decrement working_days
                            if thursday_offday and saturday_offday and friday_offday:
                                total_working_days_cairo -= 1

                start_date = date.replace(day=1)
                end_date = date.replace(day=calendar.monthrange(date.year, date.month)[1])
                end_date = end_date + relativedelta(days=1)

                total_working_days_japan = np.busday_count(start_date, end_date, weekmask='0011111')

                activities_japan = Activity.objects.filter(
                    user=user,
                    activityDate__month=date.month,
                    activityDate__year=date.year,
                ).filter(activityType__in=[constants.HOMEASSIGN])
                working_days_japan = activities_japan.count()

                # Append user data to the response
                data.append({
                    'user__id': user.id,
                    'user__first_name': user.first_name,
                    'user__last_name': user.last_name,
                    'user__email': user.email,
                    'working_days_cairo': working_days_cairo,
                    'total_days_cairo': total_working_days_cairo,
                    'working_days_japan': working_days_japan,
                    'total_days_japan': total_working_days_japan
                })

        return Response(data, status=status.HTTP_200_OK)

    @action(detail=False, methods=['PATCH'], url_path=r'edit_activity/(?P<activityId>\w+(?:-\w+)*)')
    def edit_activity(self, request, *args, **kwargs):
        serializer = EditActivitySerializer(data=request.data)
        serializer.is_valid()
        if not serializer.is_valid():
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

        request_user_id = request.user.id
        activityId = self.kwargs['activityId']

        # Check if activity exists
        activity = Activity.objects.filter(id=activityId).first()
        if not activity:
            return Response(constants.ERR_NO_ACTIVITY_ID_FOUND, status=status.HTTP_400_BAD_REQUEST)

        user = User.objects.filter(id=request_user_id).first()

        # check if user editing is the owner of that activity
        if activity.user.id != request_user_id and not (user.isAdmin or user.is_superuser):
            return Response(constants.NOT_ALLOWED_TO_ACCESS, status=status.HTTP_400_BAD_REQUEST)


        # Update the activity fields with the serializer data
        activity.userActivity = serializer.validated_data.get('userActivity', activity.userActivity)
        activity.activityType = serializer.validated_data.get('activityType', activity.activityType)
        activity.activityDate = serializer.validated_data.get('activityDate', activity.activityDate)
        activity.save()

        return Response(status=status.HTTP_200_OK)

class LatestFileView(viewsets.ModelViewSet):
    @action(detail=False, methods=['GET'], url_path=r'activities/own_timesheet/(?P<userId>\w+(?:-\w+)*)')
    def export_own_timesheet(self, request, *args, **kwargs):
        serializer = UserTimeSheetSerializer(data=request.query_params)
        serializer.is_valid()

        if not serializer.is_valid():
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

        user_id = kwargs['userId']

        # Get the first day and last day of the current month
        date = serializer.validated_data['date']

        activities = Activity.objects.filter(created__month=date.month, created__year=date.year, user__id=user_id).order_by('-created')
        user = User.objects.get(id=user_id)

        dateFont = Font(size=16)
        wb = Workbook()

        ws = wb.active
        ws.title = "My Timesheet"

        name_year_month_border = Border(
        left=Side(border_style='thin'),
        right=Side(border_style='thin'),
        top=Side(border_style='thin'),
        bottom=Side(border_style='thin')
    )

        # Write the month and year headers at the top
        ws.cell(row=1, column=1, value="Year")
        ws.cell(row=1, column=1).border = name_year_month_border

        ws.cell(row=1, column=2, value=date.year)
        ws.cell(row=1, column=2).border = name_year_month_border

        ws.cell(row=2, column=1, value="Month")
        ws.cell(row=2, column=1).border = name_year_month_border
        ws.cell(row=2, column=2, value=calendar.month_name[date.month])
        ws.cell(row=2, column=2).border = name_year_month_border
        ws.cell(row=1, column=1).font = dateFont  # Apply the font settings
        ws.cell(row=1, column=2).font = dateFont  # Apply the font settings
        ws.cell(row=2, column=1).font = dateFont  # Apply the font settings
        ws.cell(row=2, column=2).font = dateFont  # Apply the font settings

        ws.cell(row=1, column=7, value=user.get_full_name())
        ws.cell(row=1, column=7).font = dateFont
        ws.cell(row=1, column=7).border = name_year_month_border

        ws.cell(row=2, column=7, value=user.department)
        ws.cell(row=2, column=7).font = dateFont
        ws.cell(row=2, column=7).border = name_year_month_border

        # Create headers for columns
        ws.cell(row=4, column=1, value="Day").font = dateFont
        ws.cell(row=4, column=1).border = name_year_month_border
        ws.cell(row=4, column=2, value="Cairo").font = dateFont
        ws.cell(row=4, column=2).border = name_year_month_border
        ws.cell(row=4, column=3, value="Japan").font = dateFont
        ws.cell(row=4, column=3).border = name_year_month_border
        ws.cell(row=4, column=4, value="Daily Activities").font = dateFont
        ws.cell(row=4, column=4).border = name_year_month_border

        script_directory = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(script_directory, 'static', 'images', 'logo.png')
        img = Image(logo_path)
        ws.add_image(img, 'P1')

        for day in range(1, calendar.monthrange(date.year, date.month)[1] + 1):
            row_index = day + 4  # Offset by 4 to account for existing rows
            ws.cell(row=row_index, column=1, value=f"{day:02d}").font = dateFont
            ws.cell(row=row_index, column=1).alignment = Alignment(horizontal='center')
            activities_for_day = activities.filter(activityDate__day=day)
            # Display activities for the day in the second column
            activities_text = "\n".join([activity.userActivity or '' for activity in activities_for_day])

             # Merge cells for the "Activities" column
            start_column_letter = get_column_letter(4)  # Column D
            end_column_letter = get_column_letter(11)  # Adjust the last column letter as needed
            activities_column_range = f"{start_column_letter}{row_index}:{end_column_letter}{row_index}"

            # Set the width of the columns
            for col in ws.iter_cols(min_col=4, max_col=11, min_row=row_index, max_row=row_index):
                for cell in col:
                    ws.column_dimensions[cell.column_letter].width = 15  # Adjust the width as needed

            ws.merge_cells(activities_column_range)

            ws[start_column_letter + str(row_index)].alignment = Alignment(wrap_text=True)

            ws.cell(row=row_index, column=4, value=activities_text)

            activities_type = "\n".join([str(activity.get_activity_type()) for activity in activities_for_day])

            start_column_letter = get_column_letter(2)  # Column B
            end_column_letter = get_column_letter(2)

            if user.expert in [constants.LOCAL_USER, constants.EXPERT_USER]:
                ws.cell(row=row_index, column=3 if activities_type == "J" else 2, value=activities_type)

        excel_data = BytesIO()
        wb.save(excel_data)
        excel_data.seek(0)
        response = HttpResponse(excel_data, content_type="application/ms-excel")
        response["Content-Disposition"] = f'attachment; filename=my_timesheet_{date.year}_{date.month}_{user.username}.xlsx'
        return response