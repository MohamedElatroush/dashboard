from .models import User, Activity
from rest_framework.response import Response
from rest_framework import status
from .serializers import CreateUserSerializer, ListUsersSerializer, UserDeleteSerializer,\
      ActivitySerializer, CreateActivitySerializer, MakeUserAdminSerializer, ChangePasswordSerializer
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
from openpyxl.styles import Alignment
from io import BytesIO
from django.http import HttpResponse
import pandas as pd
from .utilities import utilities
from openpyxl.styles import Font, PatternFill, Border, Side
from django.db import transaction
from openpyxl.drawing.image import Image
import os

class MyTokenObtainPairSerializer(TokenObtainPairSerializer):
    @classmethod
    def get_token(cls, user):
        token = super().get_token(user)
        token['username'] = user.username
        token['isAdmin'] = user.isAdmin
        token['is_superuser'] = user.is_superuser
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
        if not (user.is_superuser or user.isAdmin):
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
        is_admin = serializer.validated_data['isAdmin']

        modified_user = User.objects.filter(id=user_id).first()
        modified_user.isAdmin = False
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
            return Response(constants.CANT_RESET_USER_PASSWORD_ERROR, status=status.HTTP_400_BAD_REQUEST)
        userObject.password = make_password("1234")
        userObject.save()
        return Response(constants.SUCCESSFULLY_DELETED_USER, status=status.HTTP_200_OK)

    @transaction.atomic
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
            return Response({'detail': 'No file uploaded'}, status=status.HTTP_400_BAD_REQUEST)
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
                nat_group = int(row['NAT Group'])
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

    def _create_activity_excel_report(self, users, selected_date):
        current_date = datetime.now()
        current_month = current_date.month
        current_year = current_date.year
        font = Font(size=8)
        dateFont = Font(size=16)

        # Create a Border object for cell borders
        thin_border = Border(top=Side(style='thin'),
                            bottom=Side(style='thin'),
                            left=Side(style='medium'),  # Make the left border thick
                            right=Side(style='medium'),)

        # Create a fill (background color) for the headers
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # Get the current month and year
        current_month_name = current_date.strftime("%B")
        current_year = current_date.year

        last_day_of_month = (current_date.replace(day=1, month=current_month % 12 + 1, year=current_year) - timedelta(days=1)).day
        day_headers = [str(day) for day in range(1, last_day_of_month + 1)]

        constants.EXPORT_ACTIVITY_COLUMNS += day_headers
        addition_headers = ["Dep NAT", "Invoiced"]
        constants.EXPORT_ACTIVITY_COLUMNS += addition_headers

        # Create a new Excel workbook and add a worksheet
        wb = Workbook()
        ws = wb.active
        ws.title = "TS"

        # Write the month and year headers at the top
        ws.cell(row=1, column=1, value=current_month_name + " " + str(current_year))
        ws.cell(row=1, column=1).font = dateFont  # Apply the font settings
        ws.cell(row=1, column=1).fill = yellow_fill

        # Create a mapping dictionary for user types
        for col_num, column_title in enumerate(constants.EXPORT_ACTIVITY_COLUMNS, 1):
            cell = ws.cell(row=2, column=col_num)
            cell.value = column_title
            cell.alignment = Alignment(horizontal="center")
            cell.font = font if column_title.isnumeric() else None

        for row_num, user in enumerate(users, 4):
            ws.cell(row=row_num, column=1, value=row_num - 3).font = font  # Add userCounter
            ws.cell(row=row_num, column=2, value=user.get_full_name()).font = Font(size=8, bold=True)
            ws.cell(row=row_num, column=3, value=str(user.get_company())).font = font
            ws.cell(row=row_num, column=4, value=str(user.position)).font = font
            ws.cell(row=row_num, column=5, value=str(user.get_expert())).font = font
            ws.cell(row=row_num, column=6, value=str(user.get_grade())).font = font
            for col, day in enumerate(range(1, last_day_of_month + 1), start=8):
                activities_for_user_and_day = Activity.objects.filter(user=user, activityDate__day=day)
                activity_types = ", ".join(activity.get_activityType_display() for activity in activities_for_user_and_day if activity.activityType is not None)
                ws.cell(row=row_num, column=col, value=activity_types).font = font
                if "C" in activity_types:
                    green_fill = PatternFill(start_color="A8D08D", end_color="A8D08D", fill_type="solid")
                    ws.cell(row=row_num, column=col).fill = green_fill
                if "X" in activity_types:
                    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                    ws.cell(row=row_num, column=col).fill = red_fill
                if "H" in activity_types:
                    grey_fill = PatternFill(start_color="A6A6A6", end_color="A6A6A6", fill_type="solid")
                    ws.cell(row=row_num, column=col).fill = grey_fill
                if "J" in activity_types:
                    pink_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
                    ws.cell(row=row_num, column=col).fill = pink_fill
                # Set horizontal alignment to "center"
                cell.alignment = Alignment(horizontal="center")
            col = 8 + last_day_of_month
            ws.cell(row=row_num, column=col, value=str(user.get_natGroup())).font = font
            col += 1
            ws.cell(row=row_num, column=col, value="X").font = font

            for cell in ws[row_num]:
                cell.border = thin_border
             # Freeze the top row containing day headers
        ws.freeze_panes = ws.cell(row=4, column=8)

        for col_idx in range(8, 8 + last_day_of_month):
            column_letter = get_column_letter(col_idx)
            ws.column_dimensions[column_letter].width = 3

        # Create sheets for each Dep NAT
        dep_nats = list(set(user.get_natGroup() for user in users if user.get_natGroup() is not None))
        for dep_nat in dep_nats:
            if dep_nat is not None:
                sheet = wb.create_sheet(title=str(dep_nat))  # Create a new sheet for each Dep NAT
                sheet.sheet_properties.tabColor = "FFA500"
                self.__customize_sheet__(sheet, dep_nat, selected_date)

        excel_data = BytesIO()
        wb.save(excel_data)
        excel_data.seek(0)
        response = HttpResponse(excel_data, content_type="application/ms-excel")
        response["Content-Disposition"] = f'attachment; filename=activity_report_{current_month_name}_{current_year}.xlsx'
        return response

    def __customize_sheet__(self, sheet, dep_nat, selected_date):
        # Add a logo at the top right
        script_directory = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(script_directory, 'static', 'images', 'logo.png')
        img = Image(logo_path)
        sheet.add_image(img, 'H1')

        # Create a reverse mapping from dep_nat to the corresponding number
        nat_group_mapping_reverse = {name: number for number, name in constants.NAT_GROUP_CHOICES}
        # Get the number for the current dep_nat
        dep_nat_number = nat_group_mapping_reverse.get(dep_nat, -1)

        users = User.objects.filter(natGroup=dep_nat_number)

        # Set the title in the middle
        title_cell = sheet.cell(row=5, column=1, value=f"Dep NAT: {dep_nat}")
        title_cell.font = Font(size=16, bold=True)
        title_cell.alignment = Alignment(horizontal="center")

        # Merge cells for the title
        sheet.merge_cells(start_row=5, start_column=1, end_row=5, end_column=8)

        # Set the selected date below the title
        date_cell = sheet.cell(row=6, column=2, value=f"{selected_date.strftime('%B %Y')}")
        date_cell.font = Font(size=14)
        date_cell.alignment = Alignment(horizontal="center")

        # Merge cells for 'No.' and 'Name'
        sheet.merge_cells(start_row=8, start_column=1, end_row=8, end_column=4)
        sheet.merge_cells(start_row=8, start_column=5, end_row=8, end_column=8)

        # Set headers 'No.' and 'Name'
        sheet.cell(row=8, column=1, value='No.').font = Font(size=12, bold=True, color='000000')
        sheet.cell(row=8, column=5, value='Name').font = Font(size=12, bold=True, color='000000')

        # Add counter under 'No.' and extract names under 'Name'
        counter = 1
        for row_num, user in enumerate(users, start=9):
            sheet.cell(row=row_num, column=1, value=counter)
            sheet.cell(row=row_num, column=5, value=user.get_full_name())
            counter += 1

        sheet.freeze_panes = sheet.cell(row=9, column=1)

    @action(detail=False, methods=['GET'])
    def export_all(self, request, *args, **kwargs):
        adminId = request.user.id
        userObj = User.objects.filter(id=adminId).first()

        if not (userObj.is_superuser or userObj.isAdmin):
            print('here')
            return Response(constants.NOT_ALLOWED_TO_ACCESS, status=status.HTTP_400_BAD_REQUEST)

        # Get the date parameter from the query parameters
        date_param = request.query_params.get('date', None)

        if date_param:
            try:
                selected_date = datetime.strptime(date_param, '%Y-%m-%d').date()
            except ValueError:
                return Response({'detail': 'Invalid date format. Please use YYYY-MM-DD format.'}, status=status.HTTP_400_BAD_REQUEST)
        else:
            selected_date = date.today()
        users = User.objects.all()
        response = self._create_activity_excel_report(users, selected_date)
        return response

    @action(detail=False, methods=['POST'])
    def create_activity(self, request, *args, **kwargs):
        """
            This endpoint allows the user to create an activity
        """
        userId = request.user.id
        today = datetime.now().date()

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

