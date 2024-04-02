import uuid  
import pytz
from ..constants import constants
from openpyxl.styles import Font, PatternFill, Border, Side
from datetime import date, datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Alignment
import calendar
import os
from openpyxl.drawing.image import Image
from ..models import (Activity, User)
import numpy as np
from openpyxl.utils import get_column_letter
from collections import defaultdict
from io import BytesIO
from dateutil.relativedelta import relativedelta

def generate_unique_id():
    new_id = "EMP_" + str(uuid.uuid1())
    return new_id

# Helper function to convert and format a datetime to Cairo timezone
def convert_to_cairo_timezone_and_format(datetime_obj):
    cairo_timezone = pytz.timezone('Africa/Cairo')
    return datetime_obj.astimezone(cairo_timezone).strftime("%Y-%m-%d %H:%M:%S")

def format_hr_codes(userGrade):
    from dashboard.models import User
    users = User.objects.filter(grade=userGrade).order_by('hrCode')
    # Iterate over the users and update their HR codes
    new_hr_code = 1
    for user in users:
        user_hr_code = f'{constants.USER_GRADE_CHOICES[userGrade][1]}{new_hr_code:03d}'
        # Check if the HR code has changed
        if user.hrCode != user_hr_code:
            user.hrCode = user_hr_code
            user.save()

        new_hr_code += 1

def check_is_admin(id):
    from dashboard.models import User
    user = User.objects.filter(id=id).first()

    # if not super user dont show users
    if not (user.is_superuser or user.isAdmin):
        return False
    return True

def convert_grade_to_choice(grade):
    grade_mappings = {
        'A1': constants.GRADE_A_1,
        'A2': constants.GRADE_A_2,
        'A3': constants.GRADE_A_3,
        'B1': constants.GRADE_B_1,
        'B2': constants.GRADE_B_2,
        'B3': constants.GRADE_B_3,
        'B4': constants.GRADE_B_4,
        'B5': constants.GRADE_B_5
    }
    return grade_mappings.get(grade, None)

def convert_expert_to_choice(expert_string):
    expert_mappings = {
        'EXP': constants.EXPERT_USER,
        'LOC': constants.LOCAL_USER
    }
    return expert_mappings.get(expert_string, None)

def convert_nat_group_to_choice(nat_group_value):
    if isinstance(nat_group_value, str):
        nat_group_string = nat_group_value.strip().upper()
        nat_group_mappings = {
            'VCH': constants.VCH,
            'FUD': constants.FUD,
            'TD': constants.TD,
            'CWD': constants.CWD,
            'PSD': constants.PSD,
            'RSMD': constants.RSMD
        }
        return nat_group_mappings.get(nat_group_string, None)
    elif isinstance(nat_group_value, int):
        return nat_group_value
    else:
        return None  # Handle other cases if needed

def convert_company_to_choice(company_string):
    company_mappings = {
        'OCG': constants.OCG,
        'NK': constants.NK,
        'EHAF': constants.EHAF,
        'ACE': constants.ACE
    }
    return company_mappings.get(company_string, None)

def generate_username_from_name(name, taken_usernames):
    # Split the name and create a base username
    name_parts = name.split()
    base_username = f"{name_parts[0]}_{name_parts[-1]}".lower()

    # Initialize a counter
    counter = 0
    unique_username = base_username

    # Check if the username exists in the taken usernames set and find a unique one by appending a number
    while unique_username in taken_usernames:
        counter += 1
        unique_username = f"{base_username}_{counter}"

    # Add the new unique username to the taken usernames set
    taken_usernames.add(unique_username)

    return unique_username

def generate_first_name(name):
    if name:
        return name.split()[0]
    return None

def generate_last_name(name):
    if name:
        return ' '.join(name.split()[1:])
    return None

def __add_local_working_days__(current_date, user, cover_ws):
    ##### EXPERT #####
    start_date = current_date.date().replace(day=1)
    last_day_of_month = calendar.monthrange(current_date.date().year, current_date.date().month)[1]
    end_date = (current_date.date() + timedelta(days=last_day_of_month)).replace(day=1)

    # Adjust end_date to the next day
    end_date = end_date + timedelta(days=1)

    total_working_days_expert = np.busday_count(start_date, end_date, weekmask='0011111')
    cover_ws['F42'].value = total_working_days_expert
    cover_ws['F42'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text
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
        cover_ws['D42'].value = working_days
        cover_ws['D42'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text
        cover_ws['G42'].value = total_working_days
        cover_ws['G42'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        start_date = current_date.date().replace(day=1)
        end_date = current_date.date().replace(day=calendar.monthrange(current_date.date().year, current_date.date().month)[1])
        end_date = end_date + relativedelta(days=1)

        activities_japan = Activity.objects.filter(
            user=user,
            activityDate__month=current_date.date().month,
            activityDate__year=current_date.date().year,
        ).filter(activityType__in=[constants.HOMEASSIGN]).count()

        cover_ws['C42'].value = activities_japan
        cover_ws['C42'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        cover_ws['C42'].value = activities_japan
        cover_ws['C42'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text
        cover_ws['J42'].value = round(cover_ws['D42'].value / cover_ws['G42'].value, 3)
        cover_ws['J42'].font = Font(size=11)
        cover_ws['J42'].alignment = Alignment(horizontal='center', vertical='center')

         # Japan NOD/TCD
        cover_ws['I42'].value = round(cover_ws['C42'].value / cover_ws['F42'].value, 3)
        cover_ws['I42'].font = Font(size=11)
        cover_ws['I42'].alignment = Alignment(horizontal='center', vertical='center')

def __add_expert_working_days__(current_date, user, cover_ws):
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
        ).filter(activityType__in=[constants.HOMEASSIGN])

        # Count the number of activities
        working_days = activities.count()

        cover_ws['C42'].value = working_days
        cover_ws['C42'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        cover_ws['F42'].value = total_working_days
        cover_ws['F42'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        start_date = current_date.date().replace(day=1)
        end_date = current_date.date().replace(day=calendar.monthrange(current_date.date().year, current_date.date().month)[1])
        end_date = end_date + relativedelta(days=1)

        activities_cairo_count = Activity.objects.filter(
            user=user,
            activityDate__month=current_date.date().month,
            activityDate__year=current_date.date().year,
            activityType__in=[constants.INCAIRO, constants.HOLIDAY]
        ).count()

        # Cairo NOD
        cover_ws['D42'].value = activities_cairo_count
        cover_ws['D38'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        # Cairo TCD
        cover_ws['G42'].value = total_working_days_cairo
        cover_ws['G42'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text

        # Japan NOD/TCD
        cover_ws['I42'].value = round(cover_ws['C42'].value / cover_ws['F42'].value, 3)
        cover_ws['I42'].font = Font(size=11)
        cover_ws['I42'].alignment = Alignment(horizontal='center', vertical='center')

        # Cairo NOD/TCD
        cover_ws['J42'].value = round(cover_ws['D42'].value / cover_ws['G42'].value, 3)
        cover_ws['J42'].font = Font(size=11)
        cover_ws['J42'].alignment = Alignment(horizontal='center', vertical='center')

def set_borders(ws, row, columns):
    for col in columns:
        cell_address = f'{col}{row}'
        ws[cell_address].border = Border(top=Side(style='thin', color='000000'),
                                        left=Side(style='thin', color='000000'),
                                        right=Side(style='thin', color='000000'),
                                        bottom=Side(style='thin', color='000000'))

def __format_cell__(cell, value, size=11, italic=False, center=True):
    cell.value = value
    cell.font = Font(size=size, italic=True)
    if center:
        cell.alignment = Alignment(horizontal='center', vertical='center')

def __add_daily_activities_sheet__(wb, current_date, user):
    # Extract month and year
    month = current_date.month
    year = current_date.year

    # get all the activities for that month by that user
    activities = Activity.objects.filter(activityDate__month=month, activityDate__year=year, user__id=user.id).order_by('-created')
    user = User.objects.get(id=user.id)

    daily_activities = wb.create_sheet(title=f"{user.first_name} (DA)")

    dateFont = Font(size=16)

    name_year_month_border = Border(
    left=Side(border_style='thin'),
    right=Side(border_style='thin'),
    top=Side(border_style='thin'),
    bottom=Side(border_style='thin')
)

    # Merge cells for the new section
    daily_activities.merge_cells(start_row=2, start_column=2, end_row=5, end_column=10)
    for row in range(2, 6):
        for col in range(2, 11):
            cell = daily_activities.cell(row=row, column=col)
            cell.border = name_year_month_border

    # Set the text and alignment in the first cell of the merged range
    merged_cell = daily_activities.cell(row=2, column=2)
    merged_cell.value = "Consulting Services for Greater Cairo Metro Line No. 4 Phase1 Project"
    merged_cell.font = dateFont
    merged_cell.alignment = Alignment(horizontal='center', vertical='center')

    script_directory = os.path.dirname(os.path.abspath(__file__))
    parent_directory = os.path.dirname(script_directory)
    parent_parent_directory = os.path.dirname(parent_directory)
    logo_path = os.path.join(parent_parent_directory, 'static', 'images', 'logo.png')
    img = Image(logo_path)
    daily_activities.add_image(img, 'M2')

    # Set font and border for "Year:" label in cell A8
    year_label_cell = daily_activities.cell(row=8, column=1, value="Year:")
    year_label_cell.font = Font(bold=True, size=12)
    year_label_cell.border = name_year_month_border

    # Align "Year:" label horizontally and vertically
    year_label_cell.alignment = Alignment(horizontal='center', vertical='center')

    # Set font and border for the year value in cell B8
    year_value_cell = daily_activities.cell(row=8, column=2, value=year)
    year_value_cell.font = Font(bold=True, size=12)
    year_value_cell.border = name_year_month_border

    # Align year value horizontally and vertically
    year_value_cell.alignment = Alignment(horizontal='center', vertical='center')
    year_digits = year % 100
    month_name = calendar.month_name[month]

    # Combine the month name and the last two digits of the year
    formatted_month = f"{month_name}-{year_digits:02d}"

    cell = daily_activities.cell(row=9, column=2, value=formatted_month)
    cell.border = name_year_month_border
    cell.font = Font(bold=True, size=12)
    cell.alignment = Alignment(horizontal='center', vertical='center')  # Center the text horizontally and vertically

    # Set the text "Month" in cell A9
    cell_month = daily_activities.cell(row=9, column=1, value="Month")
    cell_month.border = name_year_month_border
    cell_month.font = Font(bold=True, size=12)
    cell_month.alignment = Alignment(horizontal='center', vertical='center')  # Center the text horizontally and vertically

    # Adjust the width of column A to fit the text
    daily_activities.column_dimensions['A'].width = 10
    daily_activities.column_dimensions['B'].width = 10 # Adjust the height as needed

    daily_activities.merge_cells(start_row=8, start_column=7, end_row=8, end_column=11)
    daily_activities.merge_cells(start_row=9, start_column=7, end_row=9, end_column=11)
    daily_activities.merge_cells(start_row=12, start_column=6, end_row=12, end_column=15)

    for row in range(12, 13):
        for col in range(6, 16):
            cell = daily_activities.cell(row=row, column=col)
            cell.border = name_year_month_border

    cell = daily_activities.cell(row=8, column=7, value=user.get_full_name())
    cell.font = Font(bold=True, size=12)

    # Set alignment
    cell.alignment = Alignment(horizontal='center', vertical='center')

    # Create border
    border = Border(left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin'))

    # Apply border to each cell in the merged range
    for row in daily_activities.iter_rows(min_row=8, max_row=8, min_col=7, max_col=11):
        for cell in row:
            cell.border = border

    cell = daily_activities.cell(row=9, column=7, value=user.position)
    cell.font = Font(bold=True, size=12)

    # Set alignment
    cell.alignment = Alignment(horizontal='center', vertical='center')

    # Create border
    border = Border(left=Side(style='thin'), 
                    right=Side(style='thin'), 
                    top=Side(style='thin'), 
                    bottom=Side(style='thin'))

    # Apply border to each cell in the merged range
    for row in daily_activities.iter_rows(min_row=9, max_row=9, min_col=7, max_col=11):
        for cell in row:
            cell.border = border

    # Create headers for columns
    headers_border = Border(left=Side(style='medium'), 
            right=Side(style='medium'), 
            top=Side(style='medium'), 
            bottom=Side(style='medium'))
    
    for col in range(1, 4):
        header_cell = daily_activities.cell(row=12, column=col, value=["Day", "Cairo", "Japan"][col - 1])
        header_cell.font = dateFont
        header_cell.border = headers_border
        header_cell.alignment = Alignment(textRotation=90, vertical='center')

    cell = daily_activities.cell(row=12, column=6, value="DAILY ACTIVITIES")
    cell.font = dateFont
    cell.border = name_year_month_border

    # Center the text horizontally and vertically
    cell.alignment = Alignment(horizontal='center', vertical='center')
    
    for day in range(1, calendar.monthrange(year, month)[1] + 1):
        start_row_index = 2 * day + 11  # Offset by 12 to account for existing rows]
        
        # Merge two rows for each day in columns 1, 2, and 3
        daily_activities.merge_cells(start_row=start_row_index, start_column=1, end_row=start_row_index + 1, end_column=1)
        daily_activities.merge_cells(start_row=start_row_index, start_column=2, end_row=start_row_index + 1, end_column=2)
        daily_activities.merge_cells(start_row=start_row_index, start_column=3, end_row=start_row_index + 1, end_column=3)
        
        # Set value for the first cell in the merged range (representing the day)
        cell = daily_activities.cell(row=start_row_index, column=1)
        cell.value = f"{day:02d}"
        cell.font = Font(size=10)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Set borders for each cell within the merged range
        for col in ['A', 'B', 'C']:
            for row in range(start_row_index, start_row_index + 2):
                cell_address = f'{col}{row}'
                daily_activities[cell_address].border = Border(top=Side(style='thin', color='000000'),
                                                                left=Side(style='thin', color='000000'),
                                                                right=Side(style='thin', color='000000'),
                                                                bottom=Side(style='thin', color='000000'))
        
        # Fetch activities for the current day
        activities_for_day = activities.filter(activityDate__day=day)

        # Combine activities' text and type
        activities_text = "\n".join([activity.userActivity or '' for activity in activities_for_day])
        activities_type = "\n".join([str(activity.get_activity_type()) for activity in activities_for_day])

        # Calculate the number of merged rows for activities
        num_merged_rows = activities_text.count('\n') + 1

        # Merge cells for the "Activities" column (two rows)
        start_column_letter = get_column_letter(6)  # Column D
        end_column_letter = get_column_letter(15)  # Adjust the last column letter as needed
        activities_column_range = f"{start_column_letter}{start_row_index}:{end_column_letter}{(start_row_index + 2)-1}"
        daily_activities.merge_cells(activities_column_range)

        for col in range(6, 16):
            bottom_cell_address = f"{get_column_letter(col)}{start_row_index + num_merged_rows}"
            daily_activities[bottom_cell_address].border = Border(bottom=Side(style='thin', color='000000'))

        # Set value and font for the cell in column 6 (representing activities)
        
        activities_cell = daily_activities.cell(row=start_row_index, column=6)
        activities_cell.value = activities_text
        activities_cell.font = Font(size=10)

        # Adjust row height to fit the content of the cell in column 6
        activities_cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center', indent=5)
        
        # Calculate the required row height based on the number of lines in the text
        font_size = 10
        line_height = 1.2 * font_size
        required_height = line_height * num_merged_rows

        # Adjust the row height for each merged row
        for i in range(start_row_index, start_row_index + num_merged_rows):
            row_dimension = daily_activities.row_dimensions[i]
            row_dimension.height = required_height

        # Set the activities type in the appropriate column
        if user.expert in [constants.LOCAL_USER, constants.EXPERT_USER]:
            cell = daily_activities.cell(row=start_row_index, column=3 if activities_type == "J" else 2)
            cell.value = activities_type
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.font = Font(size=11)

def __add_cover_sheet__(wb, current_month_name, current_year, user, current_date, current_month):
    # Create cover page
    start_date = current_date.replace(day=1)
    last_day_of_month = calendar.monthrange(current_date.year, current_date.month)[1]
    end_date = current_date.replace(day=last_day_of_month)

    cover_ws = wb.create_sheet(title=str(user.first_name))

    # Merge and set the title
    border_style = Border(
        bottom=Side(border_style='thin')
    )

    # add the logo
    script_directory = os.path.dirname(os.path.abspath(__file__))
    parent_directory = os.path.dirname(script_directory)
    parent_parent_directory = os.path.dirname(parent_directory)
    logo_path = os.path.join(parent_parent_directory, 'static', 'images', 'logo.png')
    img = Image(logo_path)
    cover_ws.add_image(img, 'O2')


    title_cell = cover_ws['A1']
    cover_ws.merge_cells(start_row=1, start_column=1, end_row=6, end_column=13)
    title_cell.value = constants.COVER_TS_TEXT
    title_cell.font = Font(size=12)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')

    # Merge and set the Monthly Time Sheet text
    cover_ws.merge_cells(start_row=8, start_column=1, end_row=8, end_column=13)  # Modified here
    monthly_ts_cell = cover_ws['A8']  # Modified here
    monthly_ts_cell.value = 'Monthly Time Sheet'
    monthly_ts_cell.font = Font(size=16, italic=True, bold=True)
    monthly_ts_cell.alignment = Alignment(horizontal='center', vertical='center')

    # Set the month and year
    year_label_cell = cover_ws.cell(row=10, column=1, value="Month:")
    year_label_cell.font = Font(bold=True, size=12)

    year_label_cell = cover_ws.cell(row=10, column=11, value="Year:")
    year_label_cell.font = Font(bold=True, size=12)

    year_label_cell = cover_ws.cell(row=10, column=13, value=current_year)
    year_label_cell.font = Font(bold=True, size=12)
    year_label_cell.border = border_style

    value_year_cell = cover_ws.cell(row=10, column=3, value=current_month_name)
    value_year_cell.font = Font(bold=True, size=12)
    value_year_cell.border = border_style

    name_label_cell = cover_ws.cell(row=12, column=1, value="Name:")
    name_label_cell.font = Font(bold=True, size=12)

    name_label_cell = cover_ws.cell(row=12, column=3, value=str(user.get_full_name()))
    name_label_cell.font = Font(bold=True, size=12)
    name_label_cell.border = border_style

    user_grade = user.get_grade()
    name_label_cell = cover_ws.cell(row=14, column=1, value="Category:")
    name_label_cell.font = Font(bold=True, size=12)

    cover_ws.cell(row=14, column=4, value="A1")
    cover_ws.cell(row=14, column=7, value="A2")
    cover_ws.cell(row=14, column=10, value="A3")

    cover_ws.cell(row=16, column=4, value="B1")
    cover_ws.cell(row=16, column=7, value="B2")
    cover_ws.cell(row=16, column=10, value="B3")

    cover_ws.cell(row=18, column=4, value="B4")
    cover_ws.cell(row=18, column=7, value="B5")

    # Define the rows and columns
    rows = [14, 16, 18]
    columns = [4, 7, 10]

    # Iterate over each cell and set the value and alignment
    for row in rows:
        for col in columns:
            cell = cover_ws.cell(row=row, column=col)
            cell.alignment = Alignment(horizontal='center', vertical='center')


    # Mapping of grades to corresponding columns
    grade_mapping = {
        'A1': 'C14',
        'A2': 'F14',
        'A3': 'I14',
        'B1': 'C16',
        'B2': 'F16',
        'B3': 'I16',
        'B4': 'C18',
        'B5': 'F18',
    }

    # # Set the letter '/' in the corresponding cell based on the user's grade
    if user_grade in grade_mapping:
        grade_column = grade_mapping[user_grade]
        cell_address = grade_column
        cover_ws[cell_address].value = '/'
        cover_ws[cell_address].alignment = Alignment(horizontal='center', vertical='center')

    # Set borders for various cells
    set_borders(cover_ws, 14, ['C', 'F', 'I'])
    set_borders(cover_ws, 16, ['C', 'F', 'I'])
    set_borders(cover_ws, 18, ['C', 'F'])


    cell = cover_ws.cell(row=21, column=1, value="Nationality:")
    cell.font = Font(bold=True, size=12)

    cell = cover_ws.cell(row=21, column=4, value="Expatriate")
    cell.font = Font(name='Times New Roman', size=11)
    cell.alignment = Alignment(horizontal='center', vertical='center')

    cell = cover_ws.cell(row=21, column=7, value="Local")
    cell.font = Font(name='Times New Roman', size=11)
    cell.alignment = Alignment(horizontal='center', vertical='center')

    # Get the user's nationality
    user_nationality = user.get_expert()

    # Mapping of nationalities to corresponding columns
    nationality_mapping = {
        'EXP': 'C',  # Column for Expatriate
        'LOC': 'F',  # Column for Local
    }

    # Set the letter '/' in the corresponding cell and center it
    if user_nationality in nationality_mapping:
        nationality_column = nationality_mapping[user_nationality]
        cell_address = f'{nationality_column}21'
        cover_ws[cell_address].value = '/'
        cover_ws[cell_address].alignment = Alignment(horizontal='center', vertical='center')

    set_borders(cover_ws, 21, ['C', 'F'])

    cell = cover_ws.cell(row=24, column=1, value="Group Field:")
    cell.font = Font(bold=True, size=12)

    cover_ws.merge_cells(start_row=24, start_column=4, end_row=24, end_column=6)
    cover_ws.merge_cells(start_row=24, start_column=8, end_row=24, end_column=10)
    cover_ws.merge_cells(start_row=26, start_column=4, end_row=26, end_column=6)
    cover_ws.merge_cells(start_row=26, start_column=8, end_row=26, end_column=10)

    cell = cover_ws.cell(row=24, column=4, value="Management and SHQE")
    cell.font = Font(size=11)
    cell.alignment = Alignment(horizontal='center', vertical='center')

    cell = cover_ws.cell(row=24, column=8, value="Tender Evaluation and Contract Negotiation")
    cell.font = Font(size=11)
    cell.alignment = Alignment(horizontal='center', vertical='center')

    cell = cover_ws.cell(row=26, column=4, value="Construction Supervision")
    cell.font = Font(size=11)
    cell.alignment = Alignment(horizontal='center', vertical='center')

    cell = cover_ws.cell(row=26, column=8, value="O&M")
    cell.font = Font(size=11)
    cell.alignment = Alignment(horizontal='center', vertical='center')

    set_borders(cover_ws, 24, ['C', 'G'])
    set_borders(cover_ws, 26, ['C', 'G'])

    grey_fill = PatternFill(start_color='DDDDDD', end_color='DDDDDD', fill_type='solid')
    first_day_of_month = datetime(current_year, current_month, 1)
    for i in range(1, 17):
        col_address = chr(ord('A') + (i - 1))  # Alternating columns C and G
        # Set the day of the week
        day_of_week_cell = cover_ws[col_address + '31']
        day_of_week_cell.value = (first_day_of_month + timedelta(days=i - 1)).strftime("%a")
        day_of_week_cell.alignment = Alignment(horizontal='center', vertical='center')
        day_of_week_cell.fill = grey_fill
        set_borders(cover_ws, 31, [col_address])

        # Set the day of the month
        day_of_month_cell = cover_ws[col_address + '32']
        day_of_month_cell.value = str((first_day_of_month + timedelta(days=i - 1)).day)
        day_of_month_cell.alignment = Alignment(horizontal='center', vertical='center')
        day_of_month_cell.fill = grey_fill
        set_borders(cover_ws, 32, [col_address])
        cell_address_activity_type = f'{col_address}33'

        # Replace 'activity_date' with the actual date for which you want to retrieve the activityType
        activity_date = first_day_of_month + timedelta(days=i - 1)
        activity_instance = Activity.objects.filter(user=user, activityDate=activity_date).first()

        if activity_instance:
            activity_type = activity_instance.get_activity_type()
            cover_ws[cell_address_activity_type].value = str(activity_type)
            if activity_type == constants.ACTIVITY_TYPES_CHOICES[constants.HOLIDAY][1]:
                cover_ws[cell_address_activity_type].fill = grey_fill
        else:
            cover_ws[cell_address_activity_type].value = None
        cover_ws[cell_address_activity_type].alignment = Alignment(horizontal='center', vertical='center')

    first_day_of_second_half = datetime(current_year, current_month, 17).date()

    # Write abbreviated weekdays in row 30 for the second half of the month
    for i in range(0, calendar.monthrange(current_year, current_month)[1] - 16):
        col_address = chr(ord('A') + i) + '34'
        day_of_week = (first_day_of_second_half + timedelta(days=i)).strftime("%a")  # Get the abbreviated day name
        cell = cover_ws[col_address]
        cell.value = day_of_week
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.fill = grey_fill
        set_borders(cover_ws, 34, [col_address])

    # Continue writing the numeric values for the remaining days in the row below (row 30)
    for i in range(17, calendar.monthrange(current_year, current_month)[1] + 1):
        col_address = chr(ord('A') + (i - 17))
        cell_address = f'{col_address}35'
        cell = cover_ws[cell_address]
        cell.value = (first_day_of_month + timedelta(days=i - 1)).day
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.fill = grey_fill
        set_borders(cover_ws, 35, [col_address])


    for i in range(17, calendar.monthrange(current_year, current_month)[1] + 1):
        col_address_activity_type = chr(ord('A') + (i - 17)) + '36'

        activity_date = datetime(current_year, current_month, 17) + timedelta(days=i - 17)
        activity_instance = Activity.objects.filter(user=user, activityDate=activity_date).first()

        cell = cover_ws[col_address_activity_type]
        if activity_instance:
            activity_type = activity_instance.get_activity_type()
            cell.value = str(activity_type)
            if activity_type == constants.ACTIVITY_TYPES_CHOICES[constants.HOLIDAY][1]:
                cell.fill = grey_fill
        else:
            cell.value = None
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for row in range(40, 43):
        for col_letter in ['C', 'D', 'F', 'G', 'I', 'J']:
            cell = cover_ws[col_letter + str(row)]
            cell.border = Border(
                left=Side(style='thin', color='000000' if col_letter != 'C' else '000000'),
                right=Side(style='thin', color='000000' if col_letter != 'J' else '000000'),
                top=Side(style='thin', color='000000'),
                bottom=Side(style='thin', color='000000')
            )
    # add summary for working days
    cover_ws.merge_cells('C40:D40')
    cover_ws['C40'].value = "No. of Days (NOD)*"
    cover_ws['C40'].font = Font(size=11, bold=True)
    cover_ws['C40'].alignment = Alignment(horizontal='center', vertical='center')  # Center the text
    cover_ws['C40'].fill = grey_fill

    start_date = current_date.date().replace(day=1)
    last_day_of_month = calendar.monthrange(current_date.date().year, current_date.date().month)[1]
    end_date = current_date.date().replace(day=last_day_of_month)
    total_working_days = np.busday_count(start_date, end_date, weekmask='0011111')

    # Assuming user is an instance of the User model
    if user.expert == constants.LOCAL_USER:
        __add_local_working_days__(current_date, user, cover_ws)
    elif user.expert == constants.EXPERT_USER:
        __add_expert_working_days__(current_date, user, cover_ws)

    # Merge cells and set values for total calendar days (TCD)
    cover_ws.merge_cells('F40:G40')
    tcd_cell = cover_ws['F40']
    tcd_cell.value = "Total Calendar Days (TCD)"
    tcd_cell.font = Font(size=10, bold=True)
    tcd_cell.alignment = Alignment(horizontal='center', vertical='center')
    tcd_cell.fill = grey_fill

    # Set values for Japan and Cairo columns
    labels = {
        'C41': 'Japan',
        'D41': 'Cairo',
        'F41': 'Japan',
        'G41': 'Cairo'
    }
    for cell_address, label_text in labels.items():
        cell = cover_ws[cell_address]
        cell.value = label_text
        cell.font = Font(size=11, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Merge cells and set values for consumption NOD/TCD
    cover_ws.merge_cells('I40:J40')
    consumption_cell = cover_ws['I40']
    consumption_cell.value = "Consumption NOD/TCD"
    consumption_cell.font = Font(size=11, bold=True)
    consumption_cell.alignment = Alignment(horizontal='center', vertical='center')
    consumption_cell.fill = grey_fill

    # Set values for Japan and Cairo columns in consumption NOD/TCD section
    labels = {
        'I41': 'Japan',
        'J41': 'Cairo'
    }
    for cell_address, label_text in labels.items():
        cell = cover_ws[cell_address]
        cell.value = label_text
        cell.font = Font(size=11, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')


    # Project Director cell (B42)
    __format_cell__(cover_ws['B45'], "Project Director")

    # NAT Approval cell (L42)
    __format_cell__(cover_ws['L45'], "NAT Approval")

    for col_letter in range(ord('A'), ord('S')):
        col_letter = chr(col_letter)
        cell = cover_ws[col_letter + '47']
        cell.border = Border(bottom=Side(style='thick'))

    labels = {
        'B50': "J = Working day In Japan",
        'M50': "C = Working day In Cairo",
        'B52': "H = Official Holiday In Cairo",
        'M52': "X = Day off",
        'B55': "Note: According to the contract 81/M the total days are working days in Cairo plus to official holiday in Egypt *NOD=C (Working day in Cairo)+H (Official Holiday in Egypt)"
    }

    for cell_address, label_text in labels.items():
        cover_ws.merge_cells(f'{cell_address}:{chr(ord(cell_address[0]) + 2)}{cell_address[1:]}')
        cell = cover_ws[cell_address]
        cell.value = label_text
        cell.font = Font(size=11, bold=True)


def create_activity_excel_report(users, activities, selected_date, companyName, date):
    date = datetime.combine(date, datetime.min.time())

    current_date = date
    current_month = current_date.month
    current_year = current_date.year

    font = Font(size=8)

    # Create a Border object for cell borders
    thin_border = Border(top=Side(style='thin'),
                        bottom=Side(style='thin'),
                        left=Side(style='medium'),  # Make the left border thick
                        right=Side(style='medium'),)
    # Create a fill (background color) for the headers
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Get the current month and year
    current_month_name = current_date.strftime("%B")

    last_day_of_month = (current_date.replace(day=1, month=current_month % 12 + 1, year=current_year) - timedelta(days=1)).day
    day_headers = [str(day) for day in range(1, last_day_of_month + 1)]

    constants.EXPORT_ACTIVITY_COLUMNS += ["Cairo", "Japan", "Cairo %", "Japan %"]
    constants.EXPORT_ACTIVITY_COLUMNS += day_headers
    addition_headers = ["Dep NAT", "Invoiced"]
    constants.EXPORT_ACTIVITY_COLUMNS += addition_headers

    # Create a new Excel workbook and add a worksheet
    wb = Workbook()

    # Create individual timesheet/daily activities for every user
    for user in users:
        __add_cover_sheet__(wb, current_month_name, current_year, user, current_date, current_month)
        __add_daily_activities_sheet__(wb, current_date, user)

    ws = wb.active
    ws.title = "TS"

    ws.cell(row=1, column=1, value="Cairo Line 4 Phase 1 - TimeSheet")
    # Set the font to bold
    bold_font = Font(bold=True)
    ws.cell(row=1, column=1).font = bold_font
    # Combine the month abbreviation and year in the desired format
    formatted_date = f"{current_month_name.lower()[:3]}-{current_year}"
    ws.cell(row=1, column=5, value=formatted_date)
    yearFont = Font(size=8)
    ws.cell(row=1, column=5).font = yearFont
    ws.cell(row=1, column=5).fill = yellow_fill

    # Create a mapping dictionary for user types
    for col_num, column_title in enumerate(constants.EXPORT_ACTIVITY_COLUMNS, 1):
        cell = ws.cell(row=2, column=col_num)
        cell.value = column_title
        cell.alignment = Alignment(horizontal="center")
        cell.font = font if column_title.isnumeric() else None
        # Apply thick border above the headers
        cell.border = Border(outline=Side(style='medium'))

    for col_num in range(1, len(constants.EXPORT_ACTIVITY_COLUMNS) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.border = Border(bottom=Side(style='medium'))
        cell = ws.cell(row=3, column=col_num)
        cell.border = Border(bottom=Side(style='medium'))

    # Initialize a counter variable for rows
    current_row = 2
    # Use a defaultdict for activities
    default_activities = {day: '' for day in range(1, last_day_of_month + 1)}

    user_rows = {}
    for user in users:
        user_id = str(user.id)
        user_row = user_rows[user_id] = {
        'user_counter': current_row,
        'full_name': user.get_full_name(),
        'company': user.get_company(),
        'position': user.position,
        'expert': user.get_expert(),
        'grade': user.get_grade(),
        'nat_group': user.get_natGroup(),
        'invoiced': 'X',
        'Cairo': '',
        'Japan': '',
        'Cairo %': '',
        'Japan %': '',
        'activities': defaultdict(str, default_activities.copy())
    }
        # Retrieve activities for the current user
        user_activities = activities.filter(user=user)
        for activity in user_activities:
            day = activity.activityDate.day
            activity_type = activity.get_activity_type()
            user_row['activities'][day] = activity_type

        current_row += 1

    for user_id, user_data in user_rows.items():
        row_num = user_data['user_counter'] + 2
        ws.cell(row=row_num, column=1, value=row_num - 3).font = font  # Add userCounter
        ws.cell(row=row_num, column=2, value=user_data['full_name']).font = Font(size=8, bold=True)
        ws.cell(row=row_num, column=3, value=str(user_data['company'])).font = font
        ws.cell(row=row_num, column=4, value=str(user_data['position'])).font = font
        ws.cell(row=row_num, column=5, value=str(user_data['expert'])).font = font
        ws.cell(row=row_num, column=6, value=str(user_data['grade'])).font = font
        ws.cell(row=row_num, column=7 + last_day_of_month + 5, value=str(user_data['nat_group'])).font = font
        ws.cell(row=row_num, column=7 + last_day_of_month + 6, value=str(user_data['invoiced'])).font = font

        cairo_count = 0
        total_working_days_cairo = 0
        japan_count = 0
        total_working_days_japan = 0
        if user_data['expert'] == constants.EXPERT_LOCAL_CHOICES[constants.EXPERT_USER][1]:
            start_date = current_date.replace(day=1)
            end_date = current_date.replace(day=calendar.monthrange(current_date.year, current_date.month)[1])
            end_date = end_date + relativedelta(days=1)
            # Convert start_date and end_date to datetime64[D]
            start_date = np.datetime64(start_date, 'D')
            end_date = np.datetime64(end_date, 'D')

            total_working_days_cairo = np.busday_count(start_date, end_date, weekmask='1111111')
            total_working_days_japan = np.busday_count(start_date, end_date, weekmask='0011111')

        elif user_data['expert'] == constants.EXPERT_LOCAL_CHOICES[constants.LOCAL_USER][1]:
            start_date = current_date.replace(day=1)
            last_day_of_month = calendar.monthrange(current_date.year, current_date.month)[1]
            end_date = current_date.replace(day=last_day_of_month)
            last_day_of_month = calendar.monthrange(current_date.year, current_date.month)[1]
            end_date = current_date.replace(day=last_day_of_month) + timedelta(days=1)
            start_date = np.datetime64(start_date, 'D')
            end_date = np.datetime64(end_date, 'D')
            total_working_days_cairo = np.busday_count(start_date, end_date, weekmask='1111111')
            total_working_days_japan = np.busday_count(start_date, end_date, weekmask='0011111')

        for col, activity_type in user_data['activities'].items():
            if (activity_type == 'J'):
                japan_count += 1
            if activity_type in ['C', 'H']:
                cairo_count += 1

            ws.cell(row=row_num, column=col + 11, value=str(activity_type)).font = font
            if "C" in activity_type:
                green_fill = PatternFill(start_color="A8D08D", end_color="A8D08D", fill_type="solid")
                ws.cell(row=row_num, column=col + 11).fill = green_fill
            if "X" in activity_type:
                red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                ws.cell(row=row_num, column=col + 11).fill = red_fill
            if "H" in activity_type:
                grey_fill = PatternFill(start_color="A6A6A6", end_color="A6A6A6", fill_type="solid")
                ws.cell(row=row_num, column=col + 11).fill = grey_fill
            if "J" in activity_type:
                pink_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
                ws.cell(row=row_num, column=col + 11).fill = pink_fill

        # Check if total_working_days_cairo is not zero before calculating the percentage
        cairo_percentage = round(cairo_count / total_working_days_cairo, 5) if total_working_days_cairo != 0 else 0

        # Check if total_working_days_japan is not zero before calculating the percentage
        japan_percentage = round(japan_count / total_working_days_japan, 5) if total_working_days_japan != 0 else 0

        # worked days
        red_bold_italic_font = Font(size=12, color="FF0000", bold=True, italic=True)
        ws.cell(row=row_num, column=7 + 1, value=cairo_count).font = red_bold_italic_font
        ws.cell(row=row_num, column=7 + 2, value=japan_count).font = red_bold_italic_font

        ws.cell(row=row_num, column=7 + 3, value=cairo_percentage).font = red_bold_italic_font
        ws.cell(row=row_num, column=7 + 4, value=japan_percentage).font = red_bold_italic_font

        for cell in ws[row_num]:
            cell.border = thin_border
        ws.freeze_panes = ws.cell(row=4, column=8)

    for col in ws.columns:
        max_length = 0
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        ws.column_dimensions[col[0].column_letter].width = max_length

    # Center align all cells in column #
    for cell in ws['A']:
        cell.alignment = Alignment(horizontal='center')

    # Create AutoFilter for all columns starting from row 2
    ws.auto_filter.ref = f'A2:{get_column_letter(ws.max_column)}{ws.max_row}'

    for col_idx in range(12, 12 + last_day_of_month):
        column_letter = get_column_letter(col_idx)
        ws.column_dimensions[column_letter].width = 3

    # Create sheets for each Dep NAT
    dep_nats = list(set(user.get_natGroup() for user in users if user.get_natGroup() is not None))
    for dep_nat in dep_nats:
        if dep_nat is not None:
            sheet = wb.create_sheet(title=str(dep_nat))  # Create a new sheet for each Dep NAT
            sheet.sheet_properties.tabColor = "FFA500"
            __customize_sheet__(sheet, dep_nat, selected_date)
            # Write "Dep Nat" values under the appropriate column
    for user_id, user_data in user_rows.items():
        if user_data['nat_group'] == dep_nat:
            row_num = user_data['user_counter'] + 2
            dep_nat_cell = sheet.cell(row=row_num, column=constants.EXPORT_ACTIVITY_COLUMNS.index("Dep NAT") + 1)
            dep_nat_cell.value = str(dep_nat)
            dep_nat_cell.font = Font(size=8)
            dep_nat_cell.alignment = Alignment(horizontal="center")

    excel_data = BytesIO()
    wb.save(excel_data)
    excel_data.seek(0)

    excel_data = BytesIO()
    wb.save(excel_data)
    excel_data.seek(0)
    from django.http import HttpResponse
    response = HttpResponse(excel_data, content_type="application/ms-excel")
    response["Content-Disposition"] = f'attachment; filename=timesheet.xlsx'
    return response

def __department_mapping__(dep_nat):
    group = constants.DEP_MAPPING.get(dep_nat, 'Unknown')
    return group

def __customize_sheet__(sheet, dep_nat, selected_date):
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.sheet_view.showGridLines = False

    # Add a logo at the top right
    script_directory = os.path.dirname(os.path.abspath(__file__))
    parent_directory = os.path.dirname(script_directory)
    parent_parent_directory = os.path.dirname(parent_directory)
    logo_path = os.path.join(parent_parent_directory, 'static', 'images', 'logo.png')
    img = Image(logo_path)
    sheet.add_image(img, 'H1')

    # Create a reverse mapping from dep_nat to the corresponding number
    nat_group_mapping_reverse = {name: number for number, name in constants.NAT_GROUP_CHOICES}
    # Get the number for the current dep_nat
    dep_nat_number = nat_group_mapping_reverse.get(dep_nat, -1)

    department = __department_mapping__(dep_nat_number)

    users = User.objects.filter(natGroup=dep_nat_number)

    # Set the title in the middle
    title_cell = sheet.cell(row=5, column=1, value=f"Dep NAT: {department}")
    title_cell.font = Font(size=16, bold=True)
    title_cell.alignment = Alignment(horizontal="center")

    # Merge cells for the title
    sheet.merge_cells(start_row=5, start_column=1, end_row=5, end_column=8)
    selected_date = datetime.strptime(selected_date, '%Y-%m-%d')
    month_name = calendar.month_name[selected_date.month]
    # Set the selected date below the title
    date_cell = sheet.cell(row=6, column=2, value=f"{month_name}")
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

    # Hide cell lines surrounding the data of activities
    for row in sheet.iter_rows(min_row=9, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        for cell in row:
            cell.border = None

    sheet.freeze_panes = sheet.cell(row=9, column=1)

