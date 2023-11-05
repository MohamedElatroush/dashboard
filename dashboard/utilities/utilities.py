import uuid  
import pytz
from ..constants import constants

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
    if not user.is_superuser:
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

def convert_nat_group_to_choice(nat_group_string):
    nat_group_mappings = {
        'VCH': constants.VCH,
        'FUD': constants.FUD,
        'TD': constants.TD,
        'CWD': constants.CWD,
        'PSD': constants.PSD,
        'RSMD': constants.RSMD
    }
    return nat_group_mappings.get(nat_group_string, None)

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
    base_username = "_".join(name_parts[:2]).lower()  # Take the first two names and join with an underscore.
    # Initialize a counter
    counter = 0
    unique_username = base_username

    # Check if the username exists in the taken usernames set and find a unique one by appending a number
    while unique_username in taken_usernames:
        counter += 1
        unique_username = f"{base_username}{counter}"

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