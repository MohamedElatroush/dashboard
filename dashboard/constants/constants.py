from django.utils.translation import gettext_lazy as _
from openpyxl.styles import Font, Alignment, Border, Side

GRADE_A_1 = 0
GRADE_A_2 = 1
GRADE_A_3 = 2
GRADE_B_1 = 3
GRADE_B_2 = 4
GRADE_B_3 = 5
GRADE_B_4 = 6
GRADE_B_5 = 7


USER_GRADE_CHOICES = (
    (GRADE_A_1, _('A1')),
    (GRADE_A_2, _('A2')),
    (GRADE_A_3, _('A3')),
    (GRADE_B_1, _('B1')),
    (GRADE_B_2, _('B2')),
    (GRADE_B_3, _('B3')),
    (GRADE_B_4, _('B4')),
    (GRADE_B_5, _('B5')),
)

# expert / local
EXPERT_USER = 0
LOCAL_USER = 1

EXPERT_LOCAL_CHOICES = (
    (EXPERT_USER, _('EXP')),
    (LOCAL_USER, _('LOC')),
)

VCH = 0
FUD = 1
TD = 2
CWD = 3
PSD = 4
RSMD = 5

NAT_GROUP_CHOICES = (
    (VCH, _('VCH')),
    (FUD, _('FUD')),
    (TD, _('TD')),
    (CWD, _('CWD')),
    (PSD, _('PSD')),
    (RSMD, _('RSMD')),
)

OCG = 0
NK = 1
EHAF = 2
ACE = 3

COMPANY_CHOICES = (
    (OCG, _('OCG')),
    (NK, _('NK')),
    (EHAF, _('EHAF')),
    (ACE, _('ACE')),
)

SUCCESSFULLY_DELETED_USER = {
    "Success": "The user has successfully been deleted"
}

SUCCESSFULLY_UPDATED_ACTIVITY = {
    "detail": "Activity has been modified successfully"
}

NOT_ALLOWED_TO_ACCESS = {
    "detail": "You are not allowed to access that!"
}

SUCCESSFULLY_CHANGED_PASSWORD = {
    "detail": "Password has been created successfully"
}

SUCCESSFULLY_CREATED_ACTIVITY = {
    "detail": "Activity has been created successfully"
}

SUCCESSFULLY_DELETED_ACTIVITY = {
    "detail": "Activity has been deleted successfully"
}
LAST_SUPERUSER_DELETION_ERROR = {
    "detail": "Cannot delete the last admin available."
}
CANT_RESET_USER_PASSWORD_ERROR = {
    "detail": "Failed to reset password for that user"
}

ERR_NO_ACTIVITY_ID_FOUND = {
    "detail": "Failed to edit activity with the provided ID"
}

ERR_NO_AVAILABLE_TS_AVAILABLE = {
    "detail": "No available timesheet at the provided date"
}

ERR_HR_CODE_USED = {
    "detail": "An HR code was used before"
}

# Define constants for column headers
EXPORT_ACTIVITY_COLUMNS = ["#", "Name", "Source", "Position", "/", "Grade", "Group"]

HOLIDAY = 0 
INCAIRO = 1
OFFDAY = 2
HOMEASSIGN = 3

ACTIVITY_TYPES_CHOICES = (
    (HOLIDAY, _('H')),
    (INCAIRO, _('C')),
    (OFFDAY, _('X')),
    (HOMEASSIGN, _('J')),
)

ERR_PASSWORD_RESET_NEEDED = {
    "detail": "Please reset your password first before performing that action, Go to profile > Change Password button"
}

DEP_HEADERS = ['No.', 'Name']

CALENDAR_TYPE_EXPERT = 0
CALENDAR_TYPE_LOCAL = 1

CALENDAR_CHOICES = (
    (CALENDAR_TYPE_EXPERT, _('Expert Calendar')),
    (CALENDAR_TYPE_LOCAL, _('Local Calendar')),
)

COVER_TS_TEXT = "Consulting Services for Greater Cairo Metro Line No. 4 Phase 1 Project"

BORDER_STYLE = Border(top=Side(style='thin', color='000000'),
                     left=Side(style='thin', color='000000'),
                     right=Side(style='thin', color='000000'),
                     bottom=Side(style='thin', color='000000'))

DEP_MAPPING = {
    VCH: 'Vice Chairman',
    FUD: 'Follow Up Department',
    TD: 'Technical Department',
    CWD: 'Civil Work Department',
    PSD: 'Power Supply Department',
    RSMD: 'Rolling Stock & Mechanical Department',
}