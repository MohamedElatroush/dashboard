from django.utils.translation import gettext_lazy as _

GRADE_A_1 = 0
GRADE_A_2 = 1
GRADE_A_3 = 2
GRADE_B_1 = 5
GRADE_B_2 = 6
GRADE_B_3 = 7
GRADE_B_4 = 8
GRADE_B_5 = 9


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
    (VCH, _('Vice Chairman')),
    (FUD, _('Follow Up Department')),
    (TD, _('Technical Department')),
    (CWD, _('Civil Work Department')),
    (PSD, _('Power Supply Department')),
    (RSMD, _('Rolling Stock & Mechanical Department')),
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

SUCCESSFULLY_CREATED_ACTIVITY = {
    "detail": "Activity has been created successfully"
}

SUCCESSFULLY_DELETED_ACTIVITY = {
    "detail": "Activity has been deleted successfully"
}
LAST_SUPERUSER_DELETION_ERROR = {
    "detail": "Cannot delete the last admin available."
}

# Define constants for column headers
EXPORT_ACTIVITY_COLUMNS = ["Username", "Type", "Date", "Time", "User Activity"]