from django.utils.translation import gettext_lazy as _

USER_TYPE_A = 0
USER_TYPE_E = 1

USER_TYPE_CHOICES = (
    (USER_TYPE_A, _('Type A: International')),
    (USER_TYPE_E, _('Type E: Egyptian')),
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