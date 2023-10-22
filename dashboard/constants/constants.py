from django.utils.translation import gettext_lazy as _

USER_TYPE_A = 0
USER_TYPE_E = 1

USER_TYPE_CHOICES = (
    (USER_TYPE_A, _('Type A')),
    (USER_TYPE_E, _('Type E')),
)

CANT_DELETE_USER_ERROR = {
    "Error": "Failed to delete user, please make sure of the username provided in the url"
}

SUCCESSFULLY_DELETED_USER = {
    "Success": "The user has successfully been deleted"
}