from ..constants import constants
import uuid  
import pytz


def generate_unique_id():
    new_id = "EMP_" + str(uuid.uuid1())
    return new_id

# Helper function to convert and format a datetime to Cairo timezone
def convert_to_cairo_timezone_and_format(datetime_obj):
    cairo_timezone = pytz.timezone('Africa/Cairo')
    return datetime_obj.astimezone(cairo_timezone).strftime("%Y-%m-%d %H:%M:%S")
