from ..constants import constants
import uuid  

def generate_unique_id():
    new_id = "EMP_" + str(uuid.uuid1())
    return new_id