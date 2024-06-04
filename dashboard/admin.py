from django.contrib import admin
from .models import User, Activity, hrHistory, Department

# Register your models here.
admin.site.register(User)
admin.site.register(Activity)
admin.site.register(hrHistory)
admin.site.register(Department)