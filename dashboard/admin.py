from django.contrib import admin
from .models import User, Activity, ActivityFile
class ActivityFileAdmin(admin.ModelAdmin):
    list_display = ('created', 'file')
    readonly_fields = ('created',)

# Register your models here.
admin.site.register(User)
admin.site.register(Activity)
admin.site.register(ActivityFile, ActivityFileAdmin)