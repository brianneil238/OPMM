from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from django.contrib.auth.models import User
from .models import Announcement, Office, StrategicLevel, Indicator, PerformanceRecord

# We override the Save logic to automate Office creation
class OfficeUserAdmin(UserAdmin):
    def save_model(self, request, obj, form, change):
        super().save_model(request, obj, form, change)
        
        # If not an admin, create an office named exactly after the username
        if not obj.is_superuser:
            Office.objects.get_or_create(
                name=obj.username,
                defaults={'code': obj.username[:5].upper()}
            )

# UNREGISTER then RE-REGISTER the User model
try:
    admin.site.unregister(User)
except admin.sites.NotRegistered:   
    pass

admin.site.register(User, OfficeUserAdmin)

# Show Offices in the admin so you can see the automation working
@admin.register(Office)
class OfficeAdmin(admin.ModelAdmin):
    list_display = ('name', 'code')


@admin.register(PerformanceRecord)
class PerformanceRecordAdmin(admin.ModelAdmin):
    list_display = ('indicator', 'year', 'quarter', 'target_value', 'actual_value', 'status')
    list_filter = ('year', 'quarter')
    search_fields = ('indicator__description',)


@admin.register(Announcement)
class AnnouncementAdmin(admin.ModelAdmin):
    list_display = ('title', 'scope', 'is_active', 'sort_order', 'created_at')
    list_filter = ('scope', 'is_active')
    filter_horizontal = ('offices',)
    search_fields = ('title', 'body')