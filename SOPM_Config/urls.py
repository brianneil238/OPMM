"""
SOPM_Config URL Configuration
"""
from django.contrib import admin
from django.urls import path, include
from core import views
from django.contrib.auth import views as auth_views

urlpatterns = [
    # 1. Django Administrative Interface
    path('admin/', admin.site.urls),
    
    # 2. Authentication System (Login/Logout)
    # Using built-in Django auth views mapped to your custom templates
    path('accounts/login/', auth_views.LoginView.as_view(template_name='registration/login.html'), name='login'),
    path('accounts/logout/', auth_views.LogoutView.as_view(), name='logout'),
    path('accounts/', include('django.contrib.auth.urls')),

    # 3. Core Dashboard & HTMX Search
    path('dashboard/', views.dashboard_home, name='dashboard_home'),
    path('search/', views.search_suggestions, name='search_suggestions'),
    
    # 4. Institutional User & Office Administration
    # Handles Full Name + Abbreviation registration and Delete/Edit actions
    path('users/', views.user_management, name='user_management'),
    path('system/announcements/', views.announcement_manage, name='announcement_manage'),
    path(
        'system/announcements/mark-read/',
        views.announcement_mark_read,
        name='announcement_mark_read',
    ),
    path('users/edit/<int:user_id>/', views.edit_user, name='edit_user'),
    path('users/delete/<int:user_id>/', views.delete_user, name='delete_user'),
    
    # 5. Data Migration & System Maintenance
    path('upload/', views.upload_blueprint, name='upload_blueprint'),
    path('success/', views.success_page, name='success_page'),
    path('reset-offices/', views.selective_office_reset, name='selective_office_reset'),
    path('system/download-database-backup/', views.download_database_backup, name='download_database_backup'),
    path('system/restore-database-backup/', views.restore_database_backup, name='restore_database_backup'),
    path('clear-data/', views.clear_all_data, name='clear_all_data'),
    path('clear-office-data/', views.clear_office_data, name='clear_office_data'),
    path('performance-viewer/', views.performance_viewer, name='performance_viewer'),
    path('analytics/', views.analytics_dashboard, name='analytics_dashboard'),
    path('activity-log/', views.activity_log, name='activity_log'),
    path(
        'performance-viewer/indicator/<int:record_id>/',
        views.performance_viewer_indicator_detail,
        name='performance_viewer_indicator_detail',
    ),
    path(
        'performance-viewer/export-lpc-wide.xlsx',
        views.export_lpc_wide_excel,
        name='export_lpc_wide_excel',
    ),

    # 6. Root Redirect
    # Automatically sends users to the dashboard when they visit the base URL
    path('', views.dashboard_home, name='root'),
]