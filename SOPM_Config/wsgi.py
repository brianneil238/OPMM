import os
from django.core.wsgi import get_wsgi_application

# Ensure 'SOPM_Config' matches your directory name
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'SOPM_Config.settings')

application = get_wsgi_application()