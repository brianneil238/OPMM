import os
from django.core.asgi import get_asgi_application

# Ensure 'SOPM_Config' matches your directory name
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'SOPM_Config.settings')

application = get_asgi_application()