import os
from pathlib import Path

import dj_database_url

BASE_DIR = Path(__file__).resolve().parent.parent

SECRET_KEY = os.environ.get(
    'DJANGO_SECRET_KEY',
    'django-insecure-dev-only-change-for-production',
)
DEBUG = os.environ.get('DJANGO_DEBUG', 'true').lower() in ('1', 'true', 'yes')

_allowed = os.environ.get('DJANGO_ALLOWED_HOSTS', 'localhost,127.0.0.1')
ALLOWED_HOSTS = [h.strip() for h in _allowed.split(',') if h.strip()]
# Render injects the service's external hostname here automatically.
_render_host = os.environ.get('RENDER_EXTERNAL_HOSTNAME', '').strip()
if _render_host and _render_host not in ALLOWED_HOSTS:
    ALLOWED_HOSTS.append(_render_host)

_csrf_origins = os.environ.get('DJANGO_CSRF_TRUSTED_ORIGINS', '').strip()
CSRF_TRUSTED_ORIGINS = [o.strip() for o in _csrf_origins.split(',') if o.strip()]
if _render_host:
    CSRF_TRUSTED_ORIGINS.append(f'https://{_render_host}')

# Full DB wipe (/clear-data/): allowed when DEBUG, or when explicitly enabled in production.
SOPM_ENABLE_FULL_DATABASE_CLEAR = os.environ.get(
    'SOPM_ENABLE_FULL_DATABASE_CLEAR', ''
).lower() in ('1', 'true', 'yes')

# Restore DB from uploaded SQLite (/system/restore-database-backup/): DEBUG, or enable in production.
SOPM_ENABLE_DATABASE_RESTORE = os.environ.get(
    'SOPM_ENABLE_DATABASE_RESTORE', ''
).lower() in ('1', 'true', 'yes')

INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'core',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    # WhiteNoise serves collected static files directly from the WSGI app on Render.
    'whitenoise.middleware.WhiteNoiseMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]

# Ensure these match your 'SOPM_Config' folder name
ROOT_URLCONF = 'SOPM_Config.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.debug',
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
                'core.context_processors.nav_announcements',
            ],
        },
    },
]

WSGI_APPLICATION = 'SOPM_Config.wsgi.application'

# Database: use ``DATABASE_URL`` when present (Render Postgres), fall back to local SQLite for dev.
_default_db_url = f'sqlite:///{BASE_DIR / "db.sqlite3"}'
DATABASES = {
    'default': dj_database_url.config(
        default=os.environ.get('DATABASE_URL', _default_db_url),
        conn_max_age=600,
        conn_health_checks=True,
        ssl_require=os.environ.get('DATABASE_SSL_REQUIRE', '').lower() in ('1', 'true', 'yes'),
    )
}

if DEBUG:
    AUTH_PASSWORD_VALIDATORS = []
else:
    AUTH_PASSWORD_VALIDATORS = [
        {'NAME': 'django.contrib.auth.password_validation.UserAttributeSimilarityValidator'},
        {'NAME': 'django.contrib.auth.password_validation.MinimumLengthValidator'},
        {'NAME': 'django.contrib.auth.password_validation.CommonPasswordValidator'},
        {'NAME': 'django.contrib.auth.password_validation.NumericPasswordValidator'},
    ]

LANGUAGE_CODE = 'en-us'
TIME_ZONE = 'UTC'
USE_I18N = True
USE_TZ = True

STATIC_URL = '/static/'
STATIC_ROOT = BASE_DIR / 'staticfiles'
STATICFILES_STORAGE = 'whitenoise.storage.CompressedManifestStaticFilesStorage'

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

# Auth Redirects
LOGIN_REDIRECT_URL = 'dashboard_home'
LOGOUT_REDIRECT_URL = 'login'

# Production hardening: enable when not running in DEBUG (Render terminates TLS at the proxy).
if not DEBUG:
    SECURE_PROXY_SSL_HEADER = ('HTTP_X_FORWARDED_PROTO', 'https')
    SECURE_SSL_REDIRECT = os.environ.get('DJANGO_SECURE_SSL_REDIRECT', 'true').lower() in ('1', 'true', 'yes')
    SESSION_COOKIE_SECURE = True
    CSRF_COOKIE_SECURE = True
    SECURE_HSTS_SECONDS = int(os.environ.get('DJANGO_SECURE_HSTS_SECONDS', '0') or '0')
    SECURE_HSTS_INCLUDE_SUBDOMAINS = True
    SECURE_HSTS_PRELOAD = False
    SECURE_CONTENT_TYPE_NOSNIFF = True
    SECURE_REFERRER_POLICY = 'same-origin'
