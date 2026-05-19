#!/usr/bin/env bash
# Render build script for the SOPM Django app.
# Runs on every deploy: install deps, collect static assets, then migrate the Postgres DB.
set -o errexit

pip install --upgrade pip
pip install -r requirements.txt

python manage.py collectstatic --no-input
python manage.py migrate --no-input

# Free Render plans have no web shell. To create the first admin, set these env vars once
# in the web service (then remove the password var after deploy succeeds):
#   DJANGO_SUPERUSER_USERNAME, DJANGO_SUPERUSER_EMAIL, DJANGO_SUPERUSER_PASSWORD
if [ -n "${DJANGO_SUPERUSER_USERNAME:-}" ] && [ -n "${DJANGO_SUPERUSER_PASSWORD:-}" ]; then
  export DJANGO_SUPERUSER_EMAIL="${DJANGO_SUPERUSER_EMAIL:-admin@example.com}"
  python manage.py createsuperuser --noinput || true
fi
