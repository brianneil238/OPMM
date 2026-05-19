#!/usr/bin/env bash
# Render build script for the SOPM Django app.
# Runs on every deploy: install deps, collect static assets, then migrate the Postgres DB.
set -o errexit

pip install --upgrade pip
pip install -r requirements.txt

python manage.py collectstatic --no-input
python manage.py migrate --no-input
