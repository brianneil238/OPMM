"""Create or reset a superuser (local .env or Render Postgres via DATABASE_URL)."""

from django.contrib.auth import get_user_model
from django.core.management.base import BaseCommand, CommandError
from django.db import connection


class Command(BaseCommand):
    help = (
        'Create or update a superuser password in the configured database. '
        'Use with .env pointing at Render External DATABASE_URL to fix production login.'
    )

    def add_arguments(self, parser):
        parser.add_argument('--username', default='admin', help='Login username (default: admin)')
        parser.add_argument('--password', required=True, help='New password for this account')
        parser.add_argument('--email', default='', help='Email (optional)')

    def handle(self, *args, **options):
        username = (options['username'] or '').strip()
        password = options['password']
        email = (options['email'] or '').strip() or f'{username}@example.com'

        if not username:
            raise CommandError('--username is required')
        if not password:
            raise CommandError('--password is required')

        User = get_user_model()
        user, created = User.objects.get_or_create(
            username=username,
            defaults={'email': email, 'is_staff': True, 'is_superuser': True, 'is_active': True},
        )
        user.email = email or user.email
        user.is_staff = True
        user.is_superuser = True
        user.is_active = True
        user.set_password(password)
        user.save()

        db = connection.settings_dict
        host = db.get('HOST') or '(sqlite)'
        self.stdout.write(
            self.style.SUCCESS(
                f"{'Created' if created else 'Updated'} superuser '{username}' on database host {host}."
            )
        )
        self.stdout.write('Sign in at /accounts/login/ with that username and password.')
