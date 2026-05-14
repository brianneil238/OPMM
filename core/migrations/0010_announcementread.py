# Generated manually for AnnouncementRead

import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ('core', '0009_announcement'),
    ]

    operations = [
        migrations.CreateModel(
            name='AnnouncementRead',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('read_at', models.DateTimeField(auto_now_add=True)),
                (
                    'announcement',
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name='reads',
                        to='core.announcement',
                    ),
                ),
                (
                    'user',
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name='announcement_reads',
                        to=settings.AUTH_USER_MODEL,
                    ),
                ),
            ],
            options={
                'ordering': ['-read_at', '-id'],
            },
        ),
        migrations.AddConstraint(
            model_name='announcementread',
            constraint=models.UniqueConstraint(
                fields=('user', 'announcement'),
                name='uniq_announcementread_user_announcement',
            ),
        ),
    ]
