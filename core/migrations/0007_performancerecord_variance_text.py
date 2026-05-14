from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0006_performancerecord_explicit_status'),
    ]

    operations = [
        migrations.AddField(
            model_name='performancerecord',
            name='variance_text',
            field=models.TextField(blank=True),
        ),
    ]
