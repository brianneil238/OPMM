# Generated manually for nullable KPI targets (no numeric target parsed from Word).

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0004_alter_performancerecord_year'),
    ]

    operations = [
        migrations.AlterField(
            model_name='performancerecord',
            name='target_value',
            field=models.FloatField(blank=True, null=True),
        ),
    ]
