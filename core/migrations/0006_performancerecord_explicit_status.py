from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0005_alter_performancerecord_target_value'),
    ]

    operations = [
        migrations.AddField(
            model_name='performancerecord',
            name='explicit_status',
            field=models.CharField(blank=True, max_length=8, null=True),
        ),
    ]
