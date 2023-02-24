# Generated by Django 4.0.6 on 2022-12-16 08:58

from django.db import migrations
import multiselectfield.db.fields


class Migration(migrations.Migration):

    dependencies = [
        ('mailer', '0003_alter_student_days'),
    ]

    operations = [
        migrations.AlterField(
            model_name='student',
            name='days',
            field=multiselectfield.db.fields.MultiSelectField(choices=[('Monday', 'Monday'), ('Tuesday', 'Tuesday'), ('Wednesday', 'Wednesday'), ('Thursday', 'Thursday'), ('Friday', 'Friday'), ('Saturday', 'Saturday'), ('Sunday', 'Sunday')], max_length=56),
        ),
    ]
