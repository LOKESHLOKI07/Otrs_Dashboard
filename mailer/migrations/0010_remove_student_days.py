# Generated by Django 3.2.18 on 2023-04-07 05:18

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('mailer', '0009_auto_20230407_1016'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='student',
            name='days',
        ),
    ]
