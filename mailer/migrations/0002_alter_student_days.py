# Generated by Django 4.0.6 on 2022-12-16 05:50

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('mailer', '0001_initial'),
    ]

    operations = [
        migrations.AlterField(
            model_name='student',
            name='days',
            field=models.CharField(max_length=10),
        ),
    ]