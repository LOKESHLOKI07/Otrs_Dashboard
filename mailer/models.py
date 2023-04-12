import calendar
import os

from django.db import models
from django import forms
from multiselectfield import MultiSelectField

# Get the list of files that end with ".py" in the specified directory
val5 = [os.path.join(root, name) for root, dirs, files in os.walk("/home/ubuntu/reportproject/rameez", topdown=False)
        for name in files if name.endswith(".py")]

# Create the choices tuple using the list of files, with only the script names
marcas = tuple((os.path.splitext(os.path.basename(f))[0], os.path.splitext(os.path.basename(f))[0]) for f in val5)

# Define the choices for the days of the week
day_choices = [(day, day) for day in calendar.day_name]


# Create your models here.
class Student(models.Model):
    sender1 = models.EmailField()
    sender2 = models.EmailField(blank=True)
    sender3 = models.EmailField(blank=True)
    sender4 = models.EmailField(blank=True)
    cc1 = models.EmailField(blank=True)
    cc2 = models.EmailField(blank=True)
    cc3 = models.EmailField(blank=True)
    cc4 = models.EmailField(blank=True)
    hour = models.IntegerField()
    minutes = models.IntegerField()
    days = MultiSelectField(choices=day_choices, max_choices=7, blank=True)

    Engineer_Name = models.CharField(max_length=100, choices=marcas)

    def __str__(self):
        return self.sender1


class OTP(models.Model):
    email = models.EmailField()
    otp = models.CharField(max_length=6)
    created_at = models.DateTimeField(auto_now_add=True)

