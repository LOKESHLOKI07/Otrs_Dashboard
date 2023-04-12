from datetime import datetime
import os
from django.db import models
from django.forms import CheckboxSelectMultiple
from rest_framework import serializers
from multiselectfield import MultiSelectField
#Create your models here.
class Student(models.Model):
    val5 = []
    for root, dirs, files in os.walk("/home/ubuntu/reportproject/rameez", topdown=False):
        for name in files:
            path = os.path.join(name)
            if path.endswith(".py"):
                val5.append(path)

    marcas = (
        ('rameez1.py', val5[0]),
        ('rameez2.py', val5[1]),
    )
    day = (('Monday', 'Monday'),
           ('Tuesday', 'Tuesday'),
           ('Wednesday', 'Wednesday'),
           ('Thursday', 'Thursday'),
           ('Friday', 'Friday'),
           ('Saturday', 'Saturday'),
           ('Sunday', 'Sunday'))
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
    days = models.CharField(max_length=100,choices=day)
    Engineer_Name = models.CharField(max_length=100, choices=marcas)

    def __str__(self):
        return self.sender1