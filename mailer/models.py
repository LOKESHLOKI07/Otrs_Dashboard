from datetime import datetime
import os
from django.db import models
from rest_framework import serializers


# Create your models here.
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
    sender1 = models.EmailField()
    sender2 = models.EmailField()
    sender3 = models.EmailField()
    sender4 = models.EmailField()
    cc1 = models.EmailField()
    cc2 = models.EmailField()
    cc3 = models.EmailField()
    cc4 = models.EmailField()
    hour = models.IntegerField()
    minutes = models.IntegerField()
    Engineer_Name = models.CharField(max_length=100, choices=marcas)

    def __str__(self):
        return self.sender1
