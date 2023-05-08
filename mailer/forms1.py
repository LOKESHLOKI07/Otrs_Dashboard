import os

from django import forms

import mailer.models
from mailer.models import Student
# from django.forms.widgets import CheckboxSelectMultiple
from django.forms.widgets import CheckboxSelectMultiple
from .models import Student
import calendar
day_choices = [(day, day) for day in calendar.day_name]

import calendar

# Get the list of files that end with ".py" in the specified directory
val5 = [os.path.join(root, name) for root, dirs, files in os.walk("/home/ubuntu/reportproject/rameez", topdown=False)
        for name in files if name.endswith(".py")]

# Create the choices tuple using the list of files, with only the script names
marcas = tuple((os.path.splitext(os.path.basename(f))[0], os.path.splitext(os.path.basename(f))[0]) for f in val5)

class ContactForm(forms.Form):
    # Define the choices for the days dropdown using the calendar module
    day_choices = [(str(i), day) for i, day in enumerate(calendar.day_name)]

    # Define the fields for your form
    sender1 = forms.CharField(label='Sender 1')
    sender2 = forms.CharField(label='Sender 2')
    sender3 = forms.CharField(label='Sender 3')
    sender4 = forms.CharField(label='Sender 4')
    cc1 = forms.CharField(label='CC 1')
    cc2 = forms.CharField(label='CC 2')
    cc3 = forms.CharField(label='CC 3')
    cc4 = forms.CharField(label='CC 4')
    hour = forms.IntegerField(label='Hours')
    minutes = forms.IntegerField(label='Minutes')
    Engineer_Name = forms.ChoiceField(label='Script', choices=marcas, widget=forms.Select)

    days = forms.MultipleChoiceField(label='Days', choices=day_choices, widget=forms.Select)



