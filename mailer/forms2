from django import forms

import mailer.models
from mailer.models import Student
# from django.forms.widgets import CheckboxSelectMultiple
from django.forms.widgets import CheckboxSelectMultiple
from .models import Student
import calendar
day_choices = [(day, day) for day in calendar.day_name]

class ContactForm(forms.ModelForm):
    class Meta:
        model = Student
        fields = '__all__'

        days = forms.MultipleChoiceField(choices=day_choices, widget=forms.SelectMultiple())