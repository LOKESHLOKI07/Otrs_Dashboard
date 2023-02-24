from django import forms

from mailer.models import Student
# from django.forms.widgets import CheckboxSelectMultiple
from django.forms.widgets import CheckboxSelectMultiple


class ContactForm(forms.ModelForm):
    class Meta:
        model = Student
        # widgets = {'days': Student.day.RadioSelect}
        # day = forms.MultipleChoiceField(choices=Student.day, widget=forms.CheckboxSelectMultiple())

        fields = '__all__'
