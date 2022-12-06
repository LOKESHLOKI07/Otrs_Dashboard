from django import forms

from mailer.models import Student


class ContactForm(forms.ModelForm):
    # hour = forms.TimeField(widget=forms.TimeInput(format='%H'))
    # minutes = forms.TimeField(widget=forms.TimeInput(format='%M'))
    # my_time = serializers.TimeField(format='%H:%M')

    class Meta:
        model = Student
        fields = '__all__'
