from django.core.mail import send_mail, BadHeaderError
from django.http import HttpResponse
from django.shortcuts import render, redirect
from django.core.mail import EmailMessage
from datetime import datetime
# from .models import Student
import os

import reportproject.settings
from .forms import ContactForm


def contactView(request):
    if request.method == "GET":
        form = ContactForm()


    else:
        form = ContactForm(request.POST)

        if form.is_valid():
            # band = ContactForm()
            sender1 = form.cleaned_data['sender1']
            sender2 = form.cleaned_data['sender2']
            sender3 = form.cleaned_data['sender3']
            sender4 = form.cleaned_data['sender4']

            cc1 = form.cleaned_data['cc1']
            cc2 = form.cleaned_data['cc2']

            cc3 = form.cleaned_data['cc3']
            cc4 = form.cleaned_data['cc4']
            hours = form.cleaned_data['hour']
            minutes = form.cleaned_data['minutes']
            Engineer_Name = form.cleaned_data['Engineer_Name']

            form.save()

            """try:            
                email = EmailMessage(subject="daily report:" + str(hours) + str(minutes) + "", body='hello',
                                     from_email=reportproject.settings.EMAIL_HOST_USER,
                                     to=[sender1, sender2, sender3, sender4], cc=[cc1, cc2, cc3, cc4])

                email.send(fail_silently=True)

            except BadHeaderError:
                return HttpResponse('Invalid header found.')"""
            return redirect('success')
            # return render(request, "email.html", context)
    # return render(request, "email.html")
    return render(request, "email.html", {"form": form})


def successView(request):
    return HttpResponse("Success! Thank you for your message")
