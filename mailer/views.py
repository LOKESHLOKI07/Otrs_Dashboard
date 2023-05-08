from django.core.exceptions import ImproperlyConfigured
from django.core.mail import send_mail, BadHeaderError, EmailMessage
from django.http import HttpResponse
from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login
from django.contrib.auth.decorators import login_required
from django.utils.decorators import method_decorator
from django.urls import reverse_lazy
from django.views.generic.edit import FormView
from django.contrib.auth.forms import AuthenticationForm

from .forms import ContactForm
from .decorators import login_exempt


class CustomScheduleView(FormView):
    form_class = AuthenticationForm
    template_name = 'LoginView1.html'
    success_url = reverse_lazy('contact')

    def form_valid(self, form):
        username = form.cleaned_data['username']
        password = form.cleaned_data['password']
        user = authenticate(self.request, username=username, password=password)
        if user is not None:
            login(self.request, user)
            return super().form_valid(form)
        else:
            return self.form_invalid(form)


@login_required
def contactView(request):
    if request.method == "GET":
        form = ContactForm()
    else:
        form = ContactForm(request.POST)
        if form.is_valid():
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
            return redirect('success')
    return render(request, "email.html", {"form": form})

# def contact(request):
#     if request.method == 'POST':
#         form = ContactForm(request.POST)
#         if form.is_valid():
#             # Create a new instance of the Student model using the form data
#             student = Student(
#                 sender1=form.cleaned_data['sender1'],
#                 sender2=form.cleaned_data['sender2'],
#                 sender3=form.cleaned_data['sender3'],
#                 sender4=form.cleaned_data['sender4'],
#                 cc1=form.cleaned_data['cc1'],
#                 cc2=form.cleaned_data['cc2'],
#                 cc3=form.cleaned_data['cc3'],
#                 cc4=form.cleaned_data['cc4'],
#                 hour=form.cleaned_data['hour'],
#                 minutes=form.cleaned_data['minutes'],
#                 days=form.cleaned_data['days'],
#                 Engineer_Name=form.cleaned_data['Engineer_Name']
#             )
#             # Save the new student object to the database
#             student.save()
#             # Redirect to a success page
#             return redirect('success')
#     else:
#         form = ContactForm()
#     return render(request, 'contact.html', {'form': form})
def successView(request):
    return HttpResponse("Success! Thank you for your message")
