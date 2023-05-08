from django.contrib import admin
from django.urls import path

from .views import contactView, successView,CustomScheduleView

urlpatterns = [
    path("contact/", contactView, name="contact"),
    path("success/", successView, name="success"),
    path('login1/', CustomScheduleView.as_view(), name='login1'),

]
