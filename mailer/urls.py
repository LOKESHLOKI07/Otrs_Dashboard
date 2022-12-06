from django.contrib import admin
from django.urls import path

from .views import contactView, successView

urlpatterns = [
    path("contact/", contactView, name="contact"),
    path("success/", successView, name="success"),
    # path("project_index/", project_index, name="project_index"),

]