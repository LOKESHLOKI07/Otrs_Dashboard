from django.contrib import admin
from django.urls import path
from reportsapp import views

urlpatterns = [
    # path('fun/', views.fun, name='fun'),
    path('', views.home, name='home'),
    path('customer/', views.customer, name='customer'),
    path('engineer/', views.engineer, name='engineer'),
    path('fullsummary/', views.fullsummary, name='fullsummary'),
]
