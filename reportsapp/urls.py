from django.contrib import admin
from django.urls import path
from reportsapp import views

urlpatterns = [

    path('', views.dashboard, name='home'),
    path('customer/', views.customer, name='customer'),
    path('engineer/', views.engineer, name='engineer'),
    path('fullsummary/', views.fullsummary, name='fullsummary'),
    path('fullcustomer/', views.fullcustomer, name='fullcustomer'),
    path('dashboard/', views.dashboard, name='dashboard'),
    path('otp/', views.otp, name='otp'),
    path('student_detail/', views.StudentListView.as_view(), name='student_detail'),
    #path('edit_student/', views.edit_student, name='edit_student'),
    path('edit_student/<int:id>/', views.edit_student, name='edit_student'),
    # path('student_edit/', views.StudentUpdateView, name='student_edit'),
    path('plotly/', views.plotly, name='plotly'),


]
