from django.contrib import admin
from django.urls import path
from reportsapp import views
from django.contrib.auth import views as auth_views
from django.contrib.auth.decorators import login_required


urlpatterns = [
    path('', views.dashboard, name='home'),
    path('customer/', views.customer, name='customer'),
    path('engineer/', views.engineer, name='engineer'),
    path('fullsummary/', views.fullsummary, name='fullsummary'),
    path('fullcustomer/', views.fullcustomer, name='fullcustomer'),
    path('dashboard/', views.dashboard, name='dashboard'),
    path('minidash/', views.minidash, name='minidash'),
    path('otp/', views.otp, name='otp'),
    path('student_detail/', login_required(views.StudentListView.as_view()), name='student_detail'),
    path('students/<int:pk>/delete/', views.StudentDeleteView.as_view(), name='delete_student'),
    path('edit_student/<int:id>/', views.edit_student, name='edit_student'),
    path('plotly/', views.plotly, name='plotly'),
    path('junk/', views.junk, name='junk'),
    path('login/', views.CustomLoginView.as_view(), name='login'),
    path('open/', views.open, name='open'),
    path('active/', views.active_ticket, name='active'),

]
