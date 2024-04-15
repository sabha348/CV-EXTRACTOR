from django.urls import path
from . import views

urlpatterns = [
    path('upload/', views.upload_cv, name='upload_cv'),
    path('download/<path:file_path>/', views.download_cv, name='download_cv'),
    path('', views.home, name='home'),
]