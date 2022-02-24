from django.urls import path

from . import views

app_name = "myapp"

urlpatterns = [
    path('', views.index, name='index'),
    path('excel_download', views.excel_download, name='excel_download'),
]
