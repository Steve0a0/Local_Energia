from django.urls import path
from .views import FileUploadView

urlpatterns = [
    path('uploadenergia/', FileUploadView.as_view(), name='file-upload'),
]