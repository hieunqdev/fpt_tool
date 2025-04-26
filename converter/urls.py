# converter/urls.py

from django.urls import path
from .views import hello_api, delete_uploaded_pdfs, upload_excel_api, UploadedPDFListView
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path('uploaded_pdfs/', UploadedPDFListView.as_view(), name='uploaded_pdfs'),

    path('hello/', hello_api, name='hello_api'),
    path('delete_uploaded_pdfs/', delete_uploaded_pdfs, name='delete_uploaded_pdfs'),
    path('upload_excel_api/', upload_excel_api, name='upload_excel_api'),
] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
