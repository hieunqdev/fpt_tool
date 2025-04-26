# serializers.py
from rest_framework import serializers
from .models import UploadedPDF

class UploadedPDFSerializer(serializers.ModelSerializer):
    class Meta:
        model = UploadedPDF
        fields = '__all__'  # Hoặc chọn cụ thể: ['id', 'pdf_file', 'he_dao_tao', 'danh_sach_quyet_dinh']
