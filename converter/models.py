from django.db import models

class UploadedPDF(models.Model):
    pdf_file = models.FileField(upload_to='uploaded_pdfs/')
    he_dao_tao = models.CharField(max_length=100)
    danh_sach_quyet_dinh = models.CharField(max_length=100)

    def __str__(self):
        return f"{self.pdf_file.name} ({self.he_dao_tao})"

    class Meta:
        db_table = 'uploaded_pdf_files'  # Tên bảng tùy chỉnh
