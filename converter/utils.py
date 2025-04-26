# yourapp/utils.py
from openpyxl import Workbook
import pdfplumber
import re

def convert_pdf_to_excel(pdf_file):
    wb = Workbook()
    ws = wb.active

    keyword_map = {
        "mssv": ["mssv", "mã sv", "mã số sv", "student id"],
        "họ và tên": ["họ và tên", "họ tên", "họ & tên", "tên", "họ\ntên"],
        "ngày sinh": ["ngày sinh", "ns", "dob", "ngày\nsinh", "sinh ngày"],
        "giới tính": ["giới tính", "giới\ntính", "gt", "sex"],
        "dân tộc": ["dân tộc", "dân\ntộc", "ethnic", "ethnicity"]
    }

    target_columns = {key: None for key in keyword_map}
    extracted_data = []
    extra_info = {
        "so": None,
        "ngay_thang_nam": None
    }

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # Trích bảng (nếu có)
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    header_found = False
                    for row in table:
                        ws.append(row)

                        if not header_found:
                            for i, cell in enumerate(row):
                                if cell:
                                    cell_clean = cell.strip().lower().replace('\n', ' ')
                                    for key, keywords in keyword_map.items():
                                        for keyword in keywords:
                                            if keyword in cell_clean and target_columns[key] is None:
                                                target_columns[key] = i
                            if any(v is not None for v in target_columns.values()):
                                header_found = True
                        elif header_found:
                            if len(extracted_data) < 10:
                                row_data = {}
                                for key, idx in target_columns.items():
                                    if idx is not None and idx < len(row):
                                        row_data[key] = row[idx]
                                    else:
                                        row_data[key] = None
                                extracted_data.append(row_data)

            # Trích văn bản để tìm dòng "Số:" và "Hà Nội, ngày..."
            text = page.extract_text()
            if text:
                for line in text.split('\n'):
                    ws.append([line])  # ghi vào excel luôn

                    stripped = line.strip()
                    if stripped.startswith("Số:") and not extra_info["so"]:
                        extra_info["so"] = stripped
                    if stripped.startswith("Hà Nội, ngày") and not extra_info["ngay_thang_nam"]:
                        extra_info["ngay_thang_nam"] = stripped

    return wb, extracted_data, extra_info
