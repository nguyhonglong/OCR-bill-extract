import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QLabel, QPushButton, 
                             QFileDialog, QGridLayout, QMessageBox)
from tqdm import tqdm
import pytesseract
import re
import openpyxl
import os
from PIL import Image, ImageEnhance

pytesseract.pytesseract.tesseract_cmd = r'Tesseract-OCR/tesseract.exe'
def extract_bill_info(image_path):
    img = Image.open(image_path)
    enhancer = ImageEnhance.Contrast(img)
    enhanced_img = enhancer.enhance(2)  

    enhancer = ImageEnhance.Brightness(enhanced_img)
    brightened_img = enhancer.enhance(1.5)  
    text = pytesseract.image_to_string(img, lang='eng')
    text2 = pytesseract.image_to_string(img, lang='vie')
    
    lines = text.splitlines()
    lines2 = text2.splitlines()

    lines = [line for line in lines if line.strip()]
    lines2 = [line for line in lines2 if line.strip()]
    
    info = {}
    while lines[0] != 'Thanh cong':
        lines.pop(0)
    while lines2[0] != 'Thành công':
        lines2.pop(0)
   

    info['so_tien'] = lines[1][1:-1].replace(" ", "")
    ten_nguoi_nhan = lines[3].strip()
    ten_nguoi_nhan = re.sub(r'[^a-zA-Z\s]', '', ten_nguoi_nhan) 
    ten_nguoi_nhan = ten_nguoi_nhan.strip() 
    ten_nguoi_nhan = re.sub(r'^\w\s', '', ten_nguoi_nhan) 
    info['ten_nguoi_nhan'] = ten_nguoi_nhan
    stk_position = None
    for i, line in enumerate(lines2):
        if 'Thời gian' in line:
            stk_position = i - 1  
            break
    
    if stk_position is not None and 4 <= stk_position <= 6:
        info['so_tai_khoan'] = re.sub(r'\D', '', lines2[stk_position].strip())
    else:
        info['so_tai_khoan'] = "Không tìm thấy"
    for item in lines2:
        if 'Thời gian' in item:
            info['thoi_gian'] = item.split('Thời gian ')[1]
        elif 'Mã tra soát' in item:
            info['ma_tra_soat'] = item.split('Mã tra soát ')[1]
        elif 'Nội dung' in item:
            info['noi_dung'] = item.split('Nội dung ')[1]
    info['link_anh'] =  image_path
    return info

class BillExtractor(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Trích xuất thông tin từ bill")
        self.grid = QGridLayout()
        self.setLayout(self.grid)
        self.setFixedSize(480, 320) 
        self.label_folder = QLabel("Chọn thư mục chứa ảnh:")
        self.grid.addWidget(self.label_folder, 0, 0)

        self.button_folder = QPushButton("Chọn thư mục")
        self.button_folder.clicked.connect(self.select_folder)
        self.grid.addWidget(self.button_folder, 0, 1)

        self.label_status = QLabel("")
        self.grid.addWidget(self.label_status, 1, 0, 1, 2)

        self.button_extract = QPushButton("Trích xuất")
        self.button_extract.clicked.connect(self.extract_info)
        self.button_extract.setEnabled(False)
        self.grid.addWidget(self.button_extract, 2, 0, 1, 2)

        self.folder_path = ""

    def select_folder(self):
        self.folder_path = QFileDialog.getExistingDirectory(self, "Chọn thư mục")
        if self.folder_path:
            self.label_status.setText(f"Đã chọn thư mục: {self.folder_path}")
            self.button_extract.setEnabled(True)
        else:
            self.label_status.setText("Chưa chọn thư mục")
            self.button_extract.setEnabled(False)

    def extract_info(self):
        if not self.folder_path:
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn thư mục chứa ảnh!")
            return

        try:
            workbook = openpyxl.Workbook()
            worksheet = workbook.active
            worksheet.append(['Link_ảnh', 'Số tiền', 'Tên người nhận', 'Số tài khoản', 'Thời gian', 'Nội dung', 'Mã tra soát'])

            image_files = [f for f in os.listdir(self.folder_path) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]

            for filename in tqdm(image_files, desc="Processing images"):
                image_path = os.path.join(self.folder_path, filename)
                
                try:
                    bill_info = extract_bill_info(image_path)
                    worksheet.append([
                        bill_info.get('link_anh', ''),
                        bill_info.get('so_tien', ''),
                        bill_info.get('ten_nguoi_nhan', ''),
                        bill_info.get('so_tai_khoan', ''),
                        bill_info.get('thoi_gian', ''),
                        bill_info.get('noi_dung', ''),
                        bill_info.get('ma_tra_soat', '')
                    ])
                except Exception as e:
                    worksheet.append([
                        image_path,                  # Link ảnh
                        'Lỗi lấy thông tin',         # Số tiền
                        'Lỗi lấy thông tin',         # Tên người nhận
                        'Lỗi lấy thông tin',         # Số tài khoản
                        'Lỗi lấy thông tin',         # Thời gian
                        'Lỗi lấy thông tin',         # Nội dung
                        'Lỗi lấy thông tin'          # Mã tra soát
                    ])
                    # print(f"Error processing {filename}: {e}")

            workbook.save("bill_info.xlsx")
            QMessageBox.information(self, "Thành công", "Trích xuất thông tin thành công!")

        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Lỗi xảy ra: {e}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = BillExtractor()
    ex.show()
    sys.exit(app.exec_())