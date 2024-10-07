from PyQt6 import QtCore, QtGui, QtWidgets
import pandas as pd
import os
import re  # Thư viện để xử lý ký tự không hợp lệ


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 529)
        self.centralwidget = QtWidgets.QWidget(parent=MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.frame = QtWidgets.QFrame(parent=self.centralwidget)
        self.frame.setGeometry(QtCore.QRect(30, -20, 751, 511))
        self.frame.setStyleSheet("background-repeat: no-repeat;\n"
                                 "background-position: center;")
        self.frame.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.frame.setObjectName("frame")

        self.label_3 = QtWidgets.QLabel(parent=self.frame)
        self.label_3.setGeometry(QtCore.QRect(150, 80, 481, 41))
        self.label_3.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.label_3.setObjectName("label_3")

        self.label_2 = QtWidgets.QLabel(parent=self.frame)
        self.label_2.setGeometry(QtCore.QRect(130, 250, 61, 20))
        self.label_2.setObjectName("label_2")

        self.lineEdit = QtWidgets.QLineEdit(parent=self.frame)
        self.lineEdit.setGeometry(QtCore.QRect(190, 240, 431, 31))
        self.lineEdit.setObjectName("lineEdit")

        self.comboBox = QtWidgets.QComboBox(parent=self.frame)
        self.comboBox.setGeometry(QtCore.QRect(290, 190, 101, 31))
        self.comboBox.setObjectName("comboBox")

        self.pushButton = QtWidgets.QPushButton(parent=self.frame)
        self.pushButton.setGeometry(QtCore.QRect(400, 190, 93, 31))
        self.pushButton.setObjectName("pushButton")

        self.label = QtWidgets.QLabel(parent=self.frame)
        self.label.setGeometry(QtCore.QRect(120, 190, 161, 20))
        self.label.setObjectName("label")

        self.pushButton.clicked.connect(self.export_data)  # Kết nối nút Save với hàm xuất dữ liệu

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(parent=MainWindow)
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(parent=MainWindow)
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        # Thêm các giá trị vào combobox
        self.comboBox.addItems(["UNIS", "UNIMAX", "UNILUX", "LOTINA", "KAIZEN", "UNICHEMI", "TẤT CẢ"])

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label_3.setText(_translate("MainWindow", "PHẦN MỀM CHỈNH SỬA BẢNG GIÁ TỰ ĐỘNG"))
        self.label_2.setText(_translate("MainWindow", "Vị trí lưu"))
        self.pushButton.setText(_translate("MainWindow", "Save"))
        self.label.setText(_translate("MainWindow", "Xuất file Push hệ thống"))

    def export_data(self):
        save_location = self.lineEdit.text()  # Lấy vị trí lưu từ lineEdit
        company = self.comboBox.currentText()  # Lấy tên công ty từ comboBox

        # Kiểm tra nếu vị trí lưu không hợp lệ
        if not save_location:
            QtWidgets.QMessageBox.warning(None, "Warning", "Please specify a save location.")
            return

        # Đọc file Excel chỉ với các cột cần thiết
        file_path = 'C:/Users/Admin/Desktop/file_pushgia3.xlsx'
        columns_needed = ['Cong Ty Trien Khai', 'Ten Chi Nhanh', 'Mặt hàng', 'Vùng bán hàng',
                          'LOAI BAN HANG', 'HTTT', 'Tên hàng hóa, dịch vụ', 'ĐVT_1',
                          'Đơn giá niêm yết', '% Giảm giá', 'Tiền Giảm giá',
                          'Đơn giá sau chiết khấu', 'TU', 'DEN', 'Nhom Gia',
                          'Biên độ giảm', 'Biên độ tăng', 'Kích thước',
                          'Khung Giá', 'Mã SP', 'Loại', 'Gạch Ốp/Lát',
                          'ĐVT', 'TỪ', 'ĐẾN', 'GIÁ BÁN', 'GIÁ CHÍNH SÁCH',
                          'GIÁ SÀN']

        # Đọc dữ liệu từ file Excel
        df = pd.read_excel(file_path, usecols=columns_needed)

        # Đổi tên cột 'ĐVT_1' thành 'ĐVT'
        df.rename(columns={'ĐVT_1': 'ĐVT'}, inplace=True)

        # Lặp qua từng nhóm công ty
        for (cong_ty, chi_nhanh), group in df.groupby(['Cong Ty Trien Khai', 'Ten Chi Nhanh']):
            # Nếu công ty không khớp với lựa chọn, bỏ qua
            if company != "TẤT CẢ" and cong_ty != company:
                continue

            # Loại bỏ ký tự xuống dòng trong tên chi nhánh
            chi_nhanh = chi_nhanh.replace('\n', ' ')  # Thay thế xuống dòng bằng khoảng trắng

            # Tạo tên file theo định dạng yêu cầu
            file_name = f"{cong_ty}_{chi_nhanh}_2024.xlsx"

            # Loại bỏ ký tự không hợp lệ trong tên file
            file_name = re.sub(r'[<>:"/\\|?*]', '_', file_name)

            # Xóa cột 'Cong Ty Trien Khai' và 'Ten Chi Nhanh'
            group = group.drop(columns=['Cong Ty Trien Khai', 'Ten Chi Nhanh'])

            # Tạo đường dẫn đầy đủ
            full_path = os.path.join(save_location, file_name)

            # Lưu nhóm dữ liệu vào file mới
            group.to_excel(full_path, index=False)

        QtWidgets.QMessageBox.information(None, "Success", "Data exported successfully!")


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec())
