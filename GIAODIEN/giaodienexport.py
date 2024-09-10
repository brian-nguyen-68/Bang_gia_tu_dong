import pandas as pd
import pyodbc
import os
from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtWidgets import QFileDialog, QMessageBox


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(836, 366)
        self.centralwidget = QtWidgets.QWidget(parent=MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(parent=self.centralwidget)
        self.label.setGeometry(QtCore.QRect(130, 100, 91, 16))
        self.label.setObjectName("label")
        self.lineEdit = QtWidgets.QLineEdit(parent=self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(230, 100, 391, 22))
        self.lineEdit.setObjectName("lineEdit")
        self.pushButton = QtWidgets.QPushButton(parent=self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(380, 140, 93, 28))
        self.pushButton.setObjectName("pushButton")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(parent=MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 836, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(parent=MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        # Kết nối nút bấm với hàm xuất dữ liệu
        self.pushButton.clicked.connect(self.export_data)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label.setText(_translate("MainWindow", "Nhập thư mục"))
        self.pushButton.setText(_translate("MainWindow", "Xuất dữ liệu"))

    def export_data(self):
        # Lấy đường dẫn từ lineEdit
        output_directory = self.lineEdit.text()

        if not output_directory:
            self.show_message("Lỗi", "Vui lòng nhập đường dẫn thư mục.")
            return

        # Kiểm tra nếu thư mục không tồn tại thì thông báo lỗi
        if not os.path.exists(output_directory):
            self.show_message("Lỗi", "Thư mục không tồn tại.")
            return

        # Kết nối đến SQL Server
        conn_str = (
            r'DRIVER={ODBC Driver 17 for SQL Server};'
            r'SERVER=XPS7590;'  # Thay thế với tên máy chủ của bạn
            r'DATABASE=UNIS_DATA;'  # Thay thế với tên cơ sở dữ liệu của bạn
            r'TRUSTED_CONNECTION=yes;'
        )
        conn = pyodbc.connect(conn_str)

        # Đọc dữ liệu từ bảng
        query = "SELECT * FROM Phantichgia_state1"
        df = pd.read_sql(query, conn)

        # Đóng kết nối
        conn.close()

        # Lấy danh sách các mã chi nhánh
        branch_codes = df['Macn'].unique()

        for branch_code in branch_codes:
            # Lọc dữ liệu theo mã chi nhánh
            branch_df = df[df['Macn'] == branch_code]

            # Đặt tên file dựa trên mã chi nhánh
            file_name = f'{branch_code}.xlsx'

            # Tạo đường dẫn đầy đủ tới file
            file_path = os.path.join(output_directory, file_name)

            # Xuất dữ liệu ra file Excel
            branch_df.to_excel(file_path, index=False, engine='openpyxl')

            print(f'Data for branch code {branch_code} exported to {file_path}')

        self.show_message("Thông báo", "Dữ liệu đã được xuất thành công.")

    def show_message(self, title, message):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.exec()


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec())
