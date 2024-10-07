import sys
import pandas as pd
import re
from PyQt6 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 250)
        self.centralwidget = QtWidgets.QWidget(parent=MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.label = QtWidgets.QLabel(parent=self.centralwidget)
        self.label.setGeometry(QtCore.QRect(200, 20, 411, 41))
        self.label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.label.setObjectName("label")

        self.label_2 = QtWidgets.QLabel(parent=self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(70, 70, 121, 21))
        self.label_2.setObjectName("label_2")

        self.label_3 = QtWidgets.QLabel(parent=self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(70, 100, 141, 21))
        self.label_3.setObjectName("label_3")

        self.lineEdit = QtWidgets.QLineEdit(parent=self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(210, 70, 391, 22))
        self.lineEdit.setObjectName("lineEdit")

        self.lineEdit_2 = QtWidgets.QLineEdit(parent=self.centralwidget)
        self.lineEdit_2.setGeometry(QtCore.QRect(210, 100, 391, 22))
        self.lineEdit_2.setObjectName("lineEdit_2")

        self.pushButton = QtWidgets.QPushButton(parent=self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(290, 130, 93, 28))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.process_data)  # Kết nối nút Save

        self.pushButton_2 = QtWidgets.QPushButton(parent=self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(400, 130, 93, 28))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(MainWindow.close)  # Kết nối nút Thoát

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(parent=MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(parent=MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Phần Mềm Chỉnh Sửa Bảng Giá Tự Động"))
        self.label.setText(_translate("MainWindow",
                                      "<html><head/><body><p><span style=\" font-size:10pt; font-weight:600;\">PHẦN MỀM CHỈNH SỬA BẢNG GIÁ TỰ ĐỘNG</span></p></body></html>"))
        self.label_2.setText(_translate("MainWindow",
                                        "<html><head/><body><p><span style=\" font-size:9pt;\">Link input chốt giá</span></p></body></html>"))
        self.label_3.setText(_translate("MainWindow",
                                        "<html><head/><body><p><span style=\" font-size:9pt;\">Link output chốt giá</span></p></body></html>"))
        self.pushButton.setText(_translate("MainWindow", "Lưu file"))
        self.pushButton_2.setText(_translate("MainWindow", "Thoát"))

    def process_data(self):
        input_path = self.lineEdit.text()
        output_path = self.lineEdit_2.text()

        # Đoạn code xử lý dữ liệu
        def process_sheet(sheet_name, pattern, id_vars):
            df = pd.read_excel(input_path, sheet_name=sheet_name, header=[0, 1])
            df.columns = ['_'.join(col).strip() for col in df.columns.values]
            print(f"Cột hiện tại trong '{sheet_name}':", df.columns)

            loai_cols = [col for col in df.columns if re.match(pattern, col)]

            for var in id_vars:
                if var not in df.columns:
                    print(f"Cột '{var}' không có trong DataFrame.")
                    return None

            melted_df = df.melt(id_vars=id_vars, value_vars=loai_cols, var_name='Loai', value_name='GIA')
            melted_df['Macn'] = melted_df['Loai'].apply(lambda x: x.split('_')[0])
            melted_df['Loai'] = melted_df['Loai'].apply(lambda x: x.split('_')[1])
            melted_df.rename(
                columns={id_vars[0]: 'Mahang', id_vars[1]: 'Kichthuoc', id_vars[2]: 'dvt', id_vars[3]: 'Khunggiaban'},
                inplace=True)

            melted_df = melted_df[['Macn', 'Mahang', 'Kichthuoc', 'dvt', 'Khunggiaban', 'Loai', 'GIA']]
            melted_df['Macn'] = melted_df['Macn'].str.lstrip('0').replace('', '0')

            return melted_df

        try:
            df_giabanle = process_sheet('giabanle', r'^\d{3}_GBLL1K',
                                        ['Mahang_Mahang', 'Kichthuoc_Kichthuoc', 'dvt_dvt', 'Khunggiaban_Khunggiaban'])
            df_giachinhsach = process_sheet('giachinhsach', r'^\d{3}_Giachinhsach',
                                            ['Mahang_Mahang', 'Kichthuoc_Kichthuoc', 'dvt_dvt',
                                             'Khunggiaban_Khunggiaban'])
            df_giasicont = process_sheet('giasicont', r'^\d{3}_Giasicont',
                                         ['Mahang_Mahang', 'Kichthuoc_Kichthuoc', 'dvt_dvt', 'Khunggiaban_Khunggiaban'])
            df_giaomkho200 = process_sheet('giaomkho200', r'^\d{3}_Giaomkho200',
                                           ['Mahang_Mahang', 'Kichthuoc_Kichthuoc', 'dvt_dvt',
                                            'Khunggiaban_Khunggiaban'])
            df_giacatlo = process_sheet('giacatlo', r'^\d{3}_Giacatlo',
                                        ['Mahang_Mahang', 'Kichthuoc_Kichthuoc', 'dvt_dvt', 'Khunggiaban_Khunggiaban'])

            combined_df = pd.concat([df_giabanle, df_giachinhsach, df_giasicont, df_giaomkho200, df_giacatlo],
                                    ignore_index=True)
            combined_df.to_excel(output_path, sheet_name='CombinedData', index=False)
            print("Dữ liệu đã được lưu vào file Excel mới.")
            QtWidgets.QMessageBox.information(None, "Thông báo", "Dữ liệu đã được lưu thành công!")

        except Exception as e:
            QtWidgets.QMessageBox.critical(None, "Lỗi", str(e))


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec())
