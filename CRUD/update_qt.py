from PyQt6 import QtCore, QtGui, QtWidgets
import pandas as pd
import pyodbc

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 600)
        self.centralwidget = QtWidgets.QWidget(parent=MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(parent=self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(162, 270, 151, 41))
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(parent=self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(330, 270, 151, 41))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_3 = QtWidgets.QPushButton(parent=self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(500, 270, 151, 41))
        self.pushButton_3.setObjectName("pushButton_3")
        self.label = QtWidgets.QLabel(parent=self.centralwidget)
        self.label.setGeometry(QtCore.QRect(170, 150, 471, 31))
        self.label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(parent=self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(660, 520, 121, 21))
        self.label_2.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.label_2.setObjectName("label_2")
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

        # Kết nối nút bấm với phương thức xử lý
        self.pushButton.clicked.connect(self.updatePrice)
        self.pushButton_2.clicked.connect(self.addPrice)
        self.pushButton_3.clicked.connect(self.deletePrice)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.pushButton.setText(_translate("MainWindow", "Cập nhật giá bán"))
        self.pushButton_2.setText(_translate("MainWindow", "Thêm giá bán "))
        self.pushButton_3.setText(_translate("MainWindow", "Xóa giá bán"))
        self.label.setText(_translate("MainWindow", "<html><head/><body><p><span style=\" font-size:12pt; font-weight:600;\">PHẦN MỀM CHỈNH SỬA BẢNG GIÁ TỰ ĐỘNG</span></p></body></html>"))
        self.label_2.setText(_translate("MainWindow", "<html><head/><body><p><span style=\" font-size:6pt;\">Bản quyền Gia Bao</span></p></body></html>"))

    def showMessageBox(self, title, message):
        msg_box = QtWidgets.QMessageBox()
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setIcon(QtWidgets.QMessageBox.Icon.Information)  # Set icon to Information
        msg_box.setStandardButtons(QtWidgets.QMessageBox.StandardButton.Ok)  # OK button
        msg_box.exec()

    def updatePrice(self):
        file_path = 'C:/Users/Admin/Downloads/ketqua_combined.xlsx'
        df = pd.read_excel(file_path)

        # Chuyển đổi cột Mahang từ float64 sang string
        df['Mahang'] = df['Mahang'].astype(str)

        # Kiểm tra và chuyển đổi kiểu dữ liệu của cột GIA thành số nguyên
        df['GIA'] = pd.to_numeric(df['GIA'], errors='coerce')  # Chuyển đổi giá trị thành số hoặc NaN
        df = df.dropna(subset=['GIA'])  # Loại bỏ các dòng có giá trị NaN trong cột GIA
        df['GIA'] = df['GIA'].astype(int)  # Chuyển đổi các giá trị còn lại thành số nguyên

        # Kiểm tra kiểu dữ liệu của các cột
        print(df.dtypes)  # Để kiểm tra kiểu dữ liệu của từng cột

        conn_str = (
            r'DRIVER={SQL Server};'
            r'SERVER=XPS7590;'  # Thay đổi tên server của bạn
            r'DATABASE=UNIS_DATA;'  # Thay đổi tên cơ sở dữ liệu của bạn
            r'Trusted_Connection=yes;'  # Sử dụng Windows Authentication
        )

        try:
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()

            for index, row in df.iterrows():
                check_query = """
                SELECT COUNT(*)
                FROM Phantichgia_state1
                WHERE Macn = ? AND Mahang = ? AND Loai = ?
                """
                check_values = (str(row['Macn']), row['Mahang'], row['Loai'])
                cursor.execute(check_query, check_values)
                exists = cursor.fetchone()[0]

                if exists > 0:
                    sql_update_query = """
                    UPDATE Phantichgia_state1
                    SET GIA = ?, dvt = ?
                    WHERE Macn = ? AND Mahang = ? AND Loai = ?
                    """
                    values = (row['GIA'], row['dvt'], str(row['Macn']), row['Mahang'], row['Loai'])
                    try:
                        cursor.execute(sql_update_query, values)
                    except pyodbc.Error as update_error:
                        print(f"Error occurred during update: {update_error}")
                        continue

            conn.commit()
            self.showMessageBox("Cập nhật thành công", "Dữ liệu đã được cập nhật thành công!")

        except pyodbc.Error as e:
            print(f"Error occurred: {e}")
            self.showMessageBox("Lỗi", "Đã xảy ra lỗi khi cập nhật dữ liệu.")

        finally:
            cursor.close()
            conn.close()

    def addPrice(self):
        file_path = 'C:/Users/Admin/Downloads/ketqua_combined.xlsx'
        df = pd.read_excel(file_path, sheet_name='CombinedData')

        # Chỉ định kiểu dữ liệu cho từng cột
        df['Macn'] = df['Macn'].astype(str)
        df['Mahang'] = df['Mahang'].astype(str)
        df['dvt'] = df['dvt'].astype(str)
        df['Kichthuoc'] = df['Kichthuoc'].astype(str)
        df['Khunggiaban'] = df['Khunggiaban'].astype(str)
        df['Loai'] = df['Loai'].astype(str)

        # Chuyển đổi giá trị của cột GIA sang kiểu số nguyên
        try:
            df['GIA'] = pd.to_numeric(df['GIA'], errors='coerce').astype('Int64')
        except ValueError as e:
            print(f"Error converting 'GIA' column to numeric: {e}")
            self.showMessageBox("Lỗi", "Đã xảy ra lỗi khi chuyển đổi cột 'GIA'.")
            return

        conn_str = (
            r'DRIVER={SQL Server};'
            r'SERVER=XPS7590;'  # Thay đổi tên server của bạn
            r'DATABASE=UNIS_DATA;'  # Thay đổi tên cơ sở dữ liệu của bạn
            r'Trusted_Connection=yes;'  # Sử dụng Windows Authentication
        )

        try:
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()

            for index, row in df.iterrows():
                # Kiểm tra sự tồn tại của các dòng dựa trên các cột
                check_query = """
                SELECT COUNT(*)
                FROM Phantichgia_state1
                WHERE Macn = ? AND Mahang = ? AND dvt = ? AND Kichthuoc = ? AND Khunggiaban = ? AND Loai = ?
                """
                check_values = (row['Macn'], row['Mahang'], row['dvt'], row['Kichthuoc'], row['Khunggiaban'], row['Loai'])
                cursor.execute(check_query, check_values)
                exists = cursor.fetchone()[0]

                if exists == 0:
                    # Nếu dòng không tồn tại, thực hiện chèn mới
                    sql_insert_query = """
                    INSERT INTO Phantichgia_state1 (Macn, Mahang, dvt, Kichthuoc, Khunggiaban, Loai, GIA)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """
                    values = (row['Macn'], row['Mahang'], row['dvt'], row['Kichthuoc'], row['Khunggiaban'], row['Loai'], row['GIA'])
                    cursor.execute(sql_insert_query, values)

            conn.commit()
            self.showMessageBox("Thêm thành công", "Dữ liệu đã được thêm thành công!")

        except pyodbc.Error as e:
            print(f"Error occurred: {e}")
            self.showMessageBox("Lỗi", "Đã xảy ra lỗi khi thêm dữ liệu.")

        finally:
            cursor.close()
            conn.close()

    def deletePrice(self):
        # Cần implement phương thức xóa giá bán tương tự như updatePrice và addPrice
        pass


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec())
