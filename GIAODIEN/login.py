from PyQt6 import QtCore, QtGui, QtWidgets
import pyodbc
import subprocess

class LoginWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_LoginWindow()
        self.ui.setupUi(self)
        self.ui.pushButtonLogin.clicked.connect(self.handle_login)
        self.ui.labelError.setText("")  # Clear error label initially

    def handle_login(self):
        username = self.ui.lineEditUsername.text()
        password = self.ui.lineEditPassword.text()

        # Kiểm tra thông tin đăng nhập
        if self.check_credentials(username, password):
            self.showMessageBox("Đăng nhập thành công", "Bạn đã đăng nhập thành công!")
            self.open_main_window()  # Mở file Python mới
        else:
            self.showMessageBox("Lỗi đăng nhập", "Tên đăng nhập hoặc mật khẩu không đúng.")

    def check_credentials(self, username, password):
        conn_str = (
            r'DRIVER={SQL Server};'
            r'SERVER=XPS7590;'  # Thay đổi tên server của bạn
            r'DATABASE=UNIS_DATA;'  # Thay đổi tên cơ sở dữ liệu của bạn
            r'UID=' + username + ';'  # Tên đăng nhập SQL Server
            r'PWD=' + password + ';'  # Mật khẩu
        )
        try:
            conn = pyodbc.connect(conn_str)
            return True  # Kết nối thành công
        except pyodbc.Error as e:
            print(f"Error occurred: {e}")
            return False
        finally:
            if 'conn' in locals():
                conn.close()

    def showMessageBox(self, title, message):
        msg_box = QtWidgets.QMessageBox(self)
        msg_box.setIcon(QtWidgets.QMessageBox.Icon.Information if "thành công" in title else QtWidgets.QMessageBox.Icon.Critical)
        msg_box.setText(message)
        msg_box.setWindowTitle(title)
        msg_box.exec()

    def open_main_window(self):
        # Mở file Python khác
        subprocess.Popen(['python', 'C:/Users/Admin/PycharmProjects/banggia/GIAODIEN/giaodien_crud1.py'])  # Thay đổi đường dẫn đến file Python của bạn
        self.close()  # Đóng cửa sổ đăng nhập

class Ui_LoginWindow(object):
    def setupUi(self, LoginWindow):
        LoginWindow.setObjectName("LoginWindow")
        LoginWindow.resize(300, 200)
        self.centralwidget = QtWidgets.QWidget(parent=LoginWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")

        self.labelUsername = QtWidgets.QLabel(parent=self.centralwidget)
        self.labelUsername.setObjectName("labelUsername")
        self.gridLayout.addWidget(self.labelUsername, 0, 0, 1, 1)

        self.lineEditUsername = QtWidgets.QLineEdit(parent=self.centralwidget)
        self.lineEditUsername.setObjectName("lineEditUsername")
        self.gridLayout.addWidget(self.lineEditUsername, 0, 1, 1, 1)

        self.labelPassword = QtWidgets.QLabel(parent=self.centralwidget)
        self.labelPassword.setObjectName("labelPassword")
        self.gridLayout.addWidget(self.labelPassword, 1, 0, 1, 1)

        self.lineEditPassword = QtWidgets.QLineEdit(parent=self.centralwidget)
        self.lineEditPassword.setObjectName("lineEditPassword")
        self.lineEditPassword.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
        self.gridLayout.addWidget(self.lineEditPassword, 1, 1, 1, 1)

        self.pushButtonLogin = QtWidgets.QPushButton(parent=self.centralwidget)
        self.pushButtonLogin.setObjectName("pushButtonLogin")
        self.gridLayout.addWidget(self.pushButtonLogin, 2, 0, 1, 2)

        self.labelError = QtWidgets.QLabel(parent=self.centralwidget)
        self.labelError.setObjectName("labelError")
        self.gridLayout.addWidget(self.labelError, 3, 0, 1, 2)

        LoginWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(LoginWindow)
        QtCore.QMetaObject.connectSlotsByName(LoginWindow)

    def retranslateUi(self, LoginWindow):
        _translate = QtCore.QCoreApplication.translate
        LoginWindow.setWindowTitle(_translate("LoginWindow", "Login"))
        self.labelUsername.setText(_translate("LoginWindow", "Tên đăng nhập:"))
        self.labelPassword.setText(_translate("LoginWindow", "Mật khẩu:"))
        self.pushButtonLogin.setText(_translate("LoginWindow", "Đăng nhập"))
        self.labelError.setText("")

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    login_window = LoginWindow()
    login_window.show()
    sys.exit(app.exec())
