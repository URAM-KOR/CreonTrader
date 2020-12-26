import sys
import win32com.client
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from pywinauto import application

# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyStock")
        self.setGeometry(300, 300, 300, 150)

        btn1 = QPushButton("Login", self)
        btn1.move(20, 20)
        btn1.clicked.connect(self.btn1_clicked)

        btn2 = QPushButton("Check state", self)
        btn2.move(20, 70)
        btn2.clicked.connect(self.btn2_clicked)

    def btn1_clicked(self):
        QMessageBox.about(self, "message", "크레온 플러스 실행...")

        application.Application().start(
            "C:/CREON/STARTER/coStarter.exe /prj:cp /id:크레온아이디 /pwd:크레온비밀번호 /pwdcert:공인인증서비밀번호 /autostart")
        exit()

    def btn2_clicked(self):
        if bConnect == 0:
            self.statusBar().showMessage("PLUS가 정상적으로 연결되지 않음. ")
        else:
            self.statusBar().showMessage("연결 성공")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mywindow = MyWindow()
    mywindow.show()
    app.exec_()
