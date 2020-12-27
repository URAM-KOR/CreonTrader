import sys
import win32com.client
from PyQt5.QtWidgets import *
import pandas as pd
import os
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from pywinauto import application
from slacker import Slacker


# 크레온 플러스 공통 OBJECT
cpStatus = win32com.client.Dispatch("CpUtil.CpCybos")
cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')
objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")


import win32com.client



objStockMst.SetInputValue(0, 'A005930')  # 종목 코드 - 삼성전자
objStockMst.BlockRequest()

code = objStockMst.GetHeaderValue(0)  # 종목코드
name = objStockMst.GetHeaderValue(1)  # 종목명
offer = objStockMst.GetHeaderValue(16)  # 매도호가

#slack = Slacker('xoxb-1584959946071-1623576410048-8sVUKnpcl2hdROptLIV4LSgB')


# Send a message to #general channel
#slack.chat.post_message('#stock', '삼성전자 현재가: ' + str(offer))


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyStock")
        self.setGeometry(300, 300, 600, 450)

        btn1 = QPushButton("Login", self)
        btn1.move(20, 350)
        btn1.clicked.connect(self.btn1_clicked)

        btn2 = QPushButton("Check state", self)
        btn2.move(20, 400)
        btn2.clicked.connect(self.btn2_clicked)




        self.text_edit = QTextEdit(self)
        self.text_edit.setGeometry(200, 60, 250, 300)
        self.text_edit.setEnabled(False)

        # 로그인 확인
        btn2.clicked.connect(self.event_connect)


        label = QLabel('종목코드: ', self)
        label.move(200,20)

        self.code_edit = QLineEdit(self)
        self.code_edit.move(250, 20)
        self.code_edit.setText(code)


        btn3 = QPushButton("조회", self)
        btn3.move(360, 20)
        btn3.clicked.connect(self.btn3_clicked)

        btn4 = QPushButton("종목검색", self)
        btn4.move(20, 200)
       # btn4.clicked.connect(self.btn4_clicked)


    def closeEvent(self, QCloseEvent):
        ans = QMessageBox.question(self, "프로그램 종료", "종료하시겠습니까?",
                             QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ans == QMessageBox.Yes:
            QCloseEvent.accept()
        else:
            QCloseEvent.ignore()




    def btn1_clicked(self):
        QMessageBox.about(self, "message", "크레온 플러스 실행...")

        application.Application().start(
            "C:/CREON/STARTER/coStarter.exe    /prj:cp /id:##### /pwd:#####/pwdcert:##### /autostart")


    def btn2_clicked(self): #연결 확인이벤트
        if cpStatus == 0:
            self.statusBar().showMessage("PLUS가 정상적으로 연결되지 않음. ")
        else:
            self.statusBar().showMessage("연결 성공")

    def event_connect(self, cpStatus): #로그인 이벤트
        if cpStatus == 0:
            self.text_edit.append("로그인 성공")


    def btn3_clicked(self):
        code = self.code_edit.text()

        self.text_edit.append("종목코드: " + code)
        self.text_edit.append("종목명: " + name)
        self.text_edit.append("종목명: " + str(offer))



    def btn4_clicked(StockFinder):

        ret = self.kiwoom.dynamicCall("GetCodeListByMarket(QString)", ["0"])
        kospi_code_list = ret.split(';')
        kospi_code_name_list = []

        for x in kospi_code_list:
            name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", [x])
            kospi_code_name_list.append(x + " : " + name)

        self.listWidget.addItems(listAllStrategy)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    mywindow = MyWindow()
    mywindow.show()
    app.exec_()
 
