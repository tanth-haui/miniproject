# GUI.py
from PyQt5.QtCore import (QCoreApplication, QMetaObject, QRect, QSize)
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import *

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(512, 456)
        MainWindow.setMinimumSize(QSize(512, 456))
        MainWindow.setMaximumSize(QSize(512, 456))
        MainWindow.setStyleSheet(u"background-color: rgb(91, 155, 213);")
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.centralwidget.setMinimumSize(QSize(512, 0))

        font = QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setUnderline(True)
        font.setWeight(75)

        font1 = QFont()
        font1.setPointSize(17)
        font1.setBold(True)
        font1.setWeight(75)

        font2 = QFont()
        font2.setPointSize(10)
        font2.setBold(True)
        font2.setWeight(75)

        font3 = QFont()
        font3.setPointSize(9)
        font3.setBold(True)
        font3.setWeight(75)

        font4 = QFont()
        font4.setFamily("Segoe Print")
        font4.setPointSize(10)
        font4.setBold(True)
        font4.setWeight(75)

        self.label = QLabel(self.centralwidget)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(180, 30, 121, 21))
        self.label.setFont(font)

        self.label_2 = QLabel(self.centralwidget)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setGeometry(QRect(170, 60, 151, 21))
        self.label_2.setFont(font1)

        self.label_3 = QLabel(self.centralwidget)
        self.label_3.setObjectName(u"label_3")
        self.label_3.setGeometry(QRect(50, 90, 81, 21))
        self.label_3.setFont(font2)

        self.Input_Folder = QLineEdit(self.centralwidget)
        self.Input_Folder.setObjectName(u"Input_Folder")
        self.Input_Folder.setGeometry(QRect(40, 120, 331, 31))
        self.Input_Folder.setStyleSheet(u"background-color: rgb(201, 201, 201);")

        self.label_4 = QLabel(self.centralwidget)
        self.label_4.setObjectName(u"label_4")
        self.label_4.setGeometry(QRect(50, 176, 91, 20))
        self.label_4.setFont(font2)

        self.Output_Folder = QLineEdit(self.centralwidget)
        self.Output_Folder.setObjectName(u"Output_Folder")
        self.Output_Folder.setGeometry(QRect(40, 200, 331, 31))
        self.Output_Folder.setStyleSheet(u"background-color: rgb(201, 201, 201);")

        self.label_5 = QLabel(self.centralwidget)
        self.label_5.setObjectName(u"label_5")
        self.label_5.setGeometry(QRect(50, 260, 71, 16))
        self.label_5.setFont(font2)

        self.label_6 = QLabel(self.centralwidget)
        self.label_6.setObjectName(u"label_6")
        self.label_6.setGeometry(QRect(250, 260, 61, 16))
        self.label_6.setFont(font2)

        self.Input_Start_Time = QLineEdit(self.centralwidget)
        self.Input_Start_Time.setObjectName(u"Input_Start_Time")
        self.Input_Start_Time.setGeometry(QRect(40, 280, 131, 31))
        self.Input_Start_Time.setStyleSheet(u"background-color: rgb(201, 201, 201);")

        self.Input_End_Time = QLineEdit(self.centralwidget)
        self.Input_End_Time.setObjectName(u"Input_End_Time")
        self.Input_End_Time.setGeometry(QRect(240, 280, 131, 31))
        self.Input_End_Time.setStyleSheet(u"background-color: rgb(201, 201, 201);")

        self.label_7 = QLabel(self.centralwidget)
        self.label_7.setObjectName(u"label_7")
        self.label_7.setGeometry(QRect(50, 340, 51, 21))
        self.label_7.setFont(font2)


        self.label_8 = QLabel(self.centralwidget)
        self.label_8.setObjectName(u"label_8")
        self.label_8.setGeometry(QRect(380, 410, 121, 31))
        self.label_8.setFont(font4)

        self.Signal = QComboBox(self.centralwidget)
        self.Signal.addItems(["Actual Speed", "Set Speed", "Feed Forward", "AC Switch"])
        self.Signal.setObjectName(u"Signal")
        self.Signal.setGeometry(QRect(40, 370, 131, 31))
        self.Signal.setStyleSheet(u"background-color: rgb(201, 201, 201);")

        self.Input_button = QPushButton(self.centralwidget)
        self.Input_button.setObjectName(u"Input_button")
        self.Input_button.setGeometry(QRect(400, 120, 81, 31))
        self.Input_button.setFont(font2)
        self.Input_button.setStyleSheet(u"background-color: rgb(201, 201, 201);")

        self.Output_button = QPushButton(self.centralwidget)
        self.Output_button.setObjectName(u"Output_button")
        self.Output_button.setGeometry(QRect(400, 200, 81, 31))
        self.Output_button.setFont(font3)
        self.Output_button.setStyleSheet(u"background-color: rgb(201, 201, 201);")

        self.Start_button = QPushButton(self.centralwidget)
        self.Start_button.setObjectName(u"Start_button")
        self.Start_button.setGeometry(QRect(400, 370, 81, 31))
        self.Start_button.setFont(font3)
        self.Start_button.setStyleSheet(u"background-color: rgb(201, 201, 201);")

        MainWindow.setCentralWidget(self.centralwidget)

        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle("MainWindow")
        self.label.setText("Python Course")
        self.label_2.setText("Mini Project")
        self.label_3.setText("Input Folder")
        self.label_4.setText("Output Folder")
        self.label_5.setText("Start Time")
        self.label_6.setText("End Time")
        self.label_7.setText("Signal")
        self.Input_button.setText("Select")
        self.Output_button.setText("Select")
        self.Start_button.setText("Start")
        self.label_8.setText("Design by: TânCN")