
from PyQt5 import QtCore, QtGui
import pyodbc
import numpy as np
from PyQt5.QtWidgets import QTableView, QTableWidget, QTableWidgetItem
from PyQt5.QtSql import QSqlDatabase, QSqlQueryModel, QSqlQuery
from PyQt5.QtGui import QStandardItemModel, QStandardItem
from PyQt5.QtCore import Qt

import sys
from PyQt5.QtSql import QSqlDatabase, QSqlQueryModel, QSqlQuery

from PyQt5 import QtCore, QtGui, QtWidgets


msa_drivers = [x for x in pyodbc.drivers() if 'ACCESS' in x.upper()]
print(f'MS-ACCESS Drivers : {msa_drivers}')


def createConnection():
    con_String = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=C:\Users\ahmed\Desktop\filter.accdb;')

    global db
    db = QSqlDatabase.addDatabase('QODBC')

    db.setDatabaseName(con_String)

    if db.open():
        print('connect to DataBase Server successfully')

        return True
    else:
        print('connection failed')
        return False


conn = pyodbc.connect(
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\ahmed\Desktop\filter.accdb;')

cursor = conn.cursor()
fac = []
dep = []
sub = []
data = []
# test = []
# select * from (faculty inner join department on faculty.fid =department.fid ) inner join subject on subject.did= department.did
cursor.execute(
    f'select fname,dname,subname from (faculty inner join department on faculty.fid =department.fid ) inner join subject on subject.did= department.did')
for row in cursor.fetchall():
    # print(row[0])
    fac.append(row[0])
    # print(row[1])
    dep.append(row[1])
    # print(row[2])
    sub.append(row[2])
# print(sub)


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        self.y = []
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 600)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(210, 220, 47, 13))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(210, 260, 47, 13))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(210, 310, 47, 13))
        self.label_3.setObjectName("label_3")
        self.comboBox_3 = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox_3.setGeometry(QtCore.QRect(300, 300, 221, 22))
        self.comboBox_3.setObjectName("comboBox_3")
        # self.comboBox_3.addItem("")
        self.comboBox_3.addItem("")
        self.comboBox_4 = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox_4.setGeometry(QtCore.QRect(300, 250, 221, 22))
        self.comboBox_4.setObjectName("comboBox_4")
        self.comboBox_4.addItem("")
        # self.comboBox_4.addItem("")
        self.comboBox_5 = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox_5.setGeometry(QtCore.QRect(300, 210, 221, 22))
        self.comboBox_5.setObjectName("comboBox_5")
        self.comboBox_5.addItem("")
        self.comboBox_5.addItem("")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(340, 380, 75, 23))
        self.pushButton.setObjectName("pushButton")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.pushButton.clicked.connect(self.pressed)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label.setText(_translate("MainWindow", "الكلية"))
        self.label_2.setText(_translate("MainWindow", "القسم "))
        self.label_3.setText(_translate("MainWindow", "المواد"))
        x = -1
        for i in fac:
            self.comboBox_5.setItemText(x, _translate("MainWindow", i))
            x += 1
        self.comboBox_5.activated[str].connect(self.on_combobox_changed)
        self.comboBox_3.setItemText(0, _translate("MainWindow", ""))
        self.pushButton.setText(_translate("MainWindow", "بحث"))

    def on_combobox_changed(self, value):
        print(value)
        if value == fac[3]:

            for i in dep[2:5]:
                self.y.append(i)
                print(i)
                self.comboBox_4.addItem(i)
            for i in sub:
                self.comboBox_3.addItem(i)

        else:
            for i in dep[:2]:
                self.y.append(i)
                print(i)
                self.comboBox_4.addItem(i)

    def pressed(self):
        print(self.comboBox_3.currentText())

        print(self.comboBox_4.currentText())

        print(self.comboBox_5.currentText())
        qry = f"select distinct stu.name ,dname,fname from ((faculty inner join department on faculty.fid =department.fid ) inner join subject on subject.did= department.did) inner join stu on stu.facid =department.fid where department.dname='{self.comboBox_4.currentText()}'"
        self.w = displayData(qry)
        self.w.show()

        for row in cursor.fetchall():
            print(row)
            data.append(row)


def displayData(sqlStatement):
    print('processing query...')
    qry = QSqlQuery(db)
    qry.prepare(sqlStatement)
    qry.exec()
    model = QSqlQueryModel()
    model.setQuery(qry)
    view = QTableView()
    view.setMinimumSize(1250, 600)
    view.setModel(model)
    return view


if __name__ == "__main__":
    createConnection()
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
