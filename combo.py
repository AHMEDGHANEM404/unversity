
from PyQt5 import QtCore, QtGui
import pyodbc
import numpy as np
from PyQt5.QtWidgets import QTableView, QTableWidget, QTableWidgetItem
from PyQt5.QtSql import QSqlDatabase, QSqlQueryModel, QSqlQuery
from PyQt5.QtGui import QIcon, QStandardItemModel, QStandardItem
from PyQt5.QtCore import Qt

import sys
from PyQt5.QtSql import QSqlDatabase, QSqlQueryModel, QSqlQuery

from PyQt5 import QtCore, QtGui, QtWidgets


msa_drivers = [x for x in pyodbc.drivers() if 'ACCESS' in x.upper()]
print(f'MS-ACCESS Drivers : {msa_drivers}')


def createConnection():
    con_String = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=E:\project\year\code\exe project\project v1.0\v1\filter.accdb;')

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
        r'DBQ=E:\project\year\code\exe project\project v1.0\v1\filter.accdb;')

cursor = conn.cursor()
fac = []
dep = []
sub = []
data = []
level = []
# select * from (faculty inner join department on faculty.fid =department.fid ) inner join subject on subject.did= department.did
cursor.execute(
       f'select  fname,dname,subname ,levels.name  from ((faculty inner join department on faculty.fid =department.fid ) inner join subject on subject.did= department.did)inner join levels on levels.ID =subject.levelid order by levels.ID')
for row in cursor.fetchall():
    # print(row[0])
    fac.append(row[0])
    # print(row[1])
    dep.append(row[1])
    # print(row[2])
    sub.append(row[2])
    print(row[3])
    level.append(row[3])


    
# print(fac)
# print(dep)
# print(sub)
print(level)



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
        # ظ================================================
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(210, 350, 47, 13))
        self.label_4.setObjectName("label_4")
        self.comboBox_3 = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox_3.setGeometry(QtCore.QRect(300, 300, 221, 22))
        self.comboBox_3.setObjectName("comboBox_3")
        self.comboBox_3.addItem("")
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

        # =======================================================
        self.comboBox_6 = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox_6.setGeometry(QtCore.QRect(300, 350, 221, 22))
        self.comboBox_6.setObjectName("comboBox_6")
        self.comboBox_6.addItem("")
        self.comboBox_6.addItem("")
        self.comboBox_6.addItem("")
        self.comboBox_6.addItem("")
        # =====================================================
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
        MainWindow.setWindowTitle(_translate("MainWindow", "جامعة الازهر"))
        self.label.setText(_translate("MainWindow", "الكلية"))
        self.label_2.setText(_translate("MainWindow", "القسم "))
        self.label_3.setText(_translate("MainWindow", "المواد"))
        self.label_4.setText(_translate("MainWindow", "الفرقة"))

        x=0
        for i in fac[2:]:
            self.comboBox_5.setItemText(x, _translate("MainWindow", i))
            x += 1
        self.comboBox_5.activated[str].connect(self.on_combobox_changed)
        self.comboBox_3.setItemText(0, _translate("MainWindow", ""))
        self.pushButton.setText(_translate("MainWindow", "بحث"))
        
    def on_combobox_changed(self, value):
        print(value)
        if value == fac[1]:
            self.comboBox_4.clear()
            for i in dep[:3]:
                self.y.append(i)
                print(i)
                self.comboBox_4.addItem(i)
            self.comboBox_3.clear()
            for i in sub[:3]:
                self.comboBox_3.addItem(i)
            self.comboBox_6.clear()
            
            for i in level[:3]:
                self.comboBox_6.addItem(i)
                

        
        if value==fac[3]:
            self.comboBox_4.clear()
            self.y.append(dep[3])
            print(dep[3])
            self.comboBox_4.addItem(dep[3])
            self.comboBox_3.clear()
            self.comboBox_3.addItem(sub[3])
            self.comboBox_6.clear()
            self.comboBox_6.addItem(level[3])
            

    def pressed(self):
        print(self.comboBox_3.currentText())

        print(self.comboBox_4.currentText())

        print(self.comboBox_5.currentText())

        print(self.comboBox_6.currentText())


        qry = f"select distinct stu.name ,dname,fname from ((faculty inner join department on faculty.fid =department.fid ) inner join subject on subject.did= department.did) inner join stu on stu.facid =department.fid where department.dname='{self.comboBox_4.currentText()}' and subject.subname ='{self.comboBox_3.currentText()}'"
       
        self.w = displayData(qry)
        self.w.show()

        # for row in cursor.fetchall():
        #     print(row)
        #     data.append(row)


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
    app.setWindowIcon(QIcon("logo.jpg"))
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
