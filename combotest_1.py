from PyQt5 import QtGui
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QDialog, QGridLayout, QLabel, QPushButton, QTableView, QTableWidget, QTableWidgetItem
import sys

import pyodbc
edit = []
names = []
num1 = []
num2 = []
num3 = []
num4 = []
num5 = []

num6 = []
# print(num6)


class Dlg(QDialog):

    def __init__(self, nb_row, nb_col):
        QDialog.__init__(self)
        self.layout = QGridLayout(self)
        self.label1 = QLabel('النتيجة ')
        self.btn = QPushButton('تم ')
        self.btn.setFixedWidth(100)
        self.nb_row = nb_row
        self.nb_col = nb_col

        # creating empty table
        # data = [[] for i in range(self.nb_row)]
        # # self.querys("name",names)

        # for i in range(self.nb_row):
        #     for j in range(self.nb_col):
        #         data[i].append('')
        list = []
        conn = pyodbc.connect(
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r'DBQ=E:\project\year\code\exe project\project v1.0\v1\filter.accdb;')
        cursor = conn.cursor()

        sql = f'''select * from test'''
        cursor.execute(sql)
        
        for row in cursor.fetchall():
            # print(row)
            list.append(row)
            # print(row[0])
            # ALL columns value
            num1.append(row[0])
            num2.append(row[1])
            num3.append(row[2])
            num4.append(row[3])
            num5.append(row[4])
            num6.append(row[5])

        self.table1 = QTableWidget()

        self.table1.setRowCount(self.nb_row)
        self.table1.setColumnCount(self.nb_col)

        self.table1.setHorizontalHeaderLabels(
            ['Id', 'name', "s1", "s2", "s3", "s4"])
        for row in range(self.nb_row):
            for col in range(self.nb_col):
                item = QTableWidgetItem(str(list[row][col]))
                self.table1.setItem(row, col, item)

        self.layout.addWidget(self.label1, 0, 0)
        self.layout.addWidget(self.table1, 1, 0)
        self.layout.addWidget(self.btn, 2, 0)

        self.btn.clicked.connect(self.print_table_values)

    def print_table_values(self, data):

        for row in range(self.nb_row):
            for col in range(self.nb_col):
                print(row, col, self.table1.item(row, col).text())
                edit.append(self.table1.item(row, col).text())
        print(edit)

   
        #  code of ===>>>>>>> update
        conn = pyodbc.connect(
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r'DBQ=E:\project\year\code\exe project\project v1.0\v1\filter.accdb;')
        cursor = conn.cursor()
        start=0
        end=6
        while end < len(edit):
            x=edit[start:end]
            sql=f'UPDATE test SET s1={int(x[2])},s2={int(x[3])} ,s3={int(x[4])}, s4={int(x[5])} where ID={int(x[0])}'
            cursor.execute(sql)
            conn.commit()
            start+=6
            end+=6
        # print("edit 1")
       
            
            

app = QApplication(sys.argv)
w = Dlg(5, 6)
w.resize(750, 300)
w.setWindowTitle('الازهر')
w.setWindowFlags(Qt.WindowStaysOnTopHint)
w.show()
app.exit(w.exec_())
