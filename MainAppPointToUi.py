# -*- coding: utf-8 -*-
"""
Created September 2015
@author: wcz5
"""
import os
import sys
from PyQt4 import QtCore, QtGui, uic

import re
#from PyQt4.QtCore import *
#from PyQt4.QtGui import *
from PyQt4.QtSql import *

#get user login in
#user = os.environ['USER'] 
user = os.getlogin()
#usr = pwd.getpwuid(os.getuid()).pw_name

if user == 'anwar63':
    user = 'anwar'

pcName = os.path.expanduser('~') 
#os.environ['HOME']
#print(pcName)

'''
Look into
http://www.pytables.org/
http://www.hdfgroup.org/HDF5/
http://vitables.org/
'''

#valDataPath = 'C:\\Users\\' + usr + '\\OneDrive\\_AnwarPy\\Practice\\valData.py'
#sys.path.append(os.path.abspath(valDataPath))
#import importlib
#importlib.import_module(valDataPath, None)

# this is a custom file for validating data entered by users
from valData import *
from DataModels import *

#from PyQt4.QtSql import QSqlQueryModel,QSqlDatabase,QSqlQuery

form_class = uic.loadUiType("MainApp.ui")[0]                 # Load the UI

import pyodbc
#import pypyodbc



class MyWindowClass(QtGui.QMainWindow, form_class):
    def __init__(self, dbAccessConn):
        QtGui.QMainWindow.__init__(self)
        self.setupUi(self)

        self.dbAccessConn = dbAccessConn
        self.cboSalesPeople.addItems(self.getSalesPeople())

        #self.cboSalesPeople.activated.connect(self.getSalesPeople)

        self.cboSalesPeople.currentIndexChanged.connect(self.updateView)
        self.vwSales.resizeColumnsToContents()
        self.btnAddToTable.clicked.connect(self.btn_AddToTable_clicked)
        self.btnClose.clicked.connect(self.closeApp) 
        self.actionClose.triggered.connect(self.closeApp)
        
        self.lnSalesPerson.textEdited.connect(self.resetControls)
        self.lnSalesMonth.textEdited.connect(self.resetControls)
        self.lnSalesAmount.textEdited.connect(self.resetControls)

        self.defineTable()

    def closeApp(self):
        myWindow.close() 

    def resetControls(self):
        self.txtGeneral.setText("")

    def defineTable(self):
        
        self.tblSalesData.setColumnCount(3)
        self.tblSalesData.setHorizontalHeaderItem(0, QtGui.QTableWidgetItem("SalesPerson"))
        self.tblSalesData.setColumnWidth(0, 100)
        self.tblSalesData.setHorizontalHeaderItem(1, QtGui.QTableWidgetItem("Month"))
        self.tblSalesData.setColumnWidth(1, 50)
        self.tblSalesData.setHorizontalHeaderItem(2, QtGui.QTableWidgetItem("Amount"))
        self.tblSalesData.setColumnWidth(2, 50)
        
        #self.tblSalesData.width = 905
        #self.tblSalesData.height=1000
        
        self.tblSalesData.setItem(1, 1, QtGui.QTableWidgetItem('test'))
        
        curs = dbAccessConn.cursor()

        rowcount = curs.execute('''SELECT COUNT(*) FROM SalesData''').fetchone()[0]
        self.tblSalesData.setRowCount(rowcount)
        curs.execute('''SELECT * FROM SalesData''')
        for row, form in enumerate(curs):
            for column, item in enumerate(form):
                self.tblSalesData.setItem(row, column, QtGui.QTableWidgetItem(str(item)))           


        #self.tblSalesData.cellDoubleClicked.connect(self.double_clicked)


    def btn_AddToTable_clicked(self):
        
        person = self.lnSalesPerson.text()
        if len(person) < 1:
            self.txtGeneral.setText("Name cannot be empty: please try again!")
            return
            
        month = self.lnSalesMonth.text()
        if IsMonthValid(month) == False:
            self.txtGeneral.setText(month + " is not a valid month. Please enter a value between 1 and 12!")
            return
        
        amount = self.lnSalesAmount.text()
        if IsFloatValid(amount) == False:
            self.txtGeneral.setText(amount + " is not a valid amount. Please try again!")
            return
        
        curs = dbAccessConn.cursor()
        curs.execute ("INSERT INTO SalesData (SalesPerson, mon, amount) VALUES (?, ?, ?) ", person, month, amount)
        curs.commit()
        curs.close()

        self.cboSalesPeople.addItems(self.getSalesPeople())
        self.txtGeneral.setText(' New record added to table: Person: ' + person + ' month: ' + month + ' amount ' + amount)

        self.defineTable()

    def getSalesPeople(self):
        self.cboSalesPeople.clear()
        salesPeopleList=['',] #start with a blank one so the user can choose
        curs = self.dbAccessConn.cursor()
        sql = "select distinct SalesPerson from SalesData"
        curs.execute(sql)
        for row in curs:
            salesPeopleList.append(row.SalesPerson)
        curs.close()
        return  (salesPeopleList)

    def updateView(self):
        curs = dbAccessConn.cursor()
        person = str(self.cboSalesPeople.currentText())

        if len(person)> 0 :
            #Parameterized with ?
            sql = "select SalesPerson, mon, str(amount) from SalesData where SalesPerson = ? order by mon"
            curs.execute(sql, (person,) )

            rows = curs.fetchall()
            header = ['SalesPerson', 'month', 'amount']
            data = []
            for row in rows:
                #data.append('(' + str(row.SalesPerson) + ',' + str(row.mon) + ',' + str(row.amount) + ')')
                data.append(row)

            curs.close()
            tablemodel = SalesDataModel(self, data, header)
            self.vwSales.setModel(tablemodel)


            '''
            db = QSqlDatabase.addDatabase("Microsoft Access")
            db.setDatabaseName(dbAccess)
            db.open()
            projectModel = QSqlQueryModel()
            projectModel.setQuery("select * from SalesData",db)
            self.vwSalesModel.setModel(projectModel)
            '''


if pcName == '/home/anwar63':
    AccessDB = '/media/anwar63/TI10713100F/Users/Anwar/Google Drive/_AnwarPy/SQLExamples/plotData.mdb'
else:
    print('You are user: ' + user)
    print('C:\\Users\\' + user + '\\Google Drive\\_AnwarPy\\SQLExamples\\plotData.mdb')
    AccessDB = 'C:\\Users\\' + user + '\\Google Drive\\_AnwarPy\\SQLExamples\\plotData.mdb'

dbAccess = ('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' %AccessDB )
dbAccessConn = pyodbc.connect(dbAccess)

app = QtGui.QApplication(sys.argv)
myWindow = MyWindowClass(dbAccessConn)
myWindow.show()
#app.exec_()
sys.exit(app.exec_())
