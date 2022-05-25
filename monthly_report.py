from distutils.command.build_scripts import first_line_re
import os
import datetime
import sys
import json
from unicodedata import name
import PyQt5
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QWidget, QPushButton, QFileDialog
from PyQt5.QtGui import QIcon
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pkg_resources import working_set

class App(QMainWindow):
    
    # !     Important
    # TODO  ToDo
    # *     Note
    # ?     Logic
    
    
    def __init__(self):
        super().__init__()
        self.title = 'Python Test App'
        self.left = 100
        self.top = 100
        self.width = 300
        self.height = 150
        self.setStyleSheet("background-color:#363636 ")
        
        
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.statusBar().showMessage('Select master excel')
        self.setFixedSize(450,250)
        
        # Browse button
        button = QPushButton('browse',self)
        button.setToolTip('select excel file')
        button.move(20,180)
        button.clicked.connect(self.openFileNameDialog)
        
        # Qlabel name 
        person_name = QLabel(self)
        person_name.move(20,4)
        person_name.setText('Full Name')
        person_name.setStyleSheet("color: #cfcfcf;"
                                   )
        
        #Qlabel account name
        user_account = QLabel(self)
        user_account.move(140,4)
        user_account.setText('Account')
        user_account.setStyleSheet("color: #cfcfcf;"
                                   )
        
        # Create textbox full_name
        self.full_name = QLineEdit(self)
        self.full_name.move(20, 30)
        self.full_name.resize(100,20)
        self.full_name.setStyleSheet(""" 
                                     background-color: grey
                                     """)
        
        # Create textbox user_account
        self.user_account = QLineEdit(self)
        self.user_account.move(140, 30)
        self.user_account.resize(100,20)
        
        # Create textbox first date
        self.first_date = QLineEdit(self)
        self.first_date.move(20, 100)
        self.first_date.resize(100,20)
        
        
        # Create textbox last date
        self.last_date = QLineEdit(self)
        self.last_date.move(140, 100)
        self.last_date.resize(100,20)
        
        #QLabel first date
        first_date = QLabel(self)
        first_date.move(20, 75)
        first_date.setText('First Date')
        first_date.setStyleSheet("""
                                 color: #cfcfcf;
                                 """)
        
        #QLabel last date
        last_date = QLabel(self)
        last_date.move(140, 75)
        last_date.setText('Last Date')
        last_date.setStyleSheet("""
                                color: #cfcfcf;
                                
                                """)
        
        #Test Button
        test_button = QPushButton('Test Button', self)
        test_button.move(140,180)
        test_button.clicked.connect(self.testButton)
        test_button.setStyleSheet("")
        
        self.show()
    
    #test_button
    def testButton(self):
        print(self.full_name.text(),
              self.first_date.text(),
              self.last_date.text(),
              self.user_account.text()
              )
    
    
    #Open File Dialog Window
    def openFileNameDialog(self):
        
        full_name = self.full_name.text()
        first_date = self.first_date.text()
        last_date = self.last_date.text()
        user_account = self.user_account.text()
        sheet_name = first_date+"-"+last_date
        
        desktop = os.path.expanduser("~\Desktop\\") #path for current user desktop
        
        fileName, _ = QFileDialog.getOpenFileName(self,"Select Excel", "","Excel Files(*.xlsx)") #Grab File
        
        destination_workbook = Workbook() #inst destination_workbook
        destination_workbook.save(desktop+full_name+'.xlsx') #create destination_workbook
        destination_worksheet = destination_workbook.create_sheet(sheet_name)
        
        master_workbook = Workbook()
        master_workbook = load_workbook(fileName) #load_workbook from fileName
        
        master_worksheet = master_workbook.active #grabs active worksheet from master_workbook
        for iteration in master_worksheet.iter_rows( values_only=True): #returns a tuple
            
            # TODO      convert to date format-> 'dd-mm-yyyy'
            # ? temp_tuple = (0,0,0,0,0,0,0,0,0) # temp_tuple has as many variables as there will be columns in the dest_workbook, will contain the formulas and defualt values of 0
            # ? if (first_date <= iteration[0] <= last_date): # ! convert strings to dateformat
            # ?     if(iter[1]== "EC"):
            # TODO      grab values, place them in the temp_tuple on the correct positions
            # TODO      add the temp_tuple to dict
            # TODO      return the temp_tuple to the defualt form (0,0,0,0,0,0)         
            # TODO #verify date in the tuple and add it to a list? maybe
            
            # ? example dict = {
            # ?                   (0,0,0,0,0,date,val,val,0),
            # ?                  (date,val,val,0,0,0,0,0)
            # ?                  }
            # TODO    formulas will be present in the temp_tuple
            # TODO write the dictionary into the hardcoded new excel location?
            pass



if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    ex = App()
    sys.exit(app.exec_())
