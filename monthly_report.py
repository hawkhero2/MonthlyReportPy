from distutils.command.build_scripts import first_line_re
import sys
import json
from unicodedata import name
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QTableView, QTableWidget, QWidget, QPushButton, QFileDialog, QTableWidget
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtGui import QIcon
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pkg_resources import working_set

class App(QMainWindow):
    
    
    def __init__(self):
        super().__init__()
        self.title = 'Python Test App'
        self.left = 100
        self.top = 100
        self.width = 300
        self.height = 150
        
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
        
        #Qlabel account name
        user_account = QLabel(self)
        user_account.move(140,4)
        user_account.setText('Account')
        
        # Create textbox name
        self.textbox = QLineEdit(self)
        self.textbox.move(20, 30)
        self.textbox.resize(100,20)
        
        # Create textbox user_account
        self.textbox = QLineEdit(self)
        self.textbox.move(140, 30)
        self.textbox.resize(100,20)
        
        # Create textbox first date
        self.textbox = QLineEdit(self)
        self.textbox.move(20, 100)
        self.textbox.resize(100,20)
        
        
        # Create textbox last date
        self.textbox = QLineEdit(self)
        self.textbox.move(140, 100)
        self.textbox.resize(100,20)
        
        #QLabel first date
        first_date = QLabel(self)
        first_date.move(20, 75)
        first_date.setText('First Date')
        
        #QLabel last date
        first_date = QLabel(self)
        first_date.move(140, 75)
        first_date.setText('Last Date')
        
        self.show()
    
    #Open File Dialog Window
    def openFileNameDialog(self):
        
        fileName, _ = QFileDialog.getOpenFileName(self,"Select Excel", "","Excel Files(*.xlsx)") #Grab File
        
        workbook = Workbook()
        destination_workbook = Workbook() #inst destination_workbook
        # destination_workbook.save('c:/users/currentuser/desktop/'+operator_name+'.xlsx') #create destination_workbook
        workbook = load_workbook(fileName) #load_workbook from fileName
        worksheet = workbook.active
        for iteration in worksheet.iter_rows(min_row=1, max_col=2, values_only=True): #returns a tuple
            if (iteration[0]==4):
                print(iteration[0])
            #verify date in the tuple and add it to a list? maybe
            #write the list into the hardcoded new excel location?
        # newWorkbook = Workbook()
        # dest_filename = path to the new excel to be created in-> "c:\users\desktop\name.xlsx"
        # newWorkbook.save(fileName = dest_filename)
    



if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    ex = App()
    sys.exit(app.exec_())