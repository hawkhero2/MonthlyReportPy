import sys
import json
from unicodedata import name
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableView, QTableWidget, QWidget, QPushButton, QFileDialog, QTableWidget
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
        self.setFixedSize(300,150)
        
        #Browse button
        button = QPushButton('browse',self)
        button.setToolTip('select excel file')
        button.move(20,20)
        button.clicked.connect(self.openFileNameDialog)
        
        self.show()
    
    #Open File Dialog Window
    def openFileNameDialog(self):
        #Open File
        # options = QFileDialog.Options()
        # options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"Select Excel", "","Excel Files(*.xlsx)")
        
        #load_workbook from fileName
        workbook = Workbook()
        workbook = load_workbook(fileName)
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