import sys
from unicodedata import name
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableView, QTableWidget, QWidget, QPushButton, QFileDialog, QTableWidget
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtGui import QIcon
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

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
        self.statusBar().showMessage('Test App written in Python and QT')
        self.setFixedSize(300,150)
        
        #Browse button
        button = QPushButton('browse',self)
        button.setToolTip('select excel file')
        button.move(20,20)
        button.clicked.connect(self.openFileNameDialog)
        
        self.show()

    # @pyqtSlot()
    # def on_click(self):
    #     print('browse works')
    
    
    #Open File Dialog Window
    def openFileNameDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"Select Excel", "","Excel Files(*.xlsx)", options=options)
        workbook = Workbook()
        workbook = load_workbook(fileName)
        if fileName:
            print(workbook)



if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    ex = App()
    sys.exit(app.exec_())