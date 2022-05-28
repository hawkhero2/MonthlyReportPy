from cgi import test
from cgitb import text
from distutils.command.build_scripts import first_line_re
import os
from datetime import datetime
import sys
from time import time_ns
from unicodedata import name
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
        self.title = 'Monthly Excel Report'
        self.left : int = 100
        self.top = 100
        self.width = 300
        self.height = 150
        self.setStyleSheet("""
                           background-color:#363636
                           """)
        
        self.initUI()
        
    def createTextblock(self,obj, int_pos_x, int_pos_y, int_size_x, int_size_y, style):
        obj.move(int_pos_x,int_pos_y)
        obj.resize(int_size_x,int_size_y)
        obj.setStyleSheet(style)
        
    def createQLabel(self,obj,name,int_x,int_y, style):
        obj.move(int_x,int_y)
        obj.setStyleSheet(style)
        obj.setText(name)
    

    def initUI(self):
        
        button_style = """
                color: #cfcfcf;
                background-color: #474747;
        """
        
        textbox_style ="""
                color: #cfcfcf;
                background-color: #474747;
        """
        
        qlabel_style = """
                color: #cfcfcf;
        """
        
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.statusBar().showMessage('Select master excel')
        self.setFixedSize(450,250)
        
        # Browse button
        button = QPushButton('browse',self)
        button.setToolTip('select excel file')
        button.move(20,180)
        button.clicked.connect(self.openFileNameDialog)
        button.setStyleSheet(button_style)
        
        # * ----------Textbox----------------------
        
        # Create textbox 
        self.full_name = QLineEdit(self)
        self.user_account = QLineEdit(self)
        self.first_date = QLineEdit(self)
        self.last_date = QLineEdit(self)
        self.createTextblock(self.full_name,20,35,100,20,textbox_style)
        self.createTextblock(self.user_account,140,35,100,20,textbox_style)
        self.createTextblock(self.first_date,20,105,100,20,textbox_style)
        self.createTextblock(self.last_date,140,105,100,20,textbox_style)
               
       
        # * --------------QLabels----------------
        
        # Qlabel name 
        person_name = QLabel(self)
        user_account = QLabel(self)
        first_date = QLabel(self)
        last_date = QLabel(self)
        info_zone= QLabel(self)
        info_zone.setFixedSize(200,140)
        
        info_zone_txt= """
        How to :
        
        1. Fill in the fields with the
        required data 
        Date format : 24/05/2022
        
        2. Then select the master excel 
        and relax
        """""
        self.createQLabel(info_zone,info_zone_txt,250,4,qlabel_style)
        self.createQLabel(person_name,'Full Name',20,4,qlabel_style)
        self.createQLabel(user_account,'Account',140,4,qlabel_style)
        self.createQLabel(first_date,'First Date',20,75,qlabel_style)
        self.createQLabel(last_date,'Last Date',140,75,qlabel_style)
                
        # *   Test Button
        test_button = QPushButton('Test Button', self)
        test_button.move(140,180)
        test_button.clicked.connect(self.testButton)
        test_button.setStyleSheet(button_style)
        
        self.show()
    
    
    # * Test Button
    def testButton(self):
        test_wb = Workbook()
        test_wb = load_workbook(filename="test.xlsx")
        test_ws = test_wb.active
        temp_l = [0]*23
        row_nr = "1"
        row_nr2 = "2"
        temp_l.insert(21,"=SUM(A"+row_nr+":A"+row_nr2+")")
        test_ws.append(temp_l)
        test_ws.append({'G':'=SUM(A1:A2)'})
        print(temp_l)
        test_wb.save('test.xlsx')
        
        # temporary_list = ["date","activity","docs","time","speed",2,80,20,["inside","the","list"]] # cannot append list nesting lists
        # test_wb.create_sheet(title='test')
        # test_ws=test_wb["Sheet1"]
        # i=0
        # while (i<len(temporary_list)):
        #     test_ws.append(temporary_list)
        #     i+=1
        
    # * Open File Dialog Window Event
    def openFileNameDialog(self):     
        
        full_name = self.full_name.text()
        first_date_obj = datetime.strptime(self.first_date.text(),'%d/%m/%Y')
        last_date_obj = datetime.strptime(self.last_date.text(),'%d/%m/%Y')
        user_account = self.user_account.text()
        sheet_name = self.first_date.text()+"-"+self.last_date.text()
        
        excel_header = [
            "DATE",
            "EC DOCS",
            "IC DOCS",
            "GC DOCS",
            "EC TIME",
            "IC TIME",
            "GC TIME",
            "EXTRA TIME",
            "EC SPEED",
            "IC SPEED",
            "GC SPEED",
            "TOTAL TIME",
            "WORKED",
            "MINUTES",
            "MONTHLY EC SPEED",
            "MONTHLY IC SPEED",
            "MONTHLY GC SPEED",
        ]
        
        desktop = os.path.expanduser("~\Desktop\\") #path for current user desktop
        
        fileName, _ = QFileDialog.getOpenFileName(self,"Select Excel", "","Excel Files(*.xlsx)") #Grab File
        
        destination_workbook = Workbook() #inst destination_workbook
        destination_workbook.save(desktop+full_name+'.xlsx') #create destination_workbook
        destination_workbook.create_sheet(sheet_name)
        destination_worksheet = destination_workbook[sheet_name]
        
        destination_worksheet.append(excel_header)
        
        master_workbook = Workbook()
        master_workbook = load_workbook(fileName) #load_workbook from fileName
        master_worksheet = master_workbook.active #grabs active worksheet from master_workbook
        i=2
        for iteration in master_worksheet.iter_rows( values_only=True): #returns a tuple
            temp_list = [0]*17
            if ((first_date_obj <= iteration[0] <= last_date_obj)& (iteration[3] == user_account)):
                temp_list[0] = iteration[0] # write date
                docs = iteration[4]
                time = iteration[5]
                extra_time = iteration[7]
                if(iteration[1]== "Expertise"):
                    temp_list[1] = docs
                    temp_list[4] = time
                    temp_list[7] = extra_time
                    temp_list[8] = "=IFERROR(B"+i+"/E"+i+",0)" #EC SPEED
                    temp_list[11] = "=SUM(E"+i+"F"+i+"G"+i+"H"+i+")" #TOTAL TIME
                    temp_list[12] = 8 #WORKED
                    temp_list[13] = "=(L"+i+"-M"+i+")" #MINUTES
                    destination_worksheet.append(temp_list)
                    i+=1
                if(iteration[1]== "IC"):
                    temp_list[2] = docs
                    temp_list[5] = time
                    temp_list[7] = extra_time
                    temp_list[9] = "=IFERROR(C"+i+"/F"+i+",0)"
# TODO  grab values, place them in the temp_list on the correct positions
# TODO  temp_list = total_columns in destination_worksheet
# TODO  return the temp_list to the defualt form (0,0,0,0,0,0)
# TODO  verify date in the tuple and add it to a list? maybe
# TODO  formulas will be present in the temp_list
# TODO  destination_workbook.save(desktop+full_name+'.xlsx')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    ex = App()
    sys.exit(app.exec_())
