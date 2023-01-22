import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem, QLabel,QPushButton, QFileDialog
from PyQt5.QtGui import QIcon
import pandas as pd
import xlwt
import os
coloumn_count = 0

class Anim_Window(QWidget):
    def __init__(self):
        super().__init__()
        self.setGeometry(300, 300, 300, 220)
        self.setWindowTitle('LIN Simulation')
        self.setWindowIcon(QIcon('download.png'))
        #self.setFixedSize(610,480)
        self.show()

class LoadTable(QTableWidget,QWidget):
    def switch_window(self):
        self.savefile()
        self.hide()
        self.second_window = Anim_Window()
        self.second_window.show()

    def savefile(self):
        filename = str(QFileDialog.getSaveFileName(self, 'Save File', '', ".csv(*.csv)"))
        filename = filename.split(",")
        filename=filename[0]
        filename=filename[2:-1]
        print(filename)
        wbk = xlwt.Workbook()
        self.sheet = wbk.add_sheet("sheet", cell_overwrite_ok=True)
        self.add2()
        if(os.path.exists(filename)):
            os.remove(filename)
        wbk.save(filename) 

    def add2(self):
        row = 0
        col = 0         
        for i in range(self.columnCount()):
            for x in range(self.rowCount()):   
                try:     
                    teext = str(self.item(row, col).text())
                    if col > 0:
                        if teext != "Recieved" and teext != "Dont Care" and teext != "Transfer":
                            QtWidgets.QMessageBox.critical(self, "Invalid Input", "The Input for the Nodes should be either Recieved, Dont Care, or Transfer")
                    self.sheet.write(row, col, teext)
                    row += 1
                
                except AttributeError:
                    row += 1
            row = 0
            col +=1
        #content = self.combo_box.currentText()
        '''
    def Export_to_Excel(self):
        columnHeaders = []
        # create column header list
        for j in range(self.model().columnCount()):
            columnHeaders.append(self.horizontalHeaderItem(j).text())

        df = pd.DataFrame(columns=columnHeaders)

        # create dataframe object recordset
        for row in range(self.rowCount()):
            for col in range(self.columnCount()):
                df.at[row, columnHeaders[col]] = self.item(row, col).text()

        df.to_excel('Dummy File XYZ.xlsx', index=False)
        print('Excel file exported')
        '''
    def __init__(self, parent=None):
        
        super(LoadTable, self).__init__(1, 4, parent)
        self.setFont(QtGui.QFont("Helvetica", 10, QtGui.QFont.Normal, italic=False))   
        headertitle = ("Frame ID","Master","Slave A","Slave B")
        self.setHorizontalHeaderLabels(headertitle)
        self.verticalHeader().hide()
        self.horizontalHeader().setHighlightSections(False)
        self.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Fixed)

        self.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        self.setColumnWidth(0, 130)

        self.cellChanged.connect(self._cellclicked)
        


    @QtCore.pyqtSlot(int, int)
    def _cellclicked(self, r, c):
        it = self.item(r, c)
        it.setTextAlignment(QtCore.Qt.AlignCenter)        

    @QtCore.pyqtSlot()
    def _addrow(self):
        rowcount = self.rowCount()
        self.insertRow(rowcount)
 

    @QtCore.pyqtSlot()
    def _removerow(self):
        if self.rowCount() > 0:
            self.removeRow(self.rowCount()-1)


class ThirdTabLoads(QWidget):

    def __init__(self, parent=None):
        super(ThirdTabLoads, self).__init__(parent)    

        table = LoadTable()
        #layout = QVBoxLayout()
        #self.setLayout(layout)

        add_button = QtWidgets.QPushButton("Add")
        add_button.clicked.connect(table._addrow)
        
        Save_button = QtWidgets.QPushButton("Save")
        #self.Save_button = QPushButton('&Export',clicked=self.switch_window)
        #layout.addWidget(self.Save_button)
        Save_button.clicked.connect(table.switch_window)  #saves the data and transfers it into an excel sheet then opens a new window

        delete_button = QtWidgets.QPushButton("Delete")
        delete_button.clicked.connect(table._removerow)

        button_layout = QtWidgets.QVBoxLayout()
        button_layout.addWidget(add_button, alignment=QtCore.Qt.AlignBottom)
        button_layout.addWidget(delete_button, alignment=QtCore.Qt.AlignTop)
        button_layout.addWidget(Save_button, alignment=QtCore.Qt.AlignBottom)


        tablehbox = QtWidgets.QHBoxLayout()
        tablehbox.setContentsMargins(10, 10, 10, 10)
        tablehbox.addWidget(table)

        grid = QtWidgets.QGridLayout(self)
        grid.addLayout(button_layout, 0, 1)
        grid.addLayout(tablehbox, 0, 0)        


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    w = ThirdTabLoads()
    w.show()
    sys.exit(app.exec_())