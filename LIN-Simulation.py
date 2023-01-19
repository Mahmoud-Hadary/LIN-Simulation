'''
#Read from CSV
#df = pd.read_csv(r'Nodes.csv')

#Open a new window to Show animation
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem, QLabel
import sys
import openpyxl


class AnotherWindow(QWidget):
    """
    This "window" is a QWidget. If it has no parent, it
    will appear as a free-floating window as we want.
    """
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        self.label = QLabel("Another Window")
        layout.addWidget(self.label)
        self.setLayout(layout)


class Main(QWidget):
    def __init__(self):
        super(Main, self).__init__()
        self.setWindowTitle("Load Excel data to QTableWidget")
        
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        self.table_widget = QTableWidget()
        layout.addWidget(self.table_widget)
        self.send_button = QPushButton('Send', self)
        self.send_button.clicked.connect(self.show_new_window)
        self.load_data()
        
    def load_data(self):
        path = "Nodes.xlsx"
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active
        
        self.table_widget.setRowCount(sheet.max_row)
        self.table_widget.setColumnCount(sheet.max_column)
        
        list_values = list(sheet.values)
        self.table_widget.setHorizontalHeaderLabels(list_values[0])
        
        row_index = 0
        for value_tuple in list_values[1:]:
            col_index = 0
            for value in value_tuple:
                self.table_widget.setItem(row_index , col_index, QTableWidgetItem(str(value)))
                col_index += 1
            row_index += 1
    def show_new_window(self, checked):
        w = AnotherWindow()
        w.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Main()
    window.showMaximized()
    app.exec_()
'''
import openpyxl
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem, QLabel,QPushButton
from PyQt5.QtGui import QIcon
class Anim_Window(QWidget):
    def __init__(self):
        super().__init__()
        self.setGeometry(300, 300, 300, 220)
        self.setWindowTitle('LIN Simulation')
        self.setWindowIcon(QIcon('download.png'))
        #self.setFixedSize(610,480)
        self.show()

class LoadTable(QTableWidget):
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
        combox_lay_A = QtWidgets.QComboBox(self)
        combox_lay_A.addItems(["Recieve","Transfer","Dont Care"])
        combox_lay_B = QtWidgets.QComboBox(self)
        combox_lay_B.addItems(["Recieve","Transfer","Dont Care"])
        combox_lay_C = QtWidgets.QComboBox(self)
        combox_lay_C.addItems(["Recieve","Transfer","Dont Care"])
        self.setCellWidget(0, 3, combox_lay_A)
        self.setCellWidget(0, 2, combox_lay_B)
        self.setCellWidget(0, 1, combox_lay_C)
        self.cellChanged.connect(self._cellclicked)


    @QtCore.pyqtSlot(int, int)
    def _cellclicked(self, r, c):
        it = self.item(r, c)
        it.setTextAlignment(QtCore.Qt.AlignCenter)        

    @QtCore.pyqtSlot()
    def _addrow(self):
        rowcount = self.rowCount()
        self.insertRow(rowcount)
        combox_add_A = QtWidgets.QComboBox(self)
        combox_add_A.addItems(["Recieve","Transfer","Dont Care"])
        self.setCellWidget(rowcount, 3, combox_add_A)
        combox_add_B = QtWidgets.QComboBox(self)
        combox_add_B.addItems(["Recieve","Transfer","Dont Care"])
        self.setCellWidget(rowcount, 2, combox_add_B)
        combox_add_C = QtWidgets.QComboBox(self)
        combox_add_C.addItems(["Recieve","Transfer","Dont Care"])
        self.setCellWidget(rowcount, 1, combox_add_C)

    @QtCore.pyqtSlot()
    def _removerow(self):
        if self.rowCount() > 0:
            self.removeRow(self.rowCount()-1)


class ThirdTabLoads(QWidget):
    def switch_window(self):
        self.hide()
        self.second_window = Anim_Window()
        self.second_window.show()
        #export data to excel
        coloumnHeader = []

    def __init__(self, parent=None):
        super(ThirdTabLoads, self).__init__(parent)    

        table = LoadTable()

        add_button = QtWidgets.QPushButton("Add")
        add_button.clicked.connect(table._addrow)

        #Save_button = QtWidgets.QPushButton("Save")
        self.Save_button = QPushButton('&Export',clicked=lambda:'')

        #Save_button.clicked.connect(self.switch_window)  #saves the data and transfers it into an excel sheet then opens a new window

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