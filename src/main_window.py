import sys
import os
from . import excel
from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import QFileDialog


qtCreatorFile = os.path.join(os.getcwd(),'src/ui/main_window.ui')
Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)

class App(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.select_excel_button.clicked.connect(self.search_file)
        self.send_emails_button.clicked.connect(self.send_emails)
    
    def show_window(self):
        return self.show()

    def search_file(self):
        self.fname = QFileDialog.getOpenFileName(self,'Open File',os.getenv('HOME'), 'Excel 2007 file (*.xls)')
        self.filename.setText(self.fname[0])

    def send_emails(self):
        if self.fname[0] != '':
            excel.applysys.process_excel(self.fname[0])
        self.status_label.setText('Procesado!!!')
            
def main():
    app = QtWidgets.QApplication(sys.argv)
    window = App()
    window.show_window()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
