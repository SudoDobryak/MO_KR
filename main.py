import symplex
import gui
import sys
from PyQt5 import uic, QtWidgets
import des
import pandas as pd
import numpy as np
import xlrd


book = 'кр.xlsx'
sheet = 'MO'


class DlgMain(QtWidgets.QMainWindow, gui.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(lambda xl_path=book, sheet_name=sheet: self.loadEcxelData(xl_path, sheet_name))

    def loadEcxelData(self, excel_file_dir, worksheet_name):
        df = pd.read_excel(excel_file_dir, worksheet_name)
        if df.size == 0:
            return
        try:
            df.fillna('', inplace=True)
            self.listWidget.setRowCount(df.shape[0])
            self.listWidget.setColumnCount(df.shape[1])
            self.listWidget.setHorizontalHeaderLabels(df.columns)
            self.listWidget.verticalHeader().setVisible(False)
            # returns pandas array object
            for row in df.iterrows():
                values = row[1]
                for col_index, value in enumerate(values):
                    if isinstance(value, (float, int)):
                        value = '{0:0,.0f}'.format(value)
                        tableItem = QtWidgets.QTableWidgetItem(str(value))
                        self.listWidget.setItem(row[0], col_index, tableItem)
        except Exception as exc:
            print(exc)


def main():
    # Recompile ui
    with open("gui.ui") as ui_file:
        with open("gui.py", "w") as py_ui_file:
            uic.compileUi(ui_file, py_ui_file)
    app = QtWidgets.QApplication(sys.argv)  # create application
    dlgMain = DlgMain()  # create main GUI window
    dlgMain.show()  # show GUI
    sys.exit(app.exec_())  # execute application

if __name__ == '__main__':
    main()