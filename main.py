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
        self.pushButton.clicked.connect(lambda _, xl_path=book, sheet_name=sheet: self.loadEcxelData(xl_path, sheet_name))
        self.pushButton_2.clicked.connect(self.alg)

    def loadEcxelData(self, excel_file_dir, worksheet_name):
        df = pd.read_excel(excel_file_dir, worksheet_name)
        if df.size == 0:
            return
        try:
            df.fillna('', inplace=True)
            self.tableWidget.setRowCount(df.shape[0])
            self.tableWidget.setColumnCount(df.shape[1])
            self.tableWidget.setHorizontalHeaderLabels(df.columns)
            self.tableWidget.verticalHeader().setVisible(False)
            # returns pandas array object
            for row in df.iterrows():
                values = row[1]
                for col_index, value in enumerate(values):
                    if isinstance(value, (float, int)):
                        value = '{0:0,.0f}'.format(value)
                        tableItem = QtWidgets.QTableWidgetItem(str(value))
                        self.tableWidget.setItem(row[0], col_index, tableItem)
        except Exception as exc:
            print(exc)

    def alg(self):
        m = symplex.gen_matrix(7, 2)
        symplex.constrain(m, f"0,2,5,3,0,2,4,G,470")
        symplex.constrain(m, f"3,2,0,1,2,1,0,G,450")
        symplex.obj(m, '10,25,0,35,30,20,10')
        a = symplex.minz(m)
        m2 = symplex.gen_matrix(7, 2)
        symplex.constrain(m2, f"0,2,5,3,0,2,4,G,470")
        symplex.constrain(m2, f"3,2,0,1,2,1,0,G,{self.textEdit.toPlainText()}")
        symplex.obj(m2, '10,25,0,35,30,20,10')
        a2 = symplex.minz(m2)
        self.textEdit_2.setText(f"X1 = {a['x1']}" + f"\t\tY1 = {a2['x1']}\n"
                                f"X2 = {a['x2']}" + f"\t\tY2 = {a2['x2']}\n"
                                f"X3 = {a['x3']}" + f"\t\tY3 = {a2['x3']}\n"
                                f"X4 = {a['x4']}" + f"\t\tY4 = {a2['x4']}\n"
                                f"X5 = {a['x5']}" + f"\t\tY5 = {a2['x5']}\n"
                                f"X6 = {a['x6']}" + f"\t\tY6 = {a2['x6']}\n"
                                f"X7 = {a['x7']}" + f"\t\tY7 = {a2['x7']}\n"
                                f"Минимальная функция для 450 равна: {a['min']}\n"
                                f"Минимальная функция для {self.textEdit.toPlainText()} равна: {a2['min']}")


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