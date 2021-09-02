from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QTableWidgetItem, QHeaderView
from PyQt5.QtCore import QUrl
from main_ui import Ui_MainWindow
import sys
import pandas as pd
import plotly.express as px
import time

class MainWindow(QMainWindow,Ui_MainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.setupUi(self)
        self.setWindowTitle('선그래프 자동만들기')

        self.pushButton_3.clicked.connect(self.get_file)

        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget.selectionModel().selectionChanged.connect(self.select_y)
        self.pushButton_4.clicked.connect(self.make_graph)
        self.spinBox.setValue(20)

    def make_graph(self):
        try:
            txt = self.lineEdit_2.text().strip(',').split(',')
            if self.comboBox_2.currentText() != '없음':
                fig = px.line(df, x=self.comboBox.currentText(), y=txt, title=self.lineEdit.text(), color=self.comboBox_2.currentText(), text='value')
            else:
                fig = px.line(df, x=self.comboBox.currentText(), y=txt, title=self.lineEdit.text(), text='value')

            fig.update_traces(mode="markers+text+lines", textposition='top center', hovertemplate=None)
            fig.update_layout(hovermode="x")
            fig.update_layout(
                font=dict(
                    family="Courier New, monospace",
                    size=int(self.spinBox.text())
                )
            )
            fig.write_html(r'{}.html'.format(self.lineEdit.text()))
            time.sleep(1)
            url = QUrl.fromLocalFile(r'{}.html'.format(self.lineEdit.text()))
            self.webEngineView.load(url)

        except Exception as e:
            print(e)

    def select_y(self):
        try:
            col = self.tableWidget.currentColumn()
            y = list(df)[col]
            if f'{y},' not in self.lineEdit_2.text():
                self.lineEdit_2.insert(y + ',')
            elif f'{y},' in self.lineEdit_2.text():
                self.lineEdit_2.setText(self.lineEdit_2.text().replace(f'{y},', ''))
        except Exception as e:
            print(e)

    def get_file(self):
        global df

        try:
            combos = [self.comboBox, self.comboBox_2]
            for i in combos:
                i.clear()
                i.addItem('없음')
            self.lineEdit_2.setText('')
            lists = QFileDialog.getOpenFileName(self, 'Open File')
            df = pd.read_excel(lists[0])
            self.tableWidget.setRowCount(df.shape[0])
            self.tableWidget.setColumnCount(df.shape[1])
            self.tableWidget.setHorizontalHeaderLabels(df.columns)

            for i in range(df.shape[0]):
                for j in range(df.shape[1]):
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))

            self.lineEdit.setText(lists[0])
            self.comboBox.addItems(list(df))
            self.comboBox_2.addItems(list(df))
        except Exception as e:
            print(e)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()

    sys.exit(app.exec_())