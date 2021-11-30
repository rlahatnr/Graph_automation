from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QTableWidgetItem, QHeaderView
from PyQt5.QtCore import QUrl
from main_ui import Ui_MainWindow
import sys
import plotly.express as px
from selenium.webdriver.common.keys import Keys
import time
import random
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
import chromedriver_autoinstaller
import pandas as pd

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
        self.comboBox_3.addItems(['Bar', 'Line', 'Pie'])
        self.spinBox.setValue(20)
        self.comboBox_3.currentTextChanged.connect(self.status)

    def status(self):
        if self.comboBox_3.currentText() != 'Pie':
            self.label_2.setText('X')
            self.label_3.setText('Y')
            self.checkBox.setEnabled(True)
        elif self.comboBox_3.currentText() == 'Pie':
            self.label_2.setText('names')
            self.label_3.setText('values')
            self.checkBox.setDisabled(True)

    def make_graph(self):
        try:
            txt = self.lineEdit_2.text().strip(',').split(',')
            if self.comboBox_3.currentText() == 'Bar':
                fig = px.bar(df, x=self.comboBox.currentText(), y=txt, title=self.lineEdit.text(),
                              color=self.comboBox.currentText(), text='value')
                fig.update_layout(showlegend=False)

            elif self.comboBox_3.currentText() == 'Line':
                fig = px.line(df, x=self.comboBox.currentText(), y=txt, title=self.lineEdit.text(),
                              color=self.comboBox.currentText(), text='value')
                fig.update_traces(mode="markers+text+lines", textposition='top center', hovertemplate=None)
                fig.update_layout(hovermode="x")

            elif self.comboBox_3.currentText() == 'Pie':
                fig = px.pie(df, names=self.comboBox.currentText(), values=txt[0], hole=.4)
                fig.update_traces(textposition='inside', textinfo='percent+label')

            if self.checkBox.isChecked():
                fig['layout']['yaxis']['autorange'] = "reversed"

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
            self.comboBox.clear()
            self.comboBox.addItem('없음')
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
        except Exception as e:
            print(e)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()

    sys.exit(app.exec_())