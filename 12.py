import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem, QWidget, QLabel, QMessageBox, QHeaderView
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

class CuttingOptimizer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Оптимизация Раскроя')
        self.setGeometry(100, 100, 1000, 800)

        main_layout = QVBoxLayout()

        self.label = QLabel('Загрузите файл Excel с данными о панелях:', self)
        self.label.setFont(QFont('Arial', 14))
        main_layout.addWidget(self.label, alignment=Qt.AlignCenter)

        self.tableWidget = QTableWidget(self)
        self.tableWidget.setFont(QFont('Arial', 12))
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
        main_layout.addWidget(self.tableWidget)

        button_layout = QHBoxLayout()

        self.uploadButton = QPushButton('Загрузить файл Excel', self)
        self.uploadButton.setFont(QFont('Arial', 12))
        self.uploadButton.clicked.connect(self.uploadFile)
        button_layout.addWidget(self.uploadButton)

        self.runButton = QPushButton('Выполнить оптимизацию', self)
        self.runButton.setFont(QFont('Arial', 12))
        self.runButton.clicked.connect(self.runOptimization)
        button_layout.addWidget(self.runButton)

        main_layout.addLayout(button_layout)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def uploadFile(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        fileName, _ = QFileDialog.getOpenFileName(self, "Открыть файл Excel", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if fileName:
            self.data = pd.read_excel(fileName)
            self.displayData()
            self.showMessage('Файл успешно загружен')

    def displayData(self):
        if not self.data.empty:
            self.tableWidget.setColumnCount(len(self.data.columns))
            self.tableWidget.setRowCount(len(self.data.index))
            self.tableWidget.setHorizontalHeaderLabels(self.data.columns)

            for i in range(len(self.data.index)):
                for j in range(len(self.data.columns)):
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(self.data.iat[i, j])))

    def showMessage(self, message):
        QMessageBox.information(self, 'Сообщение', message)

    def runOptimization(self):
        if hasattr(self, 'data'):
            panel_size = (2500, 1250)
            saw_thickness = 6
            results = self.optimizeCutting(self.data, panel_size, saw_thickness)
            self.generatePDF(results)
            self.showMessage('Оптимизация и генерация PDF завершены')
        else:
            self.showMessage('Пожалуйста, сначала загрузите файл Excel.')

    def optimizeCutting(self, data, panel_size, saw_thickness):
        results = []
        for index, row in data.iterrows():
            length = row['Длина']
            width = row['Ширина']
            result = {
                'Панель': f"Панель {index + 1}",
                'Длина': length,
                'Ширина': width,
                'Остаток длины': panel_size[0] - length - saw_thickness,
                'Остаток ширины': panel_size[1] - width - saw_thickness
            }
            results.append(result)
        return results


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = CuttingOptimizer()
    ex.show()
    sys.exit(app.exec_())
