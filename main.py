import os
import sys

import PreShin.preshin_UI_2d
import PreShin.preshin_UI
from PySide2.QtWidgets import QWidget, QPushButton, QApplication, QMessageBox
from PreShin.loggers import logger


def messagebox(i: str):
    signBox = QMessageBox()
    signBox.setWindowTitle("Warning")
    signBox.setText(i)

    signBox.setIcon(QMessageBox.Information)
    signBox.setStandardButtons(QMessageBox.Ok)
    signBox.exec_()


def data_open():
    file_list = os.listdir(os.getcwd())
    if 'group_points_preShin.json' not in file_list:
        messagebox("Group_points_preShin.json 파일의 경로를 확인해 주세요")
        logger.error('group_points_preShin.json file location error')
        return 0
    elif 'landmark.dat' not in file_list:
        messagebox("landmark.dat 파일의 경로를 확인해 주세요")
        logger.error('landmark.dat file location error')
        return 0


def btn_PreShin_clicked():
    logger.info('btn_PreShin_3D UI start')
    if data_open() == 0:
        return
    PreShin.preshin_UI.PreShin_UI()
    logger.info('btn_PreShin_3D UI end')


def btn_PreShin_2D_clicked():
    logger.info('btn_PreShin_2D UI start')
    PreShin.preshin_UI_2d.PreShin_UI_2d()
    logger.info('btn_PreShin_2D UI end')


class Main(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        logger.info('Main start')
        btn_PreShin = QPushButton(self)
        btn_PreShin.setText("PreShin_3D")
        btn_PreShin.setGeometry(20, 35, 100, 20)
        btn_PreShin.clicked.connect(btn_PreShin_clicked)

        btn_PreShin_2d = QPushButton(self)
        btn_PreShin_2d.setText("PreShin_2D")
        btn_PreShin_2d.setGeometry(20, 60, 100, 20)
        btn_PreShin_2d.clicked.connect(btn_PreShin_2D_clicked)

        self.setWindowTitle('AI')
        self.setGeometry(500, 300, 150, 150)
        self.show()

    def closeEvent(self, QCloseEvent):
        logger.info('Main close')
        self.deleteLater()
        QCloseEvent.accept()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Main()
    sys.exit(app.exec_())
