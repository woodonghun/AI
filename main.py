import os
import sys

import PreShin.preshin_UI_2d
import PreShin.preshin_UI
import PreShin.volume_template
import PreShin.Mandibular
import PreShin.tooth
import PreShin.dentition_polygon

from PySide2.QtWidgets import QWidget, QPushButton, QApplication, QMessageBox
from PreShin.loggers import logger


def messagebox(i: str):
    signBox = QMessageBox()
    signBox.setWindowTitle("Warning")
    signBox.setText(i)

    signBox.setIcon(QMessageBox.Information)
    signBox.setStandardButtons(QMessageBox.Ok)
    signBox.exec_()


def data_open_3d():
    file_list = os.listdir(os.getcwd())
    if 'group_points_preShin.json' not in file_list:
        messagebox("Group_points_preShin.json 파일의 경로를 확인해 주세요")
        logger.error('group_points_preShin.json file location error')
        return 0
    elif 'landmark.dat' not in file_list:
        messagebox("landmark.dat 파일의 경로를 확인해 주세요")
        logger.error('landmark.dat file location error')
        return 0


# 나중에 2d 용 데이터 나오면 이용
# def data_open_2d():
#     file_list = os.listdir(os.getcwd())
#     if 'group_points_preShin.json' not in file_list:
#         messagebox("Group_points_preShin.json 파일의 경로를 확인해 주세요")
#         logger.error('group_points_preShin.json file location error')
#         return 0
#     elif 'landmark.dat' not in file_list:
#         messagebox("landmark.dat 파일의 경로를 확인해 주세요")
#         logger.error('landmark.dat file location error')
#         return 0

def btn_PreShin_clicked():
    logger.info('btn_PreShin_3D UI start')
    if data_open_3d() == 0:
        return
    PreShin.preshin_UI.PreShin_UI()
    logger.info('btn_PreShin_3D UI end')


def btn_PreShin_2D_clicked():
    logger.info('btn_PreShin_2D UI start')
    PreShin.preshin_UI_2d.PreShin_UI_2d()
    if data_open_3d() == 0:
        return
    logger.info('btn_PreShin_2D UI end')


def btn_vtp_clicked():  # volume template
    logger.info('Volume_template UI start')
    PreShin.volume_template.Vol_Template_UI()
    logger.info('Volume_template UI end')


def btn_mandibular_clicked():
    logger.info('mandibular UI start')
    PreShin.Mandibular.Mandibular_UI()
    logger.info('mandibular UI end')


def btn_tooth_clicked():
    logger.info('tooth UI start')
    PreShin.tooth.Tooth_UI()
    logger.info('tooth UI end')


def btn_manual_clicked():  # 메뉴얼 오픈
    os.startfile(f'{os.getcwd()}/AI_manual.pdf')


class Main(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        logger.info('Main start')

        btn_PreShin = QPushButton(self)
        btn_PreShin.setText("PreShin_3D")
        btn_PreShin.setGeometry(20, 85, 120, 20)
        btn_PreShin.clicked.connect(btn_PreShin_clicked)

        btn_PreShin_2d = QPushButton(self)
        btn_PreShin_2d.setText("PreShin_2D")
        btn_PreShin_2d.setGeometry(20, 110, 120, 20)
        btn_PreShin_2d.clicked.connect(btn_PreShin_2D_clicked)

        btn_vtp = QPushButton(self)
        btn_vtp.setText("Volume_template")
        btn_vtp.setGeometry(20, 60, 120, 20)
        btn_vtp.clicked.connect(btn_vtp_clicked)

        btn_mandibular = QPushButton(self)
        btn_mandibular.setText("Mandibular")
        btn_mandibular.setGeometry(20, 135, 120, 20)
        btn_mandibular.clicked.connect(btn_mandibular_clicked)

        btn_tooth = QPushButton(self)
        btn_tooth.setText("Tooth")
        btn_tooth.setGeometry(20, 160, 120, 20)
        btn_tooth.clicked.connect(btn_tooth_clicked)

        btn_manual = QPushButton(self)
        btn_manual.setText("Manual Open")
        btn_manual.setGeometry(20, 10, 120, 45)
        btn_manual.clicked.connect(btn_manual_clicked)

        self.setWindowTitle('AI')
        self.setGeometry(500, 300, 300, 300)
        self.show()

    def closeEvent(self, QCloseEvent):
        logger.info('Main close')
        self.deleteLater()
        QCloseEvent.accept()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Main()
    sys.exit(app.exec_())
