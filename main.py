import sys

from PreShin.loggers import make_logger
import PreShin.preshin_UI_2d
import PreShin.preshin_UI
from PySide2.QtWidgets import QWidget, QPushButton, QApplication

logger = make_logger()

def btn_PreShin_clicked():
    PreShin.preshin_UI.PreShin_UI()


def btn_PreShin_2D_clicked():
    PreShin.preshin_UI_2d.PreShin_UI_2d()


class Main(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        logger.info('start')
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


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Main()
    sys.exit(app.exec_())
