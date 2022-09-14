import sys

import PreShin.preshin_UI_2d
import PreShin.preshin_UI
from PySide2.QtWidgets import QWidget, QPushButton, QApplication


class Main(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        btn_preshin = QPushButton(self)
        btn_preshin.setText("PreShin_3D")
        btn_preshin.setGeometry(20, 35, 100, 20)
        btn_preshin.clicked.connect(self.btn_preshin_clicked)

        btn_preshin_2d = QPushButton(self)
        btn_preshin_2d.setText("PreShin_2D")
        btn_preshin_2d.setGeometry(20, 60, 100, 20)
        btn_preshin_2d.clicked.connect(self.btn_preshin_2D_clicked)

        self.setWindowTitle('AI')
        self.setGeometry(500, 300, 150, 150)
        self.show()

    def btn_preshin_clicked(self):
        PreShin.preshin_UI.PreShin_UI()

    def btn_preshin_2D_clicked(self):
        PreShin.preshin_UI_2d.PreShin_UI()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Main()
    sys.exit(app.exec_())
