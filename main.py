import sys

import PreShin.preshin_UI
from PySide2.QtWidgets import QWidget, QPushButton, QApplication


class Main(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        btn_preshin = QPushButton(self)
        btn_preshin.setText("preshin")
        btn_preshin.clicked.connect(self.btn_preshin_clicked)

        self.setWindowTitle('AI')
        self.setGeometry(500, 300, 550, 450)
        self.show()

    def btn_preshin_clicked(self):
        PreShin.preshin_UI.Preshin_UI()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Main()
    sys.exit(app.exec_())
