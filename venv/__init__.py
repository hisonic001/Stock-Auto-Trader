from PyQt5.QtWidgets import *
import sys
from kiwoom.kiwoom import *


class Main:
    def __init__(self):
        print("Main class 실행")

        self.app = QApplication(sys.argv)
        self.kiwoom = Kiwoom()
        self.app.exec_()


if __name__ == "__main__":
    Main()
