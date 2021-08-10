from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.error_code import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print("Kiwoom 클래스 실행")

        # event loop 모음
        self.login_event_loop = None

        # 변수 모음
        self.account_num = None

        self.get_oci_instance()
        self.event_slots()

        self.signal_login_com_connect()
        self.get_account_info()

    def get_oci_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)

    def signal_login_com_connect(self):
        self.dynamicCall("CommConnect()")

        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def login_slot(self, err_code):
        print(errors(err_code))
        self.login_event_loop.exit()

    def get_account_info(self):
        account_list = self.dynamicCall("GetLogininfo(String)", "ACCNO")
        account_num = account_list.split(';')[0]
        print(f"계좌번호 = {account_num}")
