from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.error_code import *
from PyQt5.QtTest import *
import json


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print("Kiwoom 클래스 실행")

        # event loop 모음
        self.login_event_loop = None
        self.detail_account_info_event_loop = QEventLoop()
        self.calculator_event_loop = QEventLoop()


        #

        # 변수 모음
        self.account_num = None
        self.account_stock_dict = {}
        self.unsigned_order_dict = {}
        self.calcul_data = []
        #

        # 계좌 관련 변수 모음
        self.use_money = 0
        self.use_money_percent = 0.5
        #

        # 스크린 번호 모음
        self.scr_my_info = '2000'
        self.scr_calculation_stock = "4000"
        #

        self.get_oci_instance()
        self.event_slots()

        self.signal_login_com_connect()  # 자동 아이디 로그인
        self.get_account_info()  # 계좌 번호 알아보기
        self.detail_account_info()  # 예수금 정보 요청
        self.detail_account_mystock()  # 계좌평가 잔고 내역 요청
        self.unsigned_order()  # 미체결 내역 요청

        self.calculator_fnc()  # 종목분석용//

    # OCI 가져오기
    def get_oci_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    # 이벤트 정보를 각 슬롯으로 보내주기
    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)

    # 로그인 이벤트
    def signal_login_com_connect(self):
        self.dynamicCall("CommConnect()")

        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def login_slot(self, err_code):
        print(errors(err_code))
        self.login_event_loop.exit()

    # 계좌번호 조회
    def get_account_info(self):
        account_list = self.dynamicCall("GetLogininfo(str)", "ACCNO")
        self.account_num = account_list.split(';')[1]
        print(f"계좌번호 = {self.account_num}")

    # 계좌정보 조회
    def detail_account_info(self):
        self.dynamicCall("SetInputValue(str, str)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(str, str)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(str, str)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(str, str)", "조회구분", "2")
        self.dynamicCall("CommRqData(str, str, int, str)", "예수금상세현황요청", "opw00001", "0", self.scr_my_info)

        self.detail_account_info_event_loop
        self.detail_account_info_event_loop.exec_()

    # 시장종목 코드리스트 조회
    def get_code_list_by_market(self, market_code):
        code_list = self.dynamicCall("GetCodeListByMarket(str)", market_code).split(';')[:-1]
        return code_list

    # 계좌평가잔고내역
    def detail_account_mystock(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(str, str)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(str, str)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(str, str)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(str, str)", "조회구분", "2")
        self.dynamicCall("CommRqData(str, str, int, str)", "계좌평가잔고내역요청", "opw00018", sPrevNext, self.scr_my_info)

        self.detail_account_info_event_loop
        self.detail_account_info_event_loop.exec_()

    # 미체결 정보 조회
    def unsigned_order(self, sPrevNext="0"):
        print("미체결요청")
        self.dynamicCall("SetInputValue(str, str)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(str, str)", "체결구분", "1")
        self.dynamicCall("SetInputValue(str, str)", "매매구분", "0")
        self.dynamicCall("CommRqData(str, str, int, str)", "실시간미체결요청", "opt10075", sPrevNext, self.scr_my_info)

        self.detail_account_info_event_loop.exec_()

    # 종목분석 실행용 함수
    def calculator_fnc(self):
        code_list = self.get_code_list_by_market("10")

        for idx, code in enumerate(code_list):
            self.dynamicCall("DisconnectRealData(string)", self.scr_calculation_stock)

            print(f"{idx+1}/{len(code_list)} : KOSDAQ Stock Code : {code} is updating...")

            self.day_graph(code=code)

    # 일봉 데이터 조회 - 600일치의 데이터를 받을 수 있다.
    def day_graph(self, code=None, date=None, sPrevNext="0"):
        QTest.qWait(3600)

        self.dynamicCall("SetInputValue(str, str)", "종목코드", code)
        if date is not None:
            self.dynamicCall("SetInputValue(str, str)", "기준일자", date)
        self.dynamicCall("SetInputValue(str, str)", "수정주가구분", "1")

        self.dynamicCall("CommRqData(str, str, int, str)", "주식일봉차트조회", "opt10081", sPrevNext, self.scr_calculation_stock)

        self.calculator_event_loop.exec_()

    # TR 데이터 슬롯
    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        '''
        :param sScrNo: 스크린 번호
        :param sRQName: 사용자 구분명
        :param sTrCode: TR 코드
        :param sRecordName: 사용 안함
        :param sPrevNext: 다음 페이지 유무 확인
        :return:
        '''

        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall('GetCommData(str, str, int, str)', sTrCode, sRQName, 0, "예수금")
            print(f'예수금 {int(deposit)}원')
            self.use_money = int(deposit) * self.use_money_percent
            self.use_money = self.use_money / 4

            avail_deposit = self.dynamicCall('GetCommData(str, str, int, str)', sTrCode, sRQName, 0, "출금가능금액")
            print(f"출금가능금액 {int(avail_deposit)}원")

            self.detail_account_info_event_loop.exit()

        elif sRQName == "계좌평가잔고내역요청":
            total_balance = self.dynamicCall('GetCommData(str, str, int, str)', sTrCode, sRQName, 0, "총매입금액")
            total_profit_rate = self.dynamicCall('GetCommData(str, str, int, str)', sTrCode, sRQName, 0, "총수익률(%)")

            rows = self.dynamicCall('GetRepeatCnt(str, str)', sTrCode, sRQName)
            cnt = 0
            account_stock = self.account_stock_dict

            for i in range(rows):
                code = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "종목번호").strip()[1:]
                stock_name = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "종목명").strip()
                purchase_quant = int(self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "보유수량").strip())
                purchase_price = int(self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "매입가").strip())
                stock_yield = float(self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "수익률(%)").strip())
                current_price = int(self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "현재가").strip())
                total_purchase = int(self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "매입금액").strip())
                possible_quant = int(self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "매매가능수량").strip())

                if code in account_stock:
                    pass
                else:
                    account_stock.update({code: {}})

                account_stock[code].update({"종목명": stock_name})
                account_stock[code].update({"보유수량": purchase_quant})
                account_stock[code].update({"매입가": purchase_price})
                account_stock[code].update({"수익률(%)": stock_yield})
                account_stock[code].update({"현재가": current_price})
                account_stock[code].update({"매입금액": total_purchase})
                account_stock[code].update({"매매가능수량": possible_quant})

                cnt += 1

            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()

            print(f"총 보유 종목 수 {cnt}개")
            if self.account_stock_dict:
                print("보유종목", json.dumps(account_stock, indent=4, sort_keys=True, ensure_ascii=False))
            else:
                print("보유종목 없음")

            print(f"총매입금액 {int(total_balance)}원")
            print(f"총수익률 {float(total_profit_rate)}%")

        elif sRQName == "실시간미체결요청":
            rows = self.dynamicCall('GetRepeatCnt(str, str)', sTrCode, sRQName)
            unsigned_order = self.unsigned_order_dict

            for i in range(rows):
                code = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "종목번호").strip()[1:]
                stock_name = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "종목명").strip()
                order_no = int(self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "주문번호").strip())
                order_status = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "주문상태").strip()
                order_quant = int(self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "주문수량").strip())
                order_price = int(self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "주문가격").strip())
                order_dist = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "주문구분").strip()
                unsigned_quant = int(self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "미체결수량").strip())
                signed_quant = int(self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "체결량").strip())

                unsigned_order = self.unsigned_order_dict

                if order_no in unsigned_order:
                    pass
                else:
                    self.unsigned_order_dict

                unsigned_order.update({"종목번호": code})
                unsigned_order.update({"종목명": stock_name})
                unsigned_order.update({"주문번호": order_no})
                unsigned_order.update({"주문상태": order_status})
                unsigned_order.update({"주문수량": order_quant})
                unsigned_order.update({"주문가격": order_price})
                unsigned_order.update({"주문구분": order_dist})
                unsigned_order.update({"미체결수량": unsigned_quant})
                unsigned_order.update({"체결량": signed_quant})

                print(f"미체결종목:{unsigned_order[order_no]}")

            self.detail_account_info_event_loop.exit()

        elif sRQName == "주식일봉차트조회":
            code = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, 0, "종목코드").strip()
            print(f"코드번호 : {code} 일봉데이터조회...")

            cnt = self.dynamicCall("GetRepeatCnt(string,string", sTrCode, sRQName)
            print(f'남은 일수 : {cnt}')
            # data = ['', 현재가, 거래량, 거래대금, 날짜, 시가, 고가, 저가, '']
            for i in range(cnt):
                data = []

                price = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "현재가").strip()
                value = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "거래량").strip()
                transaction_amount = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "거래대금").strip()
                date = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "일자").strip()
                init_price = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "시가").strip()
                high_price = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "고가").strip()
                low_price = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "저가").strip()

                data.append("")
                data.append(price)
                data.append(value)
                data.append(transaction_amount)
                data.append(date)
                data.append(init_price)
                data.append(high_price)
                data.append(low_price)
                data.append("")

                self.calcul_data.append(data.copy())

            if sPrevNext == "2":
                self.day_graph(code=code, sPrevNext=sPrevNext)
            else:
                print(f"총 일수 {len(self.calcul_data)}")

                pass_success = False

                # 120일 이평선 확인하기
                if self.calcul_data is None or len(self.calcul_data) < 120:
                    pass_success = False
                else:
                    total_price = 0
                    for value in self.calcul_data[:120]:
                        total_price += int(value[1])

                    quarter_moving_average_price = total_price / 120

                    # 저가 <= 120이평선 & 고가 <= 120이평선 조건 확인
                    bottom_price = False
                    check_price = None
                    if int(self.calcul_data[0][7]) <= quarter_moving_average_price <= int(self.calcul_data[0][6]):
                        bottom_price = True
                        check_price = int(self.calcul_data[0][6])

                    prev_price = None
                    if bottom_price:
                        quarter_moving_average_price = 0
                        price_top_moving = False

                        idx = 1
                        while True:
                            if len(self.calcul_data[idx:]) < 120:   # 120일치가 있는지 계속 확인
                                print("120일치의 데이터가 없음")
                                break

                            total_price = 0
                            for value in self.calcul_data[idx:120+idx]:
                                total_price += int(value[1])
                            quarter_moving_average_price_prev = total_price / 120

                            if quarter_moving_average_price_prev <= int(self.calcul_data[idx][6]) and idx <= 20:
                                print("20일 동안 주가가 120일 이평선과 같거나 위에 있으면 조건 불충분")
                                price_top_moving = False
                                break

                            elif int(self.calcul_data[idx][7]) > quarter_moving_average_price_prev and idx > 20:
                                print("일봉이 120이평선 위로 확인됨")
                                price_top_moving = True
                                prev_price = int(self.calcul_data[idx][7])
                                break

                            idx += 1

                        # 이평선이 가장 최근 일자 이평선의 가격보다 낮은지 확인하기
                        if price_top_moving is True:
                            if quarter_moving_average_price > quarter_moving_average_price_prev and check_price > prev_price:
                                print("포착된 이평선의 가격이 오늘자 이평선 가격보다 낮은 것 확인됨")
                                print("포착된 일봉 저가가 금일 일봉 고가보다 낮은 것 확인됨")
                                pass_success = True

                if pass_success:
                    print("조건부 통과됨")

                    code_nm = self.dynamicCall("GetMasterCodeName(QString)", code)

                    f = open("files/condition_stock.txt", "a", encoding="utf8")
                    f.write(f"{code}\t{code_nm}\t{str(self.calcul_data[0][1])}\n")
                    f.close()

                else:
                    print("조건부 통과 못함")

                self.calcul_data.clear()
                self.calculator_event_loop.exit()


