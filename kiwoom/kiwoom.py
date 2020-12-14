from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        print("Kiwoom init()")

        ###########변수모음
        self.account_num = None

        self.accout_stock_dict = {}
        self.not_account_stock_dict = {}

        self.will_account_stock_code = {}
        self.sell_account_stock_dict = {}
        ###################

        ##########계좌 관련변수
        self.use_money = 0  # 보유 예수금
        self.use_money_percent = 0.5    # 예수금 중 주식주문 사용 비율
        self.use_up_down_rate_percent = 7 # 신고가 조회 등락율 %
        self.use_send_order_rate = 3 # 매도 주문 조건 등락율 %
        self.use_buy_price_rate = 0.3 # 매수 주문 현재가 + 비율
        ####################

        ###이벤트 루프 모음
        self.login_event_loop = None
        self.detail_account_info_event_loop = QEventLoop()
        ###################

        #############스크린번호 모음
        self.screen_my_info = "2000"
        ########################

        self.get_ocx_instance()
        self.event_slot()

        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info() #예수금 정보 가져오기
        self.detail_account_mystock()   #계좌평가 잔고 내역
        self.not_concluded_account() #미체결정보 확인
        self.new_high_stock() #신고가 조회
        self.Send_Buy_Order() # 매수 주문
        self.Send_Sell_Order() # 매도 주문





    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")
    
    def event_slot(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)
    
    
    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()
    
    def login_slot(self, errCode):
        print(errors(errCode))
        
        self.login_event_loop.exit()
    
    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(String)","ACCNO")
        
        self.account_num = account_list.split(';')[0]

        print("나의 보유계좌번호: {}".format(self.account_num))

    def detail_account_info(self):
        print("예수금 가져오기")
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)","예수금상세현황요청","opw00001", "0", self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

    def detail_account_mystock(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)","계좌평가잔고내역","opw00018", sPrevNext, self.screen_my_info)

        
        self.detail_account_info_event_loop.exec_()

    def not_concluded_account(self, sPrevNext="0"):
        print("미체결정보 요청")
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "체결구분", "1")
        self.dynamicCall("SetInputValue(String, String)", "매매구분", "0")
        self.dynamicCall("CommRqData(String, String, int, String)","실시간미체결요청","opt10075", sPrevNext, self.screen_my_info)

        self.detail_account_info_event_loop.exec_()
    
    def new_high_stock(self, sPrevNext="0"):
        print("신고가정보 조회")
        self.dynamicCall("SetInputValue(String, String)", "시장구분", "000")
        self.dynamicCall("SetInputValue(String, String)", "신고저구분", "1")
        self.dynamicCall("CommRqData(String, String, int, String)","신고가요청","OPT10016", sPrevNext, self.screen_my_info)

        self.detail_account_info_event_loop.exec_()
    
    def Send_Buy_Order(self):
        print("매수 주문 시작")

        result = self.use_money / self.will_account_stock_code["현재가"]
        quantity = int(result)

        buy_price = self.will_account_stock_code["현재가"] + int(self.will_account_stock_code["현재가"] * self.use_buy_price_rate)

        order_succest = self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                            ["신규매수", self.screen_my_info, self.account_num, 1, self.will_account_stock_code["종목코드"], quantity, self.will_account_stock_code["현재가"],buy_price, ""]
                            )
        
        if order_succest == 0:
            print("매수주문 성공")
        else :
            print("매수주문 실패")
        
        print("종목명: {}".format(self.will_account_stock_code["종목명"]))
        print("종목코드: {}".format(self.will_account_stock_code["종목코드"]))
        print("현재가: {}".format(self.will_account_stock_code["현재가"]))
        print("buy_price: {}".format(buy_price))
        print("매수개수: {}".format(quantity))
    
    def Send_Sell_Order(self):
        print("매도 주문 시작")

        # 계좌평가 잔고내역 조회
        # 등락율 비교
        # 매도
    
    def up_down_rate_stock(self, now_price, buy_price):
        meme_rate = ((now_price - buy_price)/buy_price) * 100

        if meme_rate > self.use_send_order_rate or meme_rate < self.use_send_order_rate * -1 :
            return True
        else :
            return False


    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        
        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(String, String, int, String)",sTrCode, sRQName, 0, "예수금")
            print("예수금: {}".format(int(deposit)))

            self.use_money = int(deposit) * self.use_money_percent
            #self.use_money = int(self.use_money / 4)

            ok_deposit = self.dynamicCall("GetCommData(String, String, int, String)",sTrCode, sRQName, 0, "출금가능금액")
            print("출금가능금액: {}".format(int(ok_deposit)))

            self.detail_account_info_event_loop.exit()
        
        elif sRQName == "계좌평가잔고내역":
            total_buy_money = self.dynamicCall("GetCommData(String, String, int, String)",sTrCode, sRQName, 0, "총매입금액")
            total_buy_money = int(total_buy_money)

            print("총매입금액: {}".format(total_buy_money))
            
            total_profit_loss_rate = self.dynamicCall("GetCommData(String, String, int, String)",sTrCode, sRQName, 0, "총수익률(%)")
            total_profit_loss_rate_result = float(total_profit_loss_rate)

            print("총수익률(%): {}".format(total_profit_loss_rate_result))

            rows = self.dynamicCall("GetRepeatCnt(QString, QString)",sTrCode, sRQName) #최대조회개수 20개
            cnt = 0

            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "종목번호")
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "종목명")
                stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "보유수량")
                buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "매입가")
                learn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "수익률(%)")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "현재가")
                total_chegual_price = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "매입금액")
                possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "매매가능수량")

                if code in self.accout_stock_dict:
                    pass
                else:
                    self.accout_stock_dict.update({code:{}})

                code = code.strip()[1:]
                code_nm = code_nm.strip()
                stock_quantity = int(stock_quantity.strip())
                buy_price = int(buy_price.strip())
                learn_rate = float(learn_rate.strip())
                current_price = int(current_price.strip())
                total_chegual_price = int(total_chegual_price.strip())
                possible_quantity = int(possible_quantity.strip())

                self.accout_stock_dict[code].update({"종목명": code_nm})
                self.accout_stock_dict[code].update({"보유수량": stock_quantity})
                self.accout_stock_dict[code].update({"매입가": buy_price})
                self.accout_stock_dict[code].update({"수익률(%)": learn_rate})
                self.accout_stock_dict[code].update({"현재가": current_price})
                self.accout_stock_dict[code].update({"매입금액": total_chegual_price})
                self.accout_stock_dict[code].update({"매매가능수량": possible_quantity})

                cnt += 1

            print("계좌에 가지고 있는 종목: {}".format(cnt))
            print("보유계좌 정보: {}".format(self.accout_stock_dict))

            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()
        
        elif sRQName == "실시간미체결요청":
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)",sTrCode, sRQName) #최대조회개수 20개
            cnt = 0
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "종목번호")
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "종목명")
                order_no = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "주문번호")
                order_status = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "주문상태")
                order_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "주문수량")
                order_price = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "주문가격")
                order_gubun = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "주문구분")
                not_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "미체결수량")
                ok_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "체결량")

                code = code.strip()
                code_nm = code_nm.strip()
                order_no = int(order_no.strip()) 
                order_status = order_status.strip()
                order_quantity = int(order_quantity.strip())
                order_price = int(order_price.strip())
                order_gubun = order_gubun.strip().lstrip('+').lstrip('-')
                not_quantity = int(not_quantity.strip())
                ok_quantity = int(ok_quantity.strip())

                if order_no in self.not_account_stock_dict:
                    pass
                else:
                    self.not_account_stock_dict[order_no] = {}
                
                self.not_account_stock_dict[order_no].update({"종목코드": code})
                self.not_account_stock_dict[order_no].update({"종목명": code_nm})
                self.not_account_stock_dict[order_no].update({"주문번호": order_no})
                self.not_account_stock_dict[order_no].update({"주문상태": order_status})
                self.not_account_stock_dict[order_no].update({"주문수량": order_quantity})
                self.not_account_stock_dict[order_no].update({"주문가격": order_price})
                self.not_account_stock_dict[order_no].update({"주문구분": order_gubun})
                self.not_account_stock_dict[order_no].update({"미체결수량": not_quantity})
                self.not_account_stock_dict[order_no].update({"체결량": ok_quantity})
                
                print("미체결 종목: {}".format(self.not_account_stock_dict[order_no]))
                cnt += 1

            print("미체결종목count: {}".format(cnt))       
            
            self.detail_account_info_event_loop.exit()
        
        elif sRQName == "신고가요청":
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)",sTrCode, sRQName) #최대조회개수 20개
            cnt = 0
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "종목코드")
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "종목명")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "현재가")
                up_down_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "등락률")
                trade_count = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "거래량")

                code = code.strip()
                code_nm = code_nm.strip()
                current_price = int(current_price.strip())
                up_down_rate = up_down_rate.strip()
                trade_count = trade_count.strip()

                if '+' in up_down_rate:     #등락률 +
                    up_down_rate_temp = int(float(up_down_rate[1:]))
                    
                    if up_down_rate_temp == self.use_up_down_rate_percent or up_down_rate_temp == self.use_up_down_rate_percent +1 or up_down_rate_temp == self.use_up_down_rate_percent +2:
                        self.will_account_stock_code.update({"종목코드": code})
                        self.will_account_stock_code.update({"종목명": code_nm})
                        self.will_account_stock_code.update({"현재가": current_price})
                        self.will_account_stock_code.update({"등락률": up_down_rate})
                        self.will_account_stock_code.update({"거래량": trade_count})

                        print("신고가 종목: {}".format(self.will_account_stock_code))
                        cnt += 1

            print("신고가count: {}".format(cnt))       
            
            self.detail_account_info_event_loop.exit()