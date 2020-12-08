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
        ###################

        ##########계좌 관련변수
        self.use_money = 0
        self.use_money_percent = 0.5
        ####################

        ###이벤트 루프 모음
        self.login_event_loop = None
        self.detail_account_info_event_loop = None
        self.detail_account_info_event_loop_2 = QEventLoop()
        ###################

        self.get_ocx_instance()
        self.event_slot()

        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info() #예수금 정보 가져오기
        self.detail_account_mystock()   #계좌평가 잔고 내역


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
        self.dynamicCall("CommRqData(String, String, int, String)","예수금상세현황요청","opw00001", "0", "2000")

        self.detail_account_info_event_loop = QEventLoop()
        self.detail_account_info_event_loop.exec_()
    
    def detail_account_mystock(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)","계좌평가잔고내역","opw00018", sPrevNext, "2000")

        
        self.detail_account_info_event_loop_2.exec_()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        
        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(String, String, int, String)",sTrCode, sRQName, 0, "예수금")
            print("예수금: {}".format(int(deposit)))

            self.use_money = int(deposit) * self.use_money_percent
            self.use_money = int(self.use_money / 4)

            ok_deposit = self.dynamicCall("GetCommData(String, String, int, String)",sTrCode, sRQName, 0, "출금가능금액")
            print("출금가능금액: {}".format(int(ok_deposit)))

            self.detail_account_info_event_loop.exit()
        
        if sRQName == "계좌평가잔고내역":
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
                self.detail_account_info_event_loop_2.exit()