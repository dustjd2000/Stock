from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *
from config.kiwoomType import *
from Manage.Mail import SendMail
import time
from logManage.logManager import LogManager

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        print("Kiwoom init()")

        self.log = LogManager()

        self.realtype = RealType()

        self.objMail = SendMail()

        ###########변수모음
        self.account_num = None

        self.accout_stock_dict = {}
        self.not_account_stock_dict = {}

        self.will_account_stock_code = {}
        self.sell_account_stock_dict = {}
        self.sell_success_stock_dict = {}
        ###################

        ##########계좌 관련변수
        self.use_money_origin = 0 # 보유 예수금
        self.use_money = 0  # 보유 예수금 주식주문 비용
        self.use_money_percent = 0.5    # 예수금 중 주식주문 사용 비율
        self.use_up_down_rate_percent = 7 # 신고가 조회 등락율 %
        self.use_sell_order_rate = 0.04 # 매도 주문 조건 등락율 *100 %
        self.use_buy_price_rate = 0.3 # 매수 주문 현재가 + 비율
        ####################

        ###이벤트 루프 모음
        self.login_event_loop = None
        self.detail_account_info_event_loop = QEventLoop()
        ###################

        #############스크린번호 모음
        self.screen_my_info = "2000"
        self.screen_start_stop_real = "1000"
        ########################

        try:
            self.get_ocx_instance()
            self.event_slot()

            self.signal_login_commConnect()
            self.get_account_info()
            self.detail_account_info() #예수금 정보 가져오기
            self.detail_account_mystock()   #계좌평가 잔고 내역
            self.not_concluded_account() #미체결정보 확인
            self.new_high_stock() #신고가 조회
            self.Send_Buy_Order() # 매수 주문

            # 실시간 
            self.dynamicCall("SetRealReg(QString, QString, QString, QString)", self.screen_start_stop_real, '', self.realtype.REALTYPE["주문체결"]["주문상태"], "0")
            #self.Send_Sell_Order() # 매도 주문
        except Exception as ex:
            subject = "kiwoom 자동주식 매매 실패"
            msg = ex

            self.log.logPrint("kiwoom 자동주식 매매 실패 cause: {}".format(ex))
            self.objMail.SendMailMsgSet(subject, msg)

        
        while True:
            now = time.localtime()
            hour = int(now.tm_hour)
            min = int(now.tm_min)

            if hour == 9 and min == 40:
                self.log.logPrint("{}시{}분 주식매매 종료".format(str(hour), str(min)))
                break
        
        self.Send_Sell_Sucess_Mail()



    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")
    
    def event_slot(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)
    
    def real_event_slot(self):
        self.OnReceiveRealData.connect(self.realdata_slot)
    
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

        msg = "나의 보유계좌번호: " + self.account_num
        self.log.logPrint(msg)

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
        self.log.logPrint("매수 주문 시작")

        result = self.use_money / self.will_account_stock_code["현재가"]
        quantity = int(result)

        buy_price = self.will_account_stock_code["현재가"] + int(self.will_account_stock_code["현재가"] * self.use_buy_price_rate)

        order_success = self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                            ["신규매수", self.screen_my_info, self.account_num, 1, self.will_account_stock_code["종목코드"], quantity, self.will_account_stock_code["현재가"],buy_price, ""]
                            )
        
        if order_success == 0:
            self.log.logPrint("매수주문 성공")
        else :
            self.log.logPrint("매수주문 실패")
        
        self.log.logPrint("종목명: {}".format(self.will_account_stock_code["종목명"]))
        self.log.logPrint("종목코드: {}".format(self.will_account_stock_code["종목코드"]))
        self.log.logPrint("현재가: {}".format(self.will_account_stock_code["현재가"]))
        self.log.logPrint("buy_price: {}".format(buy_price))
        self.log.logPrint("매수개수: {}".format(quantity))
        self.log.logPrint("매수 주문 종료")
    
    def Send_Sell_Order(self):
        self.log.logPrint("매도 주문 시작")

        # 계좌평가 잔고내역 조회
        self.accout_stock_dict = {}  # 계좌평가 잔고내역 종목정보 초기화
        self.detail_account_mystock()   #계좌평가 잔고 내역 조회

        for key, value in self.accout_stock_dict:
            # 등락율 매도가격
            sell_price = self.accout_stock_dict[key]["매입가"] + int(self.accout_stock_dict[key]["매입가"] * self.use_sell_order_rate)
            quantity = self.accout_stock_dict[key]["보유수량"]

            # 매도
            order_success = self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                "신규매도", self.screen_my_info, self.account_num, 2, self.accout_stock_dict[key]["종목코드"], quantity, sell_price, "03", ""
                                )
            
            if order_success == 0:
                self.log.logPrint("매도주문 성공")
            else :
                self.log.logPrint("매도주문 실패")
            
            self.log.logPrint("*********************************************")
            self.log.logPrint("종목명: {}".format(self.accout_stock_dict[key]["종목명"]))
            self.log.logPrint("종목코드: {}".format(self.accout_stock_dict[key]["종목코드"]))
            self.log.logPrint("현재가: {}".format(self.accout_stock_dict[key]["현재가"]))
            self.log.logPrint("sell_price: {}".format(sell_price))
            self.log.logPrint("매도개수: {}".format(quantity))
        self.log.logPrint("매도 주문 끝")
    
    def Get_Real_MyAccount(self):
        self.dynamicCall("SetRealReg(QString, QString, QString, QString)", self.screen_start_stop_real, '', self.realtype.REALTYPE["잔고"]["계좌번호"], "1")
        #self.dynamicCall("SetRealReg(QString, QString, QString, QString)", self.screen_start_stop_real, '', self.realtype.REALTYPE["잔고"]["예수금"], "2")
        #self.dynamicCall("SetRealReg(QString, QString, QString, QString)", self.screen_start_stop_real, '', self.realtype.REALTYPE["잔고"]["손익율"], "3")
   
    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        
        if sRQName == "예수금상세현황요청":
            self.log.logPrint("예수금상세현황요청")
            deposit = self.dynamicCall("GetCommData(String, String, int, String)",sTrCode, sRQName, 0, "예수금")
            self.log.logPrint("예수금: {}".format(int(deposit)))

            self.use_money_origin = int(deposit)

            self.use_money = int(deposit) * self.use_money_percent
            #self.use_money = int(self.use_money / 4)

            ok_deposit = self.dynamicCall("GetCommData(String, String, int, String)",sTrCode, sRQName, 0, "출금가능금액")
            self.log.logPrint("출금가능금액: {}".format(int(ok_deposit)))

            self.detail_account_info_event_loop.exit()
        
        elif sRQName == "계좌평가잔고내역":
            self.log.logPrint("계좌평가잔고내역")
            total_buy_money = self.dynamicCall("GetCommData(String, String, int, String)",sTrCode, sRQName, 0, "총매입금액")
            total_buy_money = int(total_buy_money)

            self.log.logPrint("총매입금액: {}".format(total_buy_money))
            
            total_profit_loss_rate = self.dynamicCall("GetCommData(String, String, int, String)",sTrCode, sRQName, 0, "총수익률(%)")
            total_profit_loss_rate_result = float(total_profit_loss_rate)

            self.log.logPrint("총수익률(%): {}".format(total_profit_loss_rate_result))

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

            self.log.logPrint("계좌에 가지고 있는 종목: {}".format(cnt))
            self.log.logPrint("보유계좌 정보: {}".format(self.accout_stock_dict))

            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()
        
        elif sRQName == "실시간미체결요청":
            self.log.logPrint("실시간미체결요청")
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
                
                self.log.logPrint("미체결 종목: {}".format(self.not_account_stock_dict[order_no]))
                cnt += 1

            self.log.logPrint("미체결종목count: {}".format(cnt))       
            
            self.detail_account_info_event_loop.exit()
        
        elif sRQName == "신고가요청":
            self.log.logPrint("신고가요청")
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

                        self.log.logPrint("신고가 종목: {}".format(self.will_account_stock_code))
                        cnt += 1

            self.log.logPrint("신고가count: {}".format(cnt))  
            
            self.detail_account_info_event_loop.exit()
    
    def realdata_slot(self, sCode, sRealType, sRealData):
       
       #실시간 데이터 조회

        if sRealType == "주문체결":
           self.log.logPrint("실시간 주문체결")
           code = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realtype.REALTYPE["주문체결"]["종목코드"])
           code_nm = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realtype.REALTYPE["주문체결"]["종목명"])
           order_state = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realtype.REALTYPE["주문체결"]["주문상태"])
           result_price = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realtype.REALTYPE["주문체결"]["체결누계금액"])
           order_gubun = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realtype.REALTYPE["주문체결"]["매도수구분"])

           code = code.strip()
           code_nm = code.strip()
           order_state = order_state.strip()
           result_price = result_price.strip()
           order_gubun = self.realtype.REALTYPE["매도수구분"][order_gubun]

           if order_gubun == "매수" and order_state == "체결" : # 매수 체결 성공
               # 매도 시도
               self.log.logPrint("실시간 매수 체결 성공, 매도 시도")
               self.Send_Sell_Order()
           elif order_gubun == "매도" and order_state == "체결" :   # 매도 체결 성공
                # 매도 정보 저장
                self.log.logPrint("실시간 매도 체결 성공, 매도정보 저장")
                if code in self.sell_success_stock_dict:
                    pass
                else:
                    self.sell_success_stock_dict.update({code:{}})
                    self.sell_success_stock_dict[code].update({"종목코드": code})
                    self.sell_success_stock_dict[code].update({"종목명": code_nm})
                    self.sell_success_stock_dict[code].update({"주문상태": order_state})
                    self.sell_success_stock_dict[code].update({"체결누계금액": result_price})
                    self.sell_success_stock_dict[code].update({"매도수구분": order_gubun})
                    
    def Send_Sell_Sucess_Mail(self):

        self.detail_account_info() #예수금 정보 가져오기

        total_money = "예수금: {}".format(self.use_money_origin) + "\n\n"
        
        msg = "내역 : \n"
        msg += "**********************************" 
        for key, value in self.sell_success_stock_dict:
            msg += "종목코드: {}".format(self.sell_success_stock_dict[key]["종목코드"]) + "\n"
            msg += "종목명: {}".format(self.sell_success_stock_dict[key]["종목명"]) + "\n"
            msg += "주문상태: {}".format(self.sell_success_stock_dict[key]["주문상태"]) + "\n"
            msg += "체결누계금액: {}".format(self.sell_success_stock_dict[key]["체결누계금액"]) + "\n"
            msg += "매도수구분: {}".format(self.sell_success_stock_dict[key]["매도수구분"]) + "\n"
            msg += "**********************************" + "\n" 

        subject = "Kiwoom 자동주식매매 Sell_Success"
        sendmsg = total_money + msg

        self.log.logPrint("Send_Sell_Success_Mail()")
        self.log.logPrint(sendmsg)

        self.objMail.SendMailMsgSet(subject, sendmsg)