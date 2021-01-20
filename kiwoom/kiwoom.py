from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *
from config.kiwoomType import *
from Manage.Mail import SendMail
import time
from logManage.logManager import LogManager
import sys
import os

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

        self.will_account_stock_code = {}   # 신고가 조회
        self.sell_account_stock_dict = {}   # 매도 시도
        self.sell_success_stock_dict = {}   # 매도 체결성공

        self.will_account_stock_code_finish = [] # 매수 성공 리스트
        self.sell_success_stock_dict_finish = [] # 매도 성공 리스트
        ###################

        ##########계좌 관련변수
        self.use_money_origin = 0 # 보유 예수금
        self.use_money = 0  # 보유 예수금 주식주문 비용
        self.use_money_percent = 0.5    # 예수금 중 주식주문 사용 비율
        self.use_up_down_rate_percent = 7 # 신고가 조회 등락율 %
        self.use_sell_order_rate = 0.04 # 매도 주문 조건 등락율 *100 %
        self.use_buy_price_rate = 2 # 매수 주문  - 현재가 * 비율
        
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
            self.real_event_slot()

            self.signal_login_commConnect()
            self.get_account_info()
            self.detail_account_info(self.screen_my_info) #예수금 정보 가져오기
            self.detail_account_mystock()   #계좌평가 잔고 내역
            #self.not_concluded_account() #미체결정보 확인

            #장 시작, 끝 확인
            self.dynamicCall("SetRealReg(QString, QString, QString, QString)", self.screen_start_stop_real, '', self.realtype.REALTYPE["장시작시간"]["장운영구분"], "0")

            while(True):
                #self.new_high_stock() #신고가 
                self.high_stock() #가격급등락
                
                if self.will_account_stock_code.keys() :
                    self.merge_sell_account()
                    break
                else:
                    time.sleep(30)
            
            self.Send_Buy_Order() # 매수 주문

            # 특정 종목 실시간 
            #self.dynamicCall("SetRealReg(QString, QString, QString, QString)", self.screen_start_stop_real, '', self.realtype.REALTYPE["주문체결"]["주문상태"], "0")
            #self.Send_Sell_Order() # 매도 주문

            
            """
            while True:
                now = time.localtime()
                hour = int(now.tm_hour)
                min = int(now.tm_min)

                if hour == 9 and min == 40:
                    self.log.logPrint("{}시{}분 주식매매 종료".format(str(hour), str(min)))
                    break
            

            self.Send_Sell_Sucess_Mail()
            """

        except Exception as ex:
            subject = "kiwoom 자동주식 매매 실패"
            msg = ex

            self.log.logPrint("kiwoom 자동주식 매매 실패 cause: {}".format(msg))
            #self.objMail.SendMailMsgSet(subject, msg)
        

        #exit()

        #sys.exit()



    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")
    
    def event_slot(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)
    
    def real_event_slot(self):
        self.OnReceiveRealData.connect(self.realdata_slot)  # 특정 종목 실시간 정보 조회
        self.OnReceiveChejanData.connect(self.chejan_slot)
        self.OnReceiveMsg.connect(self.receiveMsg)
    
    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()
    
    def login_slot(self, errCode):
        if 0 == int(str(errCode)):
            self.log.logPrint("로그인 정상접속")
        else:
            self.log.logPrint("로그인 실패: {}".format(str(errors(errCode)[1:])) )
            self.objMail.SendMailMsgSet("kiwoom 자동 주식매매 로그인 접속 실패", "원인: {}".format(str(errors(errCode)[1:])))
        
        self.login_event_loop.exit()
    
    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(String)","ACCNO")
        
        self.account_num = account_list.split(';')[0]

        msg = "나의 보유계좌번호: " + self.account_num
        self.log.logPrint(msg)

    def detail_account_info(self, screen_num):
        print("예수금 가져오기")
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)","예수금상세현황요청","opw00001", "0", screen_num)

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
    
    def high_stock(self, sPrevNext="0"):
        print("가격급등 조회")
        self.dynamicCall("SetInputValue(String, String)", "시장구분", "000")
        self.dynamicCall("SetInputValue(String, String)", "등락구분", "1") # 1: 급등, 2: 급락
        self.dynamicCall("SetInputValue(String, String)", "시간구분", "1") # 1: 분전, 2: 일전
        self.dynamicCall("SetInputValue(String, String)", "거래량구분", "00050") # 5만주이상
        self.dynamicCall("CommRqData(String, String, int, String)","가격급등락요청","opt10019", sPrevNext, self.screen_my_info)

        self.detail_account_info_event_loop.exec_()
    
    def Send_Buy_Order(self):
        self.log.logPrint("##########신규 매수 주문 시작###########")

        if self.will_account_stock_code["종목코드"] in self.will_account_stock_code_finish:
            return

        self.log.logPrint("매도주문정보: {}".format(self.will_account_stock_code))

        hoga = self.hogaUnitCalc( int(self.will_account_stock_code["현재가"]) )

        buy_price = self.will_account_stock_code["현재가"] + (hoga * self.use_buy_price_rate)

        won_1 = buy_price % 10

        buy_price = buy_price - won_1
        
        result = self.use_money / buy_price
        quantity = int(result)

        order_success = self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                            ["신규매수", self.screen_my_info, self.account_num, self.realtype.REALTYPE["주문유형"]["신규매수"], 
                            self.will_account_stock_code["종목코드"], quantity, buy_price, self.realtype.REALTYPE["거래구분"]["지정가"], ""]
                            )
        
        if order_success == 0:
            self.log.logPrint("신규매수 주문전달 성공")
        else :
            self.log.logPrint("신규매수 주문전달 실패")
        
        self.log.logPrint("종목명: {}".format(self.will_account_stock_code["종목명"]))
        self.log.logPrint("종목코드: {}".format(self.will_account_stock_code["종목코드"]))
        self.log.logPrint("현재가: {}".format(self.will_account_stock_code["현재가"]))
        self.log.logPrint("buy_price: {}".format(buy_price))
        self.log.logPrint("매수개수: {}".format(quantity))

        self.will_account_stock_code_finish.append(self.will_account_stock_code["종목코드"])
        self.log.logPrint("#############신규 매수 주문 종료###########")
    
    def Send_Sell_Order(self):
        self.log.logPrint("########매도 주문 시작#########")

        # 계좌평가 잔고내역 조회
        #self.accout_stock_dict = {}  # 계좌평가 잔고내역 종목정보 초기화
        #self.detail_account_mystock()   #계좌평가 잔고 내역 조회

        self.log.logPrint("매도주문정보: {}".format(self.sell_account_stock_dict))

        for sCode in self.sell_account_stock_dict:

            hoga = self.hogaUnitCalc(self.sell_account_stock_dict[sCode]["현재가"])

            hope_price = int(self.sell_account_stock_dict[sCode]["매입단가"]) + \
                int(self.sell_account_stock_dict[sCode]["매입단가"] * self.use_sell_order_rate)

            buyhoga_count = int(
                (hope_price - int(self.sell_account_stock_dict[sCode]["매입단가"])) / hoga)

            sell_price = int(
                 self.sell_account_stock_dict[sCode]["매입단가"]) + (buyhoga_count * hoga)

            won_1 = sell_price % 10

            sell_price = sell_price - won_1

            quantity = self.sell_account_stock_dict[sCode]["보유수량"]

              # 매도
            order_success = self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                                ["신규매도", self.screen_my_info, self.account_num, int(self.realtype.REALTYPE["주문유형"]["신규매도"]),
                                                  self.sell_account_stock_dict[sCode]["종목코드"], quantity, sell_price, self.realtype.REALTYPE["거래구분"]["지정가"], ""]
                                                )

            if order_success == 0:
                self.log.logPrint("매도주문전달 성공")
            else:
                self.log.logPrint("매도주문전달 실패")

            self.log.logPrint(
                    "*********************************************")
            self.log.logPrint("종목명: {}".format(self.sell_account_stock_dict[sCode]["종목명"]))
            self.log.logPrint("종목코드: {}".format(self.sell_account_stock_dict[sCode]["종목코드"]))
            self.log.logPrint("현재가: {}".format(self.sell_account_stock_dict[sCode]["현재가"]))
            self.log.logPrint("sell_price: {}".format(sell_price))
            self.log.logPrint("매도개수: {}".format(quantity))

            self.sell_success_stock_dict_finish.append(sCode)
        self.log.logPrint("#########매도 주문 끝#########")
    
    def stop_screen_cancel(self, sScrNo):
        self.dynamicCall("DisconnectRealData(QString)", sScrNo)
   
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

            #self.stop_screen_cancel(self.screen_my_info)

            self.detail_account_info_event_loop.exit()
        
        elif sRQName == "계좌평가잔고내역":
            self.log.logPrint("계좌평가잔고내역")
            
            total_buy_money = self.dynamicCall("GetCommData(String, String, int, String)",sTrCode, sRQName, 0, "총매입금액")
            if total_buy_money == '':
                total_buy_money = 0
            total_buy_money = int(total_buy_money)

            self.log.logPrint("총매입금액: {}".format(total_buy_money))
            
            total_profit_loss_rate = self.dynamicCall("GetCommData(String, String, int, String)",sTrCode, sRQName, 0, "총수익률(%)")
            if total_profit_loss_rate == '':
                total_profit_loss_rate = 0.00
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

                code = code.strip()[1:]
                code_nm = code_nm.strip()
                
                stock_quantity = stock_quantity.strip()
                if stock_quantity == '':
                    stock_quantity = 0
                stock_quantity = int(stock_quantity)

                buy_price = buy_price.strip()
                if buy_price == '':
                    buy_price = 0
                buy_price = int(buy_price)

                learn_rate = learn_rate.strip()
                if learn_rate == '':
                    learn_rate = 0.00
                learn_rate = float(learn_rate)

                current_price = current_price.strip()
                
                if current_price == '':
                    current_price = 0
                current_price = int(current_price)
                current_price = abs(current_price)

                total_chegual_price = total_chegual_price.strip()
                if total_chegual_price == '':
                    total_chegual_price = 0
                total_chegual_price = int(total_chegual_price)

                possible_quantity = possible_quantity.strip()
                if possible_quantity == '':
                    possible_quantity = 0
                possible_quantity = int(possible_quantity)

                if code in self.accout_stock_dict:
                    pass
                else:
                    self.accout_stock_dict[code] = {}

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
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "현재가")

                code = code.strip()
                code_nm = code_nm.strip()
                order_no = int(order_no.strip()) 
                order_status = order_status.strip()
                order_quantity = int(order_quantity.strip())
                order_price = int(order_price.strip())
                order_gubun = order_gubun.strip().lstrip('+').lstrip('-')
                not_quantity = int(not_quantity.strip())
                ok_quantity = int(ok_quantity.strip())
                current_price = int(current_price.strip())
                current_price = abs(current_price)

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
                self.not_account_stock_dict[order_no].update({"현재가": current_price})
                
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
                current_price = abs(current_price)
                up_down_rate = up_down_rate.strip()
                trade_count = trade_count.strip()

                if '+' in up_down_rate:     #등락률 +
                    up_down_rate_temp = int(float(up_down_rate[1:]))
                    
                    if up_down_rate_temp == self.use_up_down_rate_percent or up_down_rate_temp == self.use_up_down_rate_percent +1 or up_down_rate_temp == self.use_up_down_rate_percent +2:
                        if self.use_money > current_price and current_price < 100000:
                            
                            if code in self.will_account_stock_code:
                                continue
                            else:
                                self.will_account_stock_code.update({"종목코드": code})
                                self.will_account_stock_code.update({"종목명": code_nm})
                                self.will_account_stock_code.update({"현재가": current_price})
                                self.will_account_stock_code.update({"등락률": up_down_rate})
                                self.will_account_stock_code.update({"거래량": trade_count})

                                self.log.logPrint("신고가 종목: {}".format(self.will_account_stock_code))
                                cnt += 1

            self.log.logPrint("신고가count: {}".format(cnt))  
            
            self.detail_account_info_event_loop.exit()
        
        elif sRQName == "가격급등락요청":
            self.log.logPrint("가격급등락요청")
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)",sTrCode, sRQName) #최대조회개수 20개
            cnt = 0
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "종목코드")
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "종목명")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "현재가")
                up_down_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "등락률")
                trade_count = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "거래량")
                high_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode, sRQName, i, "급등률")

                code = code.strip()
                code_nm = code_nm.strip()
                current_price = int(current_price.strip())
                current_price = abs(current_price)
                up_down_rate = up_down_rate.strip()
                trade_count = trade_count.strip()
                high_rate = high_rate.strip()

                if '+' in high_rate and '-' not in up_down_rate:     #등락률 +
                    high_rate_temp = int(float(high_rate[1:]))
                    
                    if high_rate_temp == self.use_up_down_rate_percent or high_rate_temp == self.use_up_down_rate_percent +1 or high_rate_temp == self.use_up_down_rate_percent +2:
                        if self.use_money > current_price and current_price < 100000:
                            
                            if code in self.will_account_stock_code:
                                continue
                            else:
                                self.will_account_stock_code.update({"종목코드": code})
                                self.will_account_stock_code.update({"종목명": code_nm})
                                self.will_account_stock_code.update({"현재가": current_price})
                                self.will_account_stock_code.update({"등락률": up_down_rate})
                                self.will_account_stock_code.update({"거래량": trade_count})
                                self.will_account_stock_code.update({"급등률": high_rate})

                                self.log.logPrint("가격급등 종목: {}".format(self.will_account_stock_code))
                                cnt += 1
                                break

            self.log.logPrint("가격급등count: {}".format(cnt))  
            
            self.detail_account_info_event_loop.exit()
    
    def realdata_slot(self, sCode, sRealType, sRealData):

       #해당 종목에 대한 실시간 데이터 조회
       #내가 주문한거에 대한 데이터가 아니라 해당 종목에 남들이한 주문, 현재가 등등의 정보 조회

       if sRealType == "장시작시간":
           fid = self.realtype.REALTYPE["장시작시간"]["장운영구분"]
           value = self.dynamicCall("GetCommRealData(QString, int)", sCode, fid)

           if value == '0':
                self.log.logPrint("장 시작 전")
           elif value == '3':
                self.log.logPrint("장 시작")
           elif value == '2':
                self.log.logPrint("장 종료, 동시호가로 넘어감")
           elif value == '4':
                self.log.logPrint("3시 30분, 장 종료")
                
                os.system("taskkill / f / im cmd.exe")
                #sys.exit()

       """
       if sRealType == "주문체결":
           self.log.logPrint("실시간 주문체결")
           code = self.dynamicCall(
               "GetCommRealData(QString, int)", sCode, self.realtype.REALTYPE["주문체결"]["종목코드"])
           code_nm = self.dynamicCall(
               "GetCommRealData(QString, int)", sCode, self.realtype.REALTYPE["주문체결"]["종목명"])
           order_state = self.dynamicCall(
               "GetCommRealData(QString, int)", sCode, self.realtype.REALTYPE["주문체결"]["주문상태"])
           result_price = self.dynamicCall(
               "GetCommRealData(QString, int)", sCode, self.realtype.REALTYPE["주문체결"]["체결누계금액"])
           order_gubun = self.dynamicCall(
               "GetCommRealData(QString, int)", sCode, self.realtype.REALTYPE["주문체결"]["매도수구분"])

           code = code.strip()
           code_nm = code.strip()
           order_state = order_state.strip()
           result_price = result_price.strip()
           order_gubun = self.realtype.REALTYPE["매도수구분"][order_gubun]

           if order_gubun == "매수" and order_state == "체결":  # 매수 체결 성공
               # 매도 시도
               self.log.logPrint("실시간 매수 체결 성공, 매도 시도")
               #self.Send_Sell_Order()
           elif order_gubun == "매도" and order_state == "체결":   # 매도 체결 성공
               # 매도 정보 저장
               self.log.logPrint("실시간 매도 체결 성공, 매도정보 저장")
               if code in self.sell_success_stock_dict:
                    pass
               else:
                    self.sell_success_stock_dict.update({code: {}})
                    self.sell_success_stock_dict[code].update({"종목코드": code})
                    self.sell_success_stock_dict[code].update({"종목명": code_nm})
                    self.sell_success_stock_dict[code].update(
                        {"주문상태": order_state})
                    self.sell_success_stock_dict[code].update(
                        {"체결누계금액": result_price})
                    self.sell_success_stock_dict[code].update(
                        {"매도수구분": order_gubun})
        """

    def chejan_slot(self, sGubun, nItemCnt, sFIdList):
        #내가 주문한 주문 체결 실시간 정보 조회
        #BSTR sGubun, // 체결구분 접수와 체결시 '0'값, 국내주식 잔고전달은 '1'값, 파생잔고 전달은 '4'
        if int(sGubun) == 0:
            self.log.logPrint("chejan 주문 체결")
            accountnum = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["주문체결"]["계좌번호"])
            code = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["주문체결"]["종목코드"])[1:]
            stock_name = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["주문체결"]["종목명"])
            order_number = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["주문체결"]["주문번호"])
            order_status = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["주문체결"]["주문상태"]) #출력 : 접수, 확인, 체결
            order_quan = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["주문체결"]["주문수량"])
            order_price = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["주문체결"]["주문가격"])
            not_order_quan = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["주문체결"]["미체결수량"])
            order_gubun = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["주문체결"]["주문구분"])    # 출력 : 매도 , 매수
            chegual_price = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["주문체결"]["체결가"])
            
            accountnum = accountnum.strip()
            code = code.strip()
            stock_name = stock_name.strip()
            order_number = order_number.strip()
            order_status = order_status.strip()
            order_quan = order_quan.strip()

            if order_quan == '':
                order_quan = 0
            else :
                order_quan = int(order_quan)

            order_price = order_price.strip()

            if order_price == '':
                order_price = 0
            else :
                order_price = int(order_price)
            
            not_order_quan = not_order_quan.strip()

            if not_order_quan == '':
                not_order_quan = 0
            else :
                not_order_quan = int(not_order_quan)

            order_gubun = order_gubun.strip().lstrip('+').lstrip('-')
            chegual_price = chegual_price.strip()

            if chegual_price == '':
                chegual_price = 0
            else :
                chegual_price = int(chegual_price)

            #매수 체결
            if order_status == "체결" and (order_gubun == "매수" or order_gubun == "매수정정") and code in self.will_account_stock_code_finish:
                self.log.logPrint("########매수 체결 성공#########")
                self.log.logPrint("계좌번호: {}".format(accountnum))
                self.log.logPrint("종목코드: {}".format(code))
                self.log.logPrint("종목명: {}".format(stock_name))
                self.log.logPrint("주문번호: {}".format(order_number))
                self.log.logPrint("주문상태: {}".format(order_status))
                self.log.logPrint("주문구분: {}".format(order_gubun))
                self.log.logPrint("##############################")

                self.will_account_stock_code_finish.remove(code)

                subject = "매수 체결"
                msg = "########매수 체결 성공#########" + "\n"
                msg += "계좌번호: {}".format(accountnum) + "\n"
                msg += "종목코드: {}".format(code) + "\n"
                msg += "종목명: {}".format(stock_name) + "\n"
                msg += "주문번호: {}".format(order_number) + "\n"
                msg += "주문상태: {}".format(order_status) + "\n"
                msg += "주문구분: {}".format(order_gubun) + "\n"
                msg += "##############################" + "\n"
                
                self.objMail.SendMailMsgSet(subject, msg)
                    
                
            #매도 체결
            elif order_status == "체결" and (order_gubun == "매도" or order_gubun == "매도정정") and code in self.sell_success_stock_dict_finish:
                self.log.logPrint("########매도 체결 성공#########")
                self.log.logPrint("계좌번호: {}".format(accountnum))
                self.log.logPrint("종목코드: {}".format(code))
                self.log.logPrint("종목명: {}".format(stock_name))
                self.log.logPrint("주문번호: {}".format(order_number))
                self.log.logPrint("주문상태: {}".format(order_status))
                self.log.logPrint("주문구분: {}".format(order_gubun))
                self.log.logPrint("##############################")

                self.sell_success_stock_dict_finish.remove(code)                
                
                #메일로 발송할 매도 성공 주문 정보 저장
                self.sell_success_stock_dict.update({code: {}})
                self.sell_success_stock_dict[code].update({"종목코드": code})
                self.sell_success_stock_dict[code].update({"종목명": stock_name})
                self.sell_success_stock_dict[code].update(
                    {"주문상태": order_status})
                self.sell_success_stock_dict[code].update(
                    {"체결누계금액": chegual_price * order_quan })
                self.sell_success_stock_dict[code].update(
                    {"매도수구분": order_gubun})

                
        elif int(sGubun) == 1:
            self.log.logPrint("chejan 잔고 조회")

            account_num = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["잔고"]["계좌번호"])
            sCode = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["잔고"]["종목코드"])[1:]
            stock_name = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["잔고"]["종목명"])
            current_price = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["잔고"]["현재가"])
            stoc_quan = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["잔고"]["보유수량"])
            like_quan = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["잔고"]["주문가능수량"])
            buy_price = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["잔고"]["매입단가"])
            total_buy_price = self.dynamicCall(
               "GetChejanData(int)", self.realtype.REALTYPE["잔고"]["총매입가"])
            
            #meme_gubun = self.dynamicCall(
            #   "GetChejanData(int)", self.realtype.REALTYPE["잔고"]["매도/매수구분"])
            
            account_num = account_num.strip()
            sCode = sCode.strip()
            stock_name = stock_name.strip()
            current_price = current_price.strip()

            if current_price == '':
                current_price = 0
            else :
                current_price = abs(int(current_price))
            
            current_price = abs(current_price)

            stoc_quan = stoc_quan.strip()

            if stoc_quan == '':
                stoc_quan = 0
            else :
                stoc_quan = int(stoc_quan)

            like_quan = like_quan.strip()

            if like_quan == '':
                like_quan = 0
            else :
                like_quan = int(like_quan)

            buy_price = buy_price.strip()

            if buy_price == '':
                buy_price = 0
            else :
                buy_price = abs(int(buy_price))

            total_buy_price = total_buy_price.strip()

            if total_buy_price == '':
                total_buy_price = 0
            else :
                total_buy_price = abs(int(total_buy_price))

            #meme_gubun = meme_gubun.strip()
            #meme_gubun = self.realtype.REALTYPE["매도수구분"][meme_gubun]

            sCode_Check = False

            if sCode in self.sell_account_stock_dict.keys():
                sCode_Check = False
                pass
            else :
                sCode_Check = True
                self.sell_account_stock_dict.update({sCode:{}})

                self.sell_account_stock_dict[sCode].update({"현재가": current_price})
                self.sell_account_stock_dict[sCode].update({"종목코드": sCode})
                self.sell_account_stock_dict[sCode].update({"종목명": stock_name})
                self.sell_account_stock_dict[sCode].update({"보유수량": stoc_quan})
                self.sell_account_stock_dict[sCode].update({"주문가능수량": like_quan})
                self.sell_account_stock_dict[sCode].update({"매입단가": buy_price})
                self.sell_account_stock_dict[sCode].update({"총매입가": total_buy_price})
                #self.sell_account_stock_dict[sCode].update({"매도매수구분": meme_gubun})

                #self.log.logPrint("잔고 매도수구분: {}".format(meme_gubun))

            if stoc_quan == 0:  # 매도 체결 끝났을때 
                self.log.logPrint("보유수량0처리 종목코드: {}, 종목명: {}".format(sCode,stock_name))
                self.Send_Sell_Sucess_Mail()
                del self.sell_account_stock_dict[sCode]
                self.dynamicCall("SetRealRemove(QString, QString)", self.screen_my_info, sCode)     # 실시간 정보 끊기 해당 종목 스크린에서
            elif sCode_Check and sCode not in self.sell_success_stock_dict_finish:
                self.Send_Sell_Order()

    def Send_Sell_Sucess_Mail(self):
        #self.detail_account_info(self.screen_my_info) #예수금 정보 가져오기
        account_num = "사용계좌: {}".format(self.account_num) + "\n\n"
        total_money = "기존 예수금: {}".format(self.use_money_origin) + "\n\n"
        
        msg = "내역 : \n"
        msg += "**********************************" + "\n"
        for key in self.sell_success_stock_dict:
            msg += "종목코드: {}".format(self.sell_success_stock_dict[key]["종목코드"]) + "\n"
            msg += "종목명: {}".format(self.sell_success_stock_dict[key]["종목명"]) + "\n"
            msg += "주문상태: {}".format(self.sell_success_stock_dict[key]["주문상태"]) + "\n"
            msg += "체결누계금액: {}".format(self.sell_success_stock_dict[key]["체결누계금액"]) + "\n"
            msg += "매도수구분: {}".format(self.sell_success_stock_dict[key]["매도수구분"]) + "\n"
            msg += "**********************************" + "\n" 

        subject = "Kiwoom 자동주식매매 Sell_Success"
        sendmsg = account_num + total_money + msg

        self.log.logPrint("Send_Sell_Success_Mail()")
        self.log.logPrint(sendmsg)

        self.objMail.SendMailMsgSet(subject, sendmsg)
    
    def hogaUnitCalc(self, price):
        hogaUnit = 1

        if price < 1000:
            hogaUnit = 10   # origin 1      주문단가 잘못입력 에러 방지
        elif price < 5000:
            hogaUnit = 10   # origin 5      주문단가 잘못입력 에러 방지
        elif price < 10000:
            hogaUnit = 10
        elif price < 50000:
            hogaUnit = 50
        elif price < 100000:
            hogaUnit = 100
        
        return hogaUnit

    def receiveMsg(self, screenNo, requestName, trCode, msg):
        """
        수신 메시지 이벤트
        서버로 어떤 요청을 했을 때(로그인, 주문, 조회 등), 그 요청에 대한 처리내용을 전달해준다.
        :param screenNo: string - 화면번호(4자리, 사용자 정의, 서버에 조회나 주문을 요청할 때 이 요청을 구별하기 위한 키값)
        :param requestName: string - TR 요청명(사용자 정의)
        :param trCode: string
        :param msg: string - 서버로 부터의 메시지
        """

        msg = "receiveMsg() - " + requestName + ": " + msg + "\n"
        self.log.logPrint(msg)
    
    def merge_sell_account(self):
        
        if not self.accout_stock_dict.keys():
            return

        for acc_code in self.accout_stock_dict:
            code = acc_code
            stock_name = self.accout_stock_dict[acc_code]["종목명"]
            current_price = self.accout_stock_dict[acc_code]["현재가"]
            current_price = abs(int(current_price))
            stoc_quan = self.accout_stock_dict[acc_code]["매매가능수량"]
            buy_price = self.accout_stock_dict[acc_code]["매입가"]

            if code in self.sell_account_stock_dict.keys():
                continue
            else:
                self.sell_account_stock_dict[code] = {}

                self.sell_account_stock_dict[code].update({"현재가": current_price})
                self.sell_account_stock_dict[code].update({"종목코드": code})
                self.sell_account_stock_dict[code].update({"종목명": stock_name})
                self.sell_account_stock_dict[code].update({"보유수량": stoc_quan})
                self.sell_account_stock_dict[code].update({"매입단가": buy_price})
