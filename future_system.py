######################### API로 종목 분석하기 158-191
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtTest import *
import re
import datetime
import pandas as pd

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        print("kiwoom() class start.")

        #### event loop를 실행하기 위한 변수 모음
        self.login_event_loop = QEventLoop()
        self.detail_account_info_event_loop = QEventLoop()
        self.detail_account_info_event_loop2 = QEventLoop()
        self.detail_account_info_event_loop_f = QEventLoop()
       # self.detail_account_info_event_loop3 = QEventLoop()
       # self.calculator_event_loop = QEventLoop()

        ### 계좌관련 변수
        self.account_stock_dict = {}
        self.not_account_stock_dict = {}
        self.account_num = None
        self.deposit = 0
        self.use_money = 0
        self.use_money_percent = 0.5
        self.output_deposit = 0
        self.total_profit_loss_money = 0
        self.total_profit_loss_rate = 0.0
        self.t_avr1 = 1
        self.t_avr2 = 3
        self.t_avr3 = 5
        #종목정보 가져오기
        self.portfolio_stock_dict = {}

        ### 종목분석용
        self.calcul_data = []

        #### 요청 스크린 번호
        self.screen_my_info = "2000"
        self.screen_calculation_stock = "4000"

        ### 초기 셋팅 함수들 바로 실행
        self.get_ocx_instance()  # 파이썬으로 api 사용정의
        self.event_slots()
        self.signal_login_commConnect()     # 로그인하기
        self.get_account_info()             # 계좌번호 가져오기
        #self.stock_list()
        self.future_list()

        self.detail_account_info()          # 선옵증거금상세내역요청
        #self.detail_account_mystock()       # 계좌평가잔고내역요청
        #self.not_concluded_account()
        #self.not_concluded_account()        # 실시간미체결요청
 #       self.calculator_fnc()

        #선물 코드
        self.f_code = '101QC000'
        self.get_future_chart()


    def stock_list(self):
        self.code_0 = self.dynamicCall("GetCodeListByMarket(QString)", "0")

        now = datetime.datetime.now()
        nowdate = now.strftime('%Y%m%d')  # 오늘 날짜를 기준으로 데이터를 받아옴
        tmp_code = self.code_0.split(';')
        for obj in tmp_code:
            #print(type(obj))
            name = self.dynamicCall("GetMasterCodeName(QString)", obj)
            print('code :', obj, 'name :', name)

        #print(nowdate,'    ', tmp_code)

    def future_list(self):
        f_code = '101QC000'
        f_name = self.dynamicCall("GetMasterCodeName(QString)", f_code)
        print('code :', f_code, 'name :', f_name)

        #self.code_10 = self.dynamicCall("GetCodeListByMarket(QString)", "10")
        #print(self.code_10)
        #for obj in self.code_0:
        #    print(obj)


    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")
        # print("get_ocx")

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")
        print("hi")
        self.login_event_loop.exec_()
        # print('hi')

    def login_slot(self, err_code):
        print(errors(err_code)[1])
        self.login_event_loop.exit()


    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(QString)", "ACCNO")  # 계좌번호
        tmp1 = account_list.rstrip(';')
        #print(tmp1)
        account_num = re.split(';', tmp1)[1]
        self.account_num = account_num
        print("계좌번호 : %s" % account_num)

    def average_price_ext(self, t_df, t_date):
        if t_date == 1:
            return t_df
        else:
            r_data = pd.DataFrame(index=list(range(0, len(t_df))), columns=list('a'))
            for i in range(len(t_df)-t_date+1):
                r_data['a'][i] = round(t_df[i:i + t_date].mean(),2)
                #print(r_data['a'][i])

            return r_data['a']

    def get_future_chart(self, sPrevNext=0):
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", self.f_code)
        self.dynamicCall("SetInputValue(QString, QString)", "시간단위", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "선물옵션분차트요청", "OPT50029", sPrevNext, self.screen_my_info)
        self.detail_account_info_event_loop_f.exec_()


    def detail_account_mystock(self, sPrevNext=0):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "계좌평가잔고내역요청", "opw00018", sPrevNext, self.screen_my_info)
        self.detail_account_info_event_loop.exec_()

    def not_concluded_account(self, sPrevNext=0):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "실시간미체결요청", "opt10075", sPrevNext, self.screen_my_info)
        self.detail_account_info_event_loop3.exec_()


#    def detail_account_info(self, sPrevNext=0):
#        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
#        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "")
#        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
#        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
#        self.dynamicCall("CommRqData(QString, QString, int, QString)", "예수금상세현황요청", "opw00001", sPrevNext,self.screen_my_info)
#        self.detail_account_info_event_loop2.exec_()

    def detail_account_info(self, sPrevNext=0):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        #self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "선옵증거금상세내역요청", "opw20012", sPrevNext,self.screen_my_info)
        self.detail_account_info_event_loop2.exec_()

    def realdata_slot(self, sCode, sRealType, sRealData):
        if sRealType == "장시작시간":
            fid = self.realTyle.REALTYPE[sRealType]['장운영구분'] #(0:장시작전,장종료전(20분),3:장시작,4,8:장종료(30분),9:장마감
            value = self.dynamicCall("GetCommRealData(QString,int)",sCode,fid)
            print(value)

        elif sRealType == "주식체결":
            a = self.dynamicCall("GetCommRealData(QString,int)", sCode, self.realType.REALTYPE[sRealType]['체결시간'])
            b = self.dynamicCall("GetCommRealData(QString,int)", sCode, self.realType.REALTYPE[sRealType]['현재가'])
            b = abs(int(b))

            c = self.dynamicCall("GetCommRealData(QString,int)", sCode, self.realType.REALTYPE[sRealType]['전일대비'])
            c = abs(int(c))

            d = self.dynamicCall("GetCommRealData(QString,int)", sCode, self.realType.REALTYPE[sRealType]['등락률'])
            d = abs(float(d))

            e = self.dynamicCall("GetCommRealData(QString,int)", sCode, self.realType.REALTYPE[sRealType]['(최우선)매도호가'])
            e = abs(int(e))

            f = self.dynamicCall("GetCommRealData(QString,int)", sCode, self.realType.REALTYPE[sRealType]['(최우선)매도호가'])
            f = abs(int(f))

            if sCode not in self.portfolio_stock_dict:
                self.portfolio_stock_dict.updata({sCode:{}})

            self.portfolio_stock_dict[sCode].update({"체결시간" : a })


            # opw00001_req
    def trdata_slot(self,sScrNo,sRQName,sTrCode,sRecordName,sPrevNext):
        print("#######" + sRQName)
#        if sRQName == "예수금상세현황요청":
#            deposit = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "예수금")
#            self.deposit = int(deposit)
#            print("예수금 : %s" % self.deposit)

        if sRQName == "선옵증거금상세내역요청":
            deposit = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "인출가능총액")
            self.deposit = int(deposit)
            print("인출가능총액 : %s" % self.deposit)

            use_money = float(self.deposit)*self.use_money_percent
            self.use_money = int(use_money)
            self.use_money = self.use_money / 4

            output_deposit = self.dynamicCall("GetCommData(QString,QString,int,QString)", sTrCode, sRQName, 0, "주문가능총액")
            self.output_deposit = int(output_deposit)
            print("주문가능총액: %s" % self.output_deposit)
            self.stop_screen_cancel(self.screen_my_info)
            self.detail_account_info_event_loop2.exit()

        elif sRQName == "선물옵션분차트요청":
            #out_list = ['현재가','거래량','체결시간','시가','고가','저가','수정주가이벤트','전일종가']
            out_list = ['체결시간', '시가', '고가', '저가','현재가','거래량']
            f_result = []
            for i in range(900):
                tmp = []
                j = 0
                for obj in out_list:
                    #print(self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, obj))
                    t_data = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, obj)
                    t_data = t_data.replace(' ','',30)
                    t_data = t_data.replace('+','',10)
                    t_data = t_data.replace('-','',10)
                    if j > 1 and j < 5:
                        tmp.append(round(float(t_data),3))
                    else:
                        tmp.append(t_data)
                    j = j + 1
                f_result.append(tmp)

            df = pd.DataFrame(columns=['일자', '시가', '고가', '저가', '종가','거래량'], index=range(len(f_result)))
            t_num = 0
            for obj in f_result:
                df.loc[t_num] = obj
                t_num = t_num + 1

            df[str(self.t_avr1)] = self.average_price_ext(df['종가'], self.t_avr1)
            df[str(self.t_avr2)] = self.average_price_ext(df['종가'], self.t_avr2)
            df[str(self.t_avr3)] = self.average_price_ext(df['종가'], self.t_avr3)
            print(df)
            #for obj in f_result:
            #    print(obj)
            self.detail_account_info_event_loop_f.exit()

        elif sRQName == "계좌평가잔고내역요청":
            total_buy_money = self.dynamicCall("GetCommData(QString,QString,int,QString)", sTrCode,sRQName,0,"총매입금액")
            self.total_buy_money = int(total_buy_money)
            total_profit_loss_money = self.dynamicCall("GetCommData(QString,QString,int,QString)", sTrCode, sRQName, 0, "총평가손익금액")
            self.total_profit_loss_money = int(total_profit_loss_money)
            total_profit_loss_rate = self.dynamicCall("GetCommData(QString,QString,int,QString)", sTrCode, sRQName, 0,
                                               "총수익률(%)")
            self.total_profit_loss_rate = float(total_profit_loss_rate)
            print("계좌평가잔고내역요청 싱글데이터 : %s -  %s - %s" % (total_buy_money, total_profit_loss_money, total_profit_loss_rate))

            rows = self.dynamicCall("GetRepeatCnt(QString,QString)", sTrCode, sRQName)
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString,QString,int,QString",sTrCode,sRQName,i,"종목번호")
                code = code.strip()[1:0]
                code_nm             = self.dynamicCall("GetCommData(QString,QString,int,QString", sTrCode, sRQName, i, "종목명")
                stock_quantity      = self.dynamicCall("GetCommData(QString,QString,int,QString", sTrCode, sRQName, i, "보유수량")
                buy_price           = self.dynamicCall("GetCommData(QString,QString,int,QString", sTrCode, sRQName, i, "매입가")
                learn_rate          = self.dynamicCall("GetCommData(QString,QString,int,QString", sTrCode, sRQName, i, "수익률(%)")
                current_price       = self.dynamicCall("GetCommData(QString,QString,int,QString", sTrCode, sRQName, i, "현재가")
                total_chegul_price  = self.dynamicCall("GetCommData(QString,QString,int,QString", sTrCode, sRQName, i, "매입금액")
                possible_quantity   = self.dynamicCall("GetCommData(QString,QString,int,QString", sTrCode, sRQName, i, "매매가능수량")
                print("종목번호: %s - 종목명 : %s - 보유수량 : %s - 매입가 : %s - 수입률 : %s - 현재가 : %s" % (code, code_nm, stock_quantity, buy_price, learn_rate, current_price))

                if code in self.account_stock_dict:
                    pass
                else:
                    self.account_stock_dict[code] = {}

                code_nm = code_nm.strip()
                stock_quantity = int(stock_quantity.strip())
                buy_price = int(buy_price.strip())
                learn_rate = float(learn_rate.strip())
                current_price  = int(current_price.strip())
                total_chegul_price = int(total_chegul_price.strip())
                possible_quantity = int(possible_quantity.strip())

                self.account_stock_dict[code].update({"종목명" : code_nm})
                self.account_stock_dict[code].update({"보유수량" : stock_quantity})
                self.account_stock_dict[code].update({"매입가" : buy_price})
                self.account_stock_dict[code].update({"수익률(%)" : learn_rate})
                self.account_stock_dict[code].update({"현재가" : current_price})
                self.account_stock_dict[code].update({"매입금액" : total_chegul_price})
                self.account_stock_dict[code].update({"매매가능수량" : possible_quantity})


            print("sPreNect : %s" % sPrevNext)
            print("계좌에 가지고 있는 종목은 %s" % rows)

            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
            #self.stop_screen_cancel(self.screen_my_info)
                print("close")
                self.detail_account_info_event_loop.exit()

        elif sRQName == "실시간미체결요청":
            rows = self.dynamicCall("GetRepeatCnt(QString,QString", sTrCode, sRQName)
            for i in range(rows):
                code            = self.dynamicCall("GetCommData(QString,QString,int,QString",sTrCode,sRQName,i,"종목코드")
                code_nm         = self.dynamicCall("GetCommData(QString,QString,int,QString", sTrCode, sRQName, i, "종목명")
                order_no        = self.dynamicCall("GetCommData(QString,QString,int,QString", sTrCode, sRQName, i, "주문번호")
                order_status    = self.dynamicCall("GetCommData(QString,QString,int,QString", sTrCode, sRQName, i, "주문상태")
                order_quantity  = self.dynamicCall("GetCommData(QString,QString,int,QString", sTrCode, sRQName, i, "주문수량")
                order_price     = self.dynamicCall("GetCommData(QString,QString,int,QString", sTrCode, sRQName, i, "주문가격")
                order_gubun     = self.dynamicCall("GetCommData(QString,QString,int,QString", sTrCode, sRQName, i, "주문구분")
                not_quantity    = self.dynamicCall("GetCommData(QString,QString,int,QString", sTrCode, sRQName, i, "미체결수량")
                ok_quantity     = self.dynamicCall("GetCommData(QString,QString,int,QString", sTrCode, sRQName, i, "체결량")


                code            = code.strip()
                code_nm         = code_nm.strip()
                order_no        = int(order_no.strip())
                order_status    = order_status.strip()
                order_quantity  = int(order_quantity.strip())
                order_price     = int(order_price.strip())
                order_gubun     = order_gubun.strip().lstrip('+').lstrip('-')
                not_quantity    = int(not_quantity.strip())
                ok_quantity     = int(ok_quantity.strip())

                if order_no in self.not_account_stock_dict:
                    pass
                else:
                    self.not_account_stock_dict[order_no] = {}

                    self.not_account_stock_dict[order_no].update({'종목코드' : code})
                    self.not_account_stock_dict[order_no].update({'종목명': code_nm})
                    self.not_account_stock_dict[order_no].update({'주문번호' : order_no})
                    self.not_account_stock_dict[order_no].update({'주문상태' : order_status})
                    self.not_account_stock_dict[order_no].update({'주문수량': order_quantity})
                    self.not_account_stock_dict[order_no].update({'주문가격': order_price})
                    self.not_account_stock_dict[order_no].update({'주문구분': order_gubun})
                    self.not_account_stock_dict[order_no].update({'미체결수량':not_quantity})
                    self.not_account_stock_dict[order_no].update({'체결량':ok_quantity})

                    print("미체결종목 : %s" % self.not_account_stock_dict[order_no])
            print("bye")
            self.detail_account_info_event_loop3.exit()

        elif sRQName == "주식일봉차트조회":
            code = self.dynamicCall("GetCommData(QString,QString,int,QString)", sTrCode, sRQName, 0, "종목코드")
            code = code.strip()
            print(code)
            #data = self.dynamicCall("GetCommData(QString,QString)", sTrCode, sRQName)
            cnt = self.dynamicCall("GetRepeatCnt(QString,QString)", sTrCode, sRQName)
            print("남은 일자 수 %s" % cnt)

            for i in range(cnt):
                data = []
                current_price   = self.dynamicCall("GetCommData(QString,QString,QString,QString)",sTrCode, sRQName, i,"현재가")
                value           = self.dynamicCall("GetCommData(QString,QString,QString,QString)",sTrCode, sRQName, i,"거래량")
                trading_value   = self.dynamicCall("GetCommData(QString,QString,QString,QString)",sTrCode, sRQName, i,"거래대금")
                date            = self.dynamicCall("GetCommData(QString,QString,QString,QString)",sTrCode, sRQName, i,"일자")
                start_price     = self.dynamicCall("GetCommData(QString,QString,QString,QString)",sTrCode, sRQName, i,"시가")
                high_price      = self.dynamicCall("GetCommData(QString,QString,QString,QString)",sTrCode, sRQName, i,"고가")
                low_price       = self.dynamicCall("GetCommData(QString,QString,QString,QString)",sTrCode, sRQName, i,"저가")

                data.append("")
                data.append(current_price.strip())
                data.append(value.strip())
                data.append(trading_value.strip())
                data.append(date.strip())
                data.append(start_price.strip())
                data.append(high_price.strip())
                data.append(low_price.strip())
                data.append("")

                self.calcul_data.append(data.copy())


            if sPrevNext == "2":
                self.day_kiwoom_db(code=code, sPrevNext=sPrevNext)
            else:
                print("총 일수 %s" % len(self.calcul_data))
                pass_success = False

                if self.calcul_data == None or len(self.calcul_data) < 120:
                    pass_success = False
                else:
                    # 120일 이평선의 최근 가격 구함
                    total_price = 0
                    for value in self.calcul_data[:120]:
                        total_price += int(value[1])
                    moving_average_price = total_price / 120

                    #오늘자 주가가 120일 이편선에 걸쳐있는지 확인
                    bottom_stock_price = False
                    check_price = None
                    if int(self.calcul_data[0][7]) <= moving_average_price and moving_average_price <= int(self.calcul_data[0][6]):
                        print("오늘의 주가가 120 이평선에 걸쳐있는지 확인")
                        bottom_stock_price = True
                        check_price = int(self.calcul_data[0][6])

                    #밑에 있으면 주가가 120일 이평선보다 위에 있던 구간이 있었는지 확인
                    prev_price = None
                    if bottom_stock_price == True:
                        moving_average_price_prev = 0
                        price_top_moving = False
                        idx = 1

                        while True:
                            if len(self.calcul_data[idx:]) < 120:
                                print("120일 치가 없음")
                                break

                            total_price = 0
                            for value in self.calcul_data[idx:120+idx]:
                                total_price += int(value[1])
                            moving_average_price_prev = total_price /120

                            if moving_average_price_prev <= int(self.calcul_data[idx][6]) and idx < 20:
                                print("20일 동안 주가가 120일 이평선과 같거나 위에 있으면 조건 통과 못함")
                                price_top_moving = False
                                break
                            elif int(self.calcul_data[idx[7]]) <= moving_average_price_prev and idx > 20:
                                print("120일 이평선 위에 있는 구간 확인됨")
                                price_top_moving = True
                                prev_price = int(self.calcul_data[idx[7]])
                                break

                            idx += 1


                        if price_top_moving == True:
                            if moving_average_price > moving_average_price_prev and check_price > prev_price:
                                print("포착된 이평선의 가격이 오늘자 이평선 가격보다 낮은것 확인")
                                print("포착된 부분의 일봉 저가가 오늘자 일봉의 고가보다 낮은지 확인")
                                pass_success = True


                if pass_success == True:
                    print("조건부 통과됨")

                    code_nm = self.dynamicCall("GetMasterCodeName(QString", code)
                    f = open("D://David\Documents//01_project//25_주식파이썬//06_120일선결과//r.txt")
                    f.write("%s\t%s\t%s\n" % (code,code_nm,str(self.calcul_data[0][1])))
                    f.close()
                elif pass_success == False:
                    print("조건부 통과 못 함")

                self.calcul_data.clear()
                self.calculator_event_loop.exit()

    def stop_screen_cancel(self,sScrNo=None):
        self.dynamicCall("DisconnectRealData(QString)", sScrNo)

    def get_code_list_by_market(self, market_code):
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market_code)
        code_list = code_list.split(';')[:-1]
        return code_list

    ## 0 : 장내 // 10 : 코스닥 // 3 : ELW // 8 : ETF // 50 : KONEX // 4 : 뮤추얼펀드 // 5 : 신주인수권 // 6 : 리츠 // 9 : 하이얼펀드 // 30 : K=OTC
    def calculator_fnc(self):
        code_list = self.get_code_list_by_market("10")
        print("갯수 %s" % len(code_list))

        for idx, code in enumerate(code_list[:10]):
            self.dynamicCall("DisconnectRealData(QString)", self.screen_calculation_stock)
            #스크린 연결 끊기

            print("%s / %s : KOSDAQ Stock Code : %s is updating... " % (idx+1, len(code_list), code))
            self.day_kiwoom_db(code=code)

    def day_kiwoom_db(self, code=None, date=None, sPrevNext = 0):
        QTest.qWait(3600)
        self.dynamicCall("SetInputValue(QString,QString)", "종목코드", code)
        self.dynamicCall("SetInputValue(QString,QString)", "수정주가구분", "1")

        if date != None:
            self.dynamicCall("SetInputValue(QString,QString)", "기준일자", date)
        self.dynamicCall("CommRqData(QString,QString,int,QString)", "주식일봉차트조회","opt10081",sPrevNext,self.screen_calculation_stock)

        self.calculator_event_loop.exec_()



if __name__ == "__main__":
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()  # MyWindow 클래스를 생성하여 myWondow 변수에 할당
    app.exec_()
    # kiwoom.show()

    # myWindow.show()  # MyWindow 클래스를 노출
    # app.exec_()
