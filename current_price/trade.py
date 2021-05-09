# current_price.py
import win32com.client
import sqlite3
import ctypes


class CurrentPrice:
    def __init__(self):
        # 연결 여부 체크
        self.dicflag1 = {
                            ord(' '): '현금', ord('Y'): '융자', ord('D'): '대주',
                            ord('B'): '담보', ord('M'): '매입담보', ord('P'): '플러스론',
                            ord('I'): '자기융자',
                        }

        self.objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        self.objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
        self.objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')

        bConnect = self.objCpCybos.IsConnect
        if bConnect == 0:
            print("PLUS가 정상적으로 연결되지 않음. ")
            exit()
        conn = sqlite3.connect("../stock_kind.db", isolation_level=None)  # sqlite 연결
        self.c = conn.cursor()

    # PLUS 실행 기본 체크 함수
    def InitPlusCheck(self):
        # 프로세스가 관리자 권한으로 실행 여부
        if ctypes.windll.shell32.IsUserAnAdmin():
            print('정상: 관리자권한으로 실행된 프로세스입니다.')
        else:
            print('오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요')
            return False

        # 연결 여부 체크
        if self.objCpCybos.IsConnect == 0:
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False

        # 주문 관련 초기화
        if self.objCpTrade.TradeInit(0) != 0:
            print("주문 초기화 실패")
            return False
        return True

    # 현재가 구하기
    def get_current_price(self, name):
        self.InitPlusCheck()

        result = self.c.execute('SELECT CODE FROM STOCK_KIND WHERE name="' + name + '"')
        fetch_ = result.fetchall()
        code = ''
        if len(fetch_) == 0:
            return None
        else:
            code = fetch_[0][0]

        objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
        objStockMst.SetInputValue(0, code)  # 종목 코드 - 삼성전자
        objStockMst.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = objStockMst.GetDibStatus()
        rqRet = objStockMst.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            exit()

        return objStockMst.GetHeaderValue(11)

    def trade(self, buy, name, quantity):
        # 주문 초기화
        self.InitPlusCheck()

        result = self.c.execute('SELECT CODE FROM STOCK_KIND WHERE name="' + name + '"')
        fetch_ = result.fetchall()
        code = ''
        if len(fetch_) == 0:
            return None
        else:
            code = fetch_[0][0]

        price = self.get_current_price(name)

        # 주식 매수 주문
        acc = self.objCpTrade.AccountNumber[0]  # 계좌번호
        accFlag = self.objCpTrade.GoodsList(acc, 1)  # 주식상품 구분
        objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
        objStockOrder.SetInputValue(0, buy)  # 2: 매수, 1: 매도
        objStockOrder.SetInputValue(1, acc)  # 계좌번호
        objStockOrder.SetInputValue(2, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        objStockOrder.SetInputValue(3, code)  # 종목코드 - A003540 - 대신증권 종목
        objStockOrder.SetInputValue(4, quantity)  # 매수수량 10주
        objStockOrder.SetInputValue(5, price)  # 주문단가  - 14,100원
        objStockOrder.SetInputValue(7, "0")  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
        objStockOrder.SetInputValue(8, "01")  # 주문호가 구분코드 - 01: 보통

        # 매수 주문 요청
        objStockOrder.BlockRequest()

        rqStatus = objStockOrder.GetDibStatus()
        rqRet = objStockOrder.GetDibMsg1()
        
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return rqStatus
        return 0

    def get_portfolio(self):
        self.InitPlusCheck()

        acc = self.objCpTrade.AccountNumber[0]  # 계좌번호
        accFlag = self.objCpTrade.GoodsList(acc, 1)  # 주식상품 구분

        self.objRq = win32com.client.Dispatch("CpTrade.CpTd6033")
        self.objRq.SetInputValue(0, acc)  # 계좌번호
        self.objRq.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objRq.SetInputValue(2, 50)  # 요청 건수(최대 50)
        self.objRq.SetInputValue(3, "1")    # 수익률

        result = []

        while True:
            self.objRq.BlockRequest()
            # 통신 및 통신 에러 처리
            rqStatus = self.objRq.GetDibStatus()
            rqRet = self.objRq.GetDibMsg1()
            print("통신상태", rqStatus, rqRet)
            if rqStatus != 0:
                return False

            if len(result) == 0:        # +2 예수금
                result.append(self.objRq.GetHeaderValue(9))

            cnt = self.objRq.GetHeaderValue(7)

            for i in range(cnt):
                item = {}
                code = self.objRq.GetDataValue(12, i)  # 종목코드
                item['code'] = code
                item['name'] = self.objRq.GetDataValue(0, i)  # 종목명
                item['quantity'] = self.objRq.GetDataValue(7, i)  # 체결잔고수량
                item['evaluation'] = self.objRq.GetDataValue(9, i)  # 평가금액(천원미만은 절사 됨)
                item['profit'] = self.objRq.GetDataValue(11, i)  # 평가손익(천원미만은 절사 됨)
                item['cprice'] = self.get_current_price(item['name'])   # 현재 단가

                result.append(item)

            if not self.objRq.Continue:
                break
        return result


if __name__ == "__main__":
    print(CurrentPrice().test()[0])