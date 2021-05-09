import ctypes
import os
import time

import win32com.client
from pywinauto import application


class AutoLogin:
    def __init__(self):
        pass

    def login(self):  # 로그인
        os.system('taskkill /IM ncStarter* /F /T')
        os.system('taskkill /IM CpStart* /F /T')
        os.system('taskkill /IM DibServer* /F /T')

        os.system('wmic process where "name like \'%ncStarter%\'" call terminate')
        os.system('wmic process where "name like \'%CpStart%\'" call terminate')
        os.system('wmic process where "name like \'%DibServer%\'" call terminate')

        time.sleep(5)
        app = application.Application()
        app.start('C:\daishin\STARTER\\ncStarter.exe /prj:cp /id:아이디 /pwd:비밀번호 /pwdcert:인증서비밀번호 /autostart')
        time.sleep(30)

        return self.check_connect()    # 연결 실패시 10초 후 재시도


    def check_connect(self):  # 연결 확인
        cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
        cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')

        if not ctypes.windll.shell32.IsUserAnAdmin():  # 관리자 권한 체크
            print('check system : admin user fail')
            return False

        if cpStatus.IsConnect == 0:  # 사이보스 연결 체크
            print('check system : connect server fail')
            return False

        if cpTradeUtil.TradeInit(0) != 0:  # 거래 초기화 체크
            print('check system : init trade fail')
            return False
        print('success connect')
        return True
