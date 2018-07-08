#-*-coding: utf-8 -*-

import win32com.client
import pythoncom
import sys
import pandas as pd
import datetime
import time
import os


class XASessionEvents:
    logInState = 0

    def __init__(self):
        print("XASessionEvents 객체생성")

    def OnLogin(self, code, msg):
        print("onLogin method is called\n")
        print(str(code))
        print(str(msg))
        # 0000이 입력될 때만 로그인 성공
        if str(code) == '0000':
            XASessionEvents.logInState = 1

    def OnLogout(self):
        print("OnLogout method is called\n")

    def OnDisconnect(self):
        print("OnDisconnect method is called\n")


# 데이터 조회시 사용하는 클래스
class XAQueryEvents:
    queryState = 0

    def __init__(self):
        print("XAQueryEvents 객체생성")

    def OnReceiveData(self, szTrCode):
        print("ReceiveData %s\n" % szTrCode)
        XAQueryEvents.queryState = 1

    def OnReceiveMessage(self, systemError, messageCode, message):
        print(messageCode, message)

    def GetPrice(self, szItemNumber): # 현재가 조회함수
        inXAQuery.LoadFromResFile("res\\t1102.res")  # res 등록 (주식 현재가)
        inXAQuery.SetFieldData("t1102InBlock", "shcode", 0, szItemNumber)  # 종목코드 입력
        inXAQuery.Request(0)

        while XAQueryEvents.queryState == 0:
            pythoncom.PumpWaitingMessages()

        # Get FieldData
        item_name = inXAQuery.GetFieldData("t1102OutBlock", "hname", 0)
        now_price = inXAQuery.GetFieldData("t1102OutBlock", "price", 0)

        XAQueryEvents.queryState = 0
        return int(now_price)

    def GetBalance(self, szAccountNum, szAccountPwd): # 잔고조회 함수
        inXAQuery.LoadFromResFile("res\\CSPAQ12300.res")  # res 등록 (계좌잔고조회)
        inXAQuery.SetFieldData("CSPAQ12300InBlock1", "RecCnt", 0, "00001")
        inXAQuery.SetFieldData("CSPAQ12300InBlock1", "AcntNo", 0, szAccountNum)
        inXAQuery.SetFieldData("CSPAQ12300InBlock1", "Pwd", 0, szAccountPwd)
        inXAQuery.SetFieldData("CSPAQ12300InBlock1", "BalCreTp", 0, "0")
        inXAQuery.SetFieldData("CSPAQ12300InBlock1", "CmsnAppTpCode", 0, "0")
        inXAQuery.SetFieldData("CSPAQ12300InBlock1", "D2balBaseQryTp", 0, "1")
        inXAQuery.SetFieldData("CSPAQ12300InBlock1", "UprcTpCode", 0, "0")
        inXAQuery.Request(0)

        while XAQueryEvents.queryState == 0:
            pythoncom.PumpWaitingMessages()

        # 조회 결과값 처리
        df_balance = pd.DataFrame(columns=['item_number', 'item_name', 'balance_quantity', 'average_price_paid',
                                              'now_price', 'purchased_amount'])
        count = inXAQuery.GetBlockCount("CSPAQ12300OutBlock3")
        # print(count)
        for i in range(count):
            item_number = inXAQuery.GetFieldData("CSPAQ12300OutBlock3", "IsuNo", i)
            item_name = inXAQuery.GetFieldData("CSPAQ12300OutBlock3", "IsuNm", i)
            balance_quantity = inXAQuery.GetFieldData("CSPAQ12300OutBlock3", "BnsBaseBalQty", i)
            average_price_paid = inXAQuery.GetFieldData("CSPAQ12300OutBlock3", "AvrUprc", i)
            now_price = inXAQuery.GetFieldData("CSPAQ12300OutBlock3", "NowPrc", i)
            purchased_amount = inXAQuery.GetFieldData("CSPAQ12300OutBlock3", "PchsAmt", i)
            df_balance.loc[i] = [item_number, item_name, balance_quantity, average_price_paid, now_price, purchased_amount]

        df_balance[['balance_quantity', 'average_price_paid', 'now_price', 'purchased_amount']] = df_balance[['balance_quantity', 'average_price_paid', 'now_price', 'purchased_amount']].apply(pd.to_numeric)
        XAQueryEvents.queryState = 0
        return df_balance