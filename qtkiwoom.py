import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
import pandas as pd


class QtKiwoom:
    def __init__(self, callbacks=None):
        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self._set_callbacks(callbacks)
        self._set_signal_slots()

    def _set_callbacks(self, callbacks):
        if callbacks is None:
            default_callbacks = {
                "login": None,
                "tr_data": None
            }
            self.callbacks = default_callbacks
        else:
            self.callbacks = callbacks

    def _set_signal_slots(self):
        self.ocx.OnReceiveTrData.connect(self._handler_tr_data)
        self.ocx.OnReceiveRealData.connect(self._handler_real_data)
        self.ocx.OnReceiveMsg.connect(self._handler_msg)
        self.ocx.OnReceiveChejanData.connect(self._handler_chejan_data)
        self.ocx.OnEventConnect.connect(self._handler_login)
        self.ocx.OnReceiveRealCondition.connect(self._handler_real_condition)
        self.ocx.OnReceiveTrCondition.connect(self._handler_tr_condition)
        self.ocx.OnReceiveConditionVer.connect(self._handler_condition_ver)

    # -------------------------------------------------------------------------------------------------------------------
    # OpenAPI+ 메서드
    # -------------------------------------------------------------------------------------------------------------------
    def CommConnect(self):
        """
        1) CommConnect
        로그인 윈도우를 실행합니다.
        :param block: True: 로그인완료까지 블록킹 됨, False: 블록킹 하지 않음
        :return: None
        """
        self.ocx.dynamicCall("CommConnect()")

    def CommRqData(self, rqname, trcode, next, screen):
        """
        3) CommRqData
        TR을 서버로 송신합니다.
        :param rqname: 사용자가 임의로 지정할 수 있는 요청 이름
        :param trcode: 요청하는 TR의 코드
        :param next: 0: 처음 조회, 2: 연속 조회
        :param screen: 화면번호 ('0000' 또는 '0' 제외한 숫자값으로 200개로 한정된 값
        :return: None
        """
        self.ocx.dynamicCall("CommRqData(QString, QString, int, QString)", rqname, trcode, next, screen)

    def GetLoginInfo(self, tag):
        """
        4) GEtLoginInfo
        로그인한 사용자 정보를 반환하는 메서드
        :param tag: ("ACCOUNT_CNT, "ACCNO", "USER_ID", "USER_NAME", "KEY_BSECGB", "FIREW_SECGB")
        :return: tag에 대한 데이터 값
        """
        data = self.ocx.dynamicCall("GetLoginInfo(QString)", tag)

        if tag == "ACCNO":
            return data.split(';')[:-1]
        else:
            return data

    def SendOrder(self, rqname, screen, accno, order_type, code, quantity, price, hoga, order_no):
        """
        5) SendOrder
        주식 주문을 서버로 전송하는 메서드
        시장가 주문시 주문단가는 0으로 입력해야 함 (가격을 입력하지 않음을 의미)
        :param rqname: 사용자가 임의로 지정할 수 있는 요청 이름
        :param screen: 화면번호 ('0000' 또는 '0' 제외한 숫자값으로 200개로 한정된 값
        :param accno: 계좌번호 10자리
        :param order_type: 1: 신규매수, 2: 신규매도, 3: 매수취소, 4: 매도취소, 5: 매수정정, 6: 매도정정
        :param code: 종목코드
        :param quantity: 주문수량
        :param price: 주문단가
        :param hoga: 00: 지정가, 03: 시장가,
                     05: 조건부지정가, 06: 최유리지정가, 07: 최우선지정가,
                     10: 지정가IOC, 13: 시장가IOC, 16: 최유리IOC,
                     20: 지정가FOK, 23: 시장가FOK, 26: 최유리FOK,
                     61: 장전시간외종가, 62: 시간외단일가, 81: 장후시간외종가
        :param order_no: 원주문번호로 신규 주문시 공백, 정정이나 취소 주문시에는 원주문번호를 입력
        :return:
        """
        ret = self.ocx.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                   [rqname, screen, accno, order_type, code, quantity, price, hoga, order_no])
        return ret

    def SendOrderCredit(self, rqname, screen, accno, order_type, code, quantity, price, hoga, credit, loan_date, order_no):
        """
        6) SendOrderCredit
        주식 주문을 서버로 전송하는 메서드
        시장가 주문시 주문단가는 0으로 입력해야 함 (가격을 입력하지 않음을 의미)
        :param rqname: 사용자가 임의로 지정할 수 있는 요청 이름
        :param screen: 화면번호 ('0000' 또는 '0' 제외한 숫자값으로 200개로 한정된 값
        :param accno: 계좌번호 10자리
        :param order_type: 1: 신규매수, 2: 신규매도, 3: 매수취소, 4: 매도취소, 5: 매수정정, 6: 매도정정
        :param code: 종목코드
        :param quantity: 주문수량
        :param price: 주문단가
        :param hoga: 00: 지정가, 03: 시장가,
                     05: 조건부지정가, 06: 최유리지정가, 07: 최우선지정가,
                     10: 지정가IOC, 13: 시장가IOC, 16: 최유리IOC,
                     20: 지정가FOK, 23: 시장가FOK, 26: 최유리FOK,
                     61: 장전시간외종가, 62: 시간외단일가, 81: 장후시간외종가
        :param credit: 신용구분
        :param load_date: 대출일
        :param order_no: 원주문번호로 신규 주문시 공백, 정정이나 취소 주문시에는 원주문번호를 입력
        :return:
        """
        ret = self.ocx.dynamicCall("SendOrderCredit(QString, QString, QString, int, QString, int, int, QString, QString, QString, QString)",
                                   [rqname, screen, accno, order_type, code, quantity, price, hoga, credit, loan_date, order_no])
        return ret

    def SetInputValue(self, id, value):
        """
        7) SetInputValue
        TR 입력값을 설정하는 메서드
        :param id: TR INPUT의 아이템명
        :param value: 입력 값
        :return: None
        """
        self.ocx.dynamicCall("SetInputValue(QString, QString)", id, value)

    def DisconnectRealData(self, screen):
        """
        10) DisconnectRealData
        화면번호에 대한 리얼 데이터 요청을 해제하는 메서드
        :param screen: 화면번호
        :return: None
        """
        self.ocx.dynamicCall("DisconnectRealData(QString)", screen)

    def GetRepeatCnt(self, trcode, rqname):
        """
        11) GetRepeatCnt
        멀티데이터의 행(row)의 개수를 얻는 메서드
        :param trcode: TR코드
        :param rqname: 사용자가 설정한 요청이름
        :return: 멀티데이터의 행의 개수
        """
        count = self.ocx.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        return count

    def CommKwRqData(self, arr_code, next, code_count, type, rqname, screen):
        """
        12) CommKwRqData
        여러 종목 (한 번에 100종목)에 대한 TR을 서버로 송신하는 메서드
        :param arr_code: 여러 종목코드 예: '000020:000040'
        :param next: 0: 처음조회
        :param code_count: 종목코드의 개수
        :param type: 0: 주식종목 3: 선물종목
        :param rqname: 사용자가 설정하는 요청이름
        :param screen: 화면번호
        :return:
        """
        ret = self.ocx.dynamicCall("CommKwRqData(QString, bool, int, int, QString, QString)", arr_code, next,
                                   code_count, type, rqname, screen);
        return ret

    def GetAPIModulePath(self):
        """
        13) GetAPIModulePath
        OpenAPI 모듈의 경로를 반환하는 메서드
        :return: 모듈의 경로
        """
        ret = self.ocx.dynamicCall("GetAPIModulePath()")
        return ret

    def GetCodeListByMarket(self, market):
        """
        14) GetCodeListByMarket
        시장별 상장된 종목코드를 반환하는 메서드
        :param market: 0: 코스피, 3: ELW, 4: 뮤추얼펀드 5: 신주인수권 6: 리츠
                       8: ETF, 9: 하이일드펀드, 10: 코스닥, 30: K-OTC, 50: 코넥스(KONEX)
        :return: 종목코드 리스트 예: ["000020", "000040", ...]
        """
        data = self.ocx.dynamicCall("GetCodeListByMarket(QString)", market)
        tokens = data.split(';')[:-1]
        return tokens

    def GetConnectState(self):
        """
        15) GetConnectState
        현재접속 상태를 반환하는 메서드
        :return: 0:미연결, 1: 연결완료
        """
        ret = self.ocx.dynamicCall("GetConnectState()")
        return ret

    def GetMasterCodeName(self, code):
        """
        16) GetMasterCodeName
        종목코드에 대한 종목명을 얻는 메서드
        :param code: 종목코드
        :return: 종목명
        """
        data = self.ocx.dynamicCall("GetMasterCodeName(QString)", code)
        return data

    def GetMasterListedStockCnt(self, code):
        """
        17) GetMasterListedStockCnt
        종목에 대한 상장주식수를 리턴하는 메서드
        :param code: 종목코드
        :return: 상장주식수
        """
        data = self.ocx.dynamicCall("GetMasterListedStockCnt(QString)", code)
        return data

    def GetMasterConstruction(self, code):
        """
        18) GetMasterConstruction
        종목코드에 대한 감리구분을 리턴
        :param code: 종목코드
        :return: 감리구분 (정상, 투자주의 투자경고, 투자위험, 투자주의환기종목)
        """
        data = self.ocx.dynamicCall("GetMasterConstruction(QString)", code)
        return data

    def GetMasterListedStockDate(self, code):
        """
        19) GetMasterListedStockDate
        종목코드에 대한 상장일을 반환
        :param code: 종목코드
        :return: 상장일 예: "20100504"
        """
        data = self.ocx.dynamicCall("GetMasterListedStockDate(QString)", code)
        return datetime.datetime.strptime(data, "%Y%m%d")

    def GetMasterLastPrice(self, code):
        """
        20) GetMasterLastPrice
        종목코드의 전일가를 반환하는 메서드
        :param code: 종목코드
        :return: 전일가
        """
        data = self.ocx.dynamicCall("GetMasterLastPrice(QString)", code)
        return int(data)

    def GetMasterStockState(self, code):
        """
        21) GetMasterStockState
        종목의 종목상태를 반환하는 메서드
        :param code: 종목코드
        :return: 종목상태
        """
        data = self.ocx.dynamicCall("GetMasterStockState(QString)", code)
        return data.split("|")

    def GetDataCount(self, record):
        """
        22) GetDataCount
        :param record: 레코드명
        :return: 레코드 반복개수
        """
        count = self.ocx.dynamicCall("GetDataCount(QString)", record)
        return count

    def GetOutputValue(self, record, repeat_index, item_index):
        """
        23) GetOutputValue
        :param record:
        :param repeat_index:
        :param item_index:
        :return:
        """
        count = self.ocx.dynamicCall("GetOutputValue(QString, int, int)", record, repeat_index, item_index)
        return count

    def GetCommData(self, trcode, rqname, index, item):
        """
        24) GetCommData
        수순 데이터를 가져가는 메서드
        :param trcode: TR 코드
        :param rqname: 요청 이름
        :param index: 멀티데이터의 경우 row index
        :param item: 얻어오려는 항목 이름
        :return:
        """
        data = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, index, item)
        return data.strip()

    def GetCommRealData(self, code, fid):
        """
        25) GetCommRealData
        :param code:
        :param fid:
        :return:
        """
        data = self.ocx.dynamicCall("GetCommRealData(QString, int)", code, fid)
        return data

    def GetChejanData(self, fid):
        """
        26) GetChejanData
        :param fid:
        :return:
        """
        data = self.ocx.dynamicCall("GetChejanData(int)", fid)
        return data

    def GetThemeGroupList(self, type=1):
        """
        27) GetThemeGroupList
        :param type:
        :return:
        """
        data = self.ocx.dynamicCall("GetThemeGroupList(int)", type)
        tokens = data.split(';')
        if type == 0:
            grp = {x.split('|')[0]: x.split('|')[1] for x in tokens}
        else:
            grp = {x.split('|')[1]: x.split('|')[0] for x in tokens}
        return grp

    def GetThemeGroupCode(self, theme_code):
        """
        28) GetThemeGroupCode
        :param theme_code:
        :return:
        """
        data = self.ocx.dynamicCall("GetThemeGroupCode(QString)", theme_code)
        data = data.split(';')
        return [x[1:] for x in data]

    def GetFutureList(self):
        """
        29) GetFutureList
        :return:
        """
        data = self.ocx.dynamicCall("GetFutureList()")
        return data


    def SetRealReg(self, screen, code_list, fid_list, real_type):
        """
        49) SetRealReg
        :param screen:
        :param code_list:
        :param fid_list:
        :param real_type:
        :return:
        """
        ret = self.ocx.dynamicCall("SetRealReg(QString, QString, QString, QString)", screen, code_list, fid_list, real_type)
        return ret

    def SetRealRemove(self, screen, del_code):
        """
        50) SetRealRemove
        :param screen:
        :param del_code:
        :return:
        """
        ret = self.ocx.dynamicCall("SetRealRemove(QString, QString)", screen, del_code)
        return ret

    def GetConditionLoad(self):
        """
        51) GetConditionLoad
        :return:
        """
        self.ocx.dynamicCall("GetConditionLoad()")

    def SendCondition(self, screen, cond_name, cond_index, search):
        """
        53) SendCondition
        :param screen:
        :param cond_name:
        :param cond_index:
        :param search:
        :return:
        """
        self.ocx.dynamicCall("SendCondition(QString, QString, int, int)", screen, cond_name, cond_index, search)

    def SendConditionStop(self, screen, cond_name, index):
        """
        54) SendConditionStop
        :param screen:
        :param cond_name:
        :param index:
        :return:
        """
        self.ocx.dynamicCall("SendConditionStop(QString, QString, int)", screen, cond_name, index)

    def GetCommDataEx(self, trcode, record):
        """
        55) GetCommDataEx
        :param trcode:
        :param record:
        :return:
        """
        data = self.ocx.dynamicCall("GetCommDataEx(QString, QString)", trcode, record)
        return data

    # -------------------------------------------------------------------------------------------------------------------
    # OpenAPI+ 이벤트 핸들러
    # -------------------------------------------------------------------------------------------------------------------
    def _handler_tr_data(self, screen, rqname, trcode, record, next):
        """
        1) OnReceiveTrData
        :param screen:
        :param rqname:
        :param trcode:
        :param record:
        :param next:
        :return:
        """
        if trcode == "opt10001":
            df = self.parse_opt10001(screen, trcode, rqname)

        # callback
        # pass the DataFrame
        fn = self.callbacks['tr_data']
        if callable(fn):
            fn(rqname=rqname, data=df)


    def _handler_real_data(self, code, real_type, real_data):
        """
        2) OnReceiveRealData
        :param code:
        :param real_type:
        :param real_data:
        :return:
        """
        pass

    def _handler_msg(self, screen, rqname, trcode, msg):
        """
        3) OnReceiveMsg
        :param screen:
        :param rqname:
        :param trcode:
        :param msg:
        :return:
        """
        pass

    def _handler_chejan_data(self, gubun, item_cnt, fid_list):
        """
        4) OnReceiveChejanData
        :param gubun:
        :param item_cnt:
        :param fid_list:
        :return:
        """
        pass

    def _handler_login(self, err_code):
        """
        5) OnEventConnect
        :param err_code:
        :return:
        """
        fn = self.callbacks['login']
        if callable(fn):
            fn(err_code=err_code)

    def _handler_real_condition(self, code, type, condition_name, condition_index):
        """
        6) OnReceiveRealCondition
        :param code:
        :param type:
        :param condition_name:
        :param condition_index:
        :return:
        """
        pass

    def _handler_tr_condition(self, screen, code_list, condition_name, index, next):
        """
        7) OnReceiveTrCondition
        :param screen:
        :param code_list:
        :param condition_name:
        :param index:
        :param next:
        :return:
        """
        pass

    def _handler_condition_ver(self, ret, msg):
        """
        8) OnReceiveConditionVer
        :param ret:
        :param msg:
        :return:
        """
        pass

    # -------------------------------------------------------------------------------------------------------------------
    # TR Parser
    # -------------------------------------------------------------------------------------------------------------------
    def parse_opt10001(self, screen, trcode, rqname):
        index = 0
        종목코드 = self.GetCommData(trcode, rqname, index, "종목코드")
        종목명 = self.GetCommData(trcode, rqname, index, "종목명")

        data = {
            "종목코드": [종목코드],
            "종목명": [종목명]
        }
        df = pd.DataFrame(data)
        return df


    # -------------------------------------------------------------------------------------------------------------------
    # Real Parser
    # -------------------------------------------------------------------------------------------------------------------



