from qtkiwoom import *


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setGeometry(300, 300, 300, 300)

        # callback 등록
        callbacks = {
            "login": self.callback_login,
            "tr_data": self.callback_tr_data,
        }

        self.qtkiwoom = QtKiwoom(callbacks)
        self.qtkiwoom.CommConnect()

    def callback_login(self, *args, **kwargs):
        if kwargs['err_code'] == 0:
            self.statusBar().showMessage("로그인 완료")
            #self.get_login_info()
            self.request_tr()

    def callback_tr_data(self, *args, **kwargs):
        rqname = kwargs['rqname']
        data = kwargs['data']
        print(rqname)
        print(data)

    def get_login_info(self):
        accounts = self.qtkiwoom.GetLoginInfo("ACCNO")
        print(accounts)

    def request_tr(self):
        self.qtkiwoom.SetInputValue("종목코드", "005930")
        self.qtkiwoom.CommRqData("opt10001_삼성전자", "opt10001", 0, "1001")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()