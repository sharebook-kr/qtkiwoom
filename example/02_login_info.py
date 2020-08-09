from qtkiwoom import *


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setGeometry(300, 300, 300, 300)

        # callback 메소드
        callbacks = {
            "login": self.callback_login,
            "tr_data": None,
        }

        self.qtkiwoom = QtKiwoom(callbacks)
        self.qtkiwoom.CommConnect()

    def callback_login(self, *args, **kwargs):
        if kwargs['err_code'] == 0:
            self.statusBar().showMessage("로그인 완료")
            self.get_login_info()

    def get_login_info(self):
        accounts = self.qtkiwoom.GetLoginInfo("ACCNO")
        print(accounts)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()