from Manage.Mail import SendMail
from ui.ui import *

class Main(object):
    def __init__(self):
        print("Start Class init")
        
    def sendMail(self):
        print("sendMail() start")
        objMail = SendMail()

        objMail.setSubject("Kiwoom 자동주식매매 Start")
        objMail.setMsg("자동 주식 매매 시작!")

        objMail.sendMail()

    def start_kiwoom(self):
        UI_class()

if __name__ == '__main__':
    start = Main()
    start.sendMail()
    start.start_kiwoom()