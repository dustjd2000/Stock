from Manage.Mail import SendMail
from ui.ui import *

class Main():
    def __init__(self):
        print("Start Class init")
        UI_class()
        
    
    def sendMail(self):
        print("sendMail() start")
        objMail = SendMail()

        objMail.setSubject("Python Mail send Test")
        objMail.setMsg("내용 잘 갔나?")

        objMail.sendMail()


if __name__ == '__main__':
    start = Main()
    #start.sendMail()