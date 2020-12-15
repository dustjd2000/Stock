import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText 
from email.mime.image import MIMEImage

"""
recipients = ['dustjd2000@kakao.com'] # 수신인 입력
strFrom = 'dustjd2000@gmail.com'
strTo = ", ".join(recipients)
msgRoot = MIMEMultipart('related') #그대로 작성
msgRoot['Subject'] = 'MailTest'
# 여기부터 그대로 작성
msgRoot['From'] = strFrom
msgRoot['To'] = strTo
#msgRoot.preamble = 'This is a multi-part message in MIME format.'
msgAlternative = MIMEMultipart('alternative') 
msgRoot.attach(msgAlternative)
textMsg = "테스트ㅌ트트트트트" 
msgText = MIMEText(textMsg)
msgAlternative.attach(msgText) 

s = smtplib.SMTP('smtp.gmail.com: 587')
s.starttls()
s.login('dustjd2000@gmail.com', '#qhwhehrh3031')
s.sendmail(strFrom, strTo, msgRoot.as_string())

s.quit()
"""

class SendMail(object):
    def __init__(self):
        self.loginMail = 'dustjd2000@gmail.com'
        self.loginPass = '#qhwhehrh3031'
        self.strFrom = 'dustjd2000@gmail.com'
        self.recipients = ['dustjd2000@kakao.com']
        self.textMsg = "MsgTest"
        self.subject = "SubjectTest"
    
    def addRecipients(self, recipient):
        self.recipients.append(recipient)
    
    def setFrom(self, fromMail):
        self.strFrom = fromMail
    
    def setMsg(self, textMsg):
        self.textMsg = textMsg

    def setSubject(self, subject):
        self.subject = subject

    def sendMail(self): 
        strTo = ", ".join(self.recipients)
        msgRoot = MIMEMultipart('related') #그대로 작성
        msgRoot['Subject'] = self.subject
        # 여기부터 그대로 작성
        msgRoot['From'] = self.strFrom
        msgRoot['To'] = strTo
        #msgRoot.preamble = 'This is a multi-part message in MIME format.'
        msgAlternative = MIMEMultipart('alternative') 
        msgRoot.attach(msgAlternative)
        msgText = MIMEText(self.textMsg)
        msgAlternative.attach(msgText) 

        s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        s.login(self.loginMail, self.loginPass)
        s.sendmail(self.strFrom, strTo, msgRoot.as_string())

        s.quit()

        print("Send Mail() Success!")