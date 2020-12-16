import time
import os.path

class LogManager(object):

    def __init__(self):
        filePath = os.path.dirname(os.path.realpath(os.path.dirname(os.path.realpath(__file__))))
        filePath = os.path.join(filePath,"log")
        self.fileFullPath = ""
        self.makelogfile(filePath)

    def logPrint(self, content):
        print(content)

        now = time.localtime()

        hour = str(now.tm_hour)
        min = str(now.tm_min)
        sec = str(now.tm_sec)

        msg = hour+":"+min+":"+sec+":\t\t"
        msg += content+"\n"

        if os.path.exists(self.fileFullPath):
            pass
        else:
            f = open(self.fileFullPath, 'w')
            f.close()
        
        f = open(self.fileFullPath, 'a')
        f.write(msg)
        f.close()
    
    def makelogfile(self, filepath):
        
        now = time.localtime()
        year = str(now.tm_year)
        mon = str(now.tm_mon)
        day = str(now.tm_mday)
        
        filename = year+"."+mon+"."+day+".txt"

        #self.fileFullPath = filepath + filename
        self.fileFullPath = os.path.join(filepath, filename)
        self.fileFullPath = self.fileFullPath.replace("\\","/")

        if os.path.isdir(filepath):
            pass
        else:
            os.mkdir(filepath)

        if os.path.exists(self.fileFullPath):
            pass
        else:
            f = open(self.fileFullPath, 'w')
            f.close()