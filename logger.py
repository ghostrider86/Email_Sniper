import os
import datetime as dt

class sniper_logger(object):
    """description of class"""

    def __init__(self):
        self.filename = "search.txt"

    def start(self, folder_name):
        if not os.path.isdir(folder_name):
            os.mkdir(folder_name)

        filepath = os.path.join(folder_name, self.filename)
        print("Logging to ",filepath)

        self.f = open(filepath,'a',encoding='utf-8')
        self.f.write('Email Sniper search starting at '+str(dt.datetime.now()))
        self.f.write('\n')

    def write(self,text):
        self.f.write(text+'\n')

    def close(self):
        self.f.close();
