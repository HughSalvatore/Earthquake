# -*- coding: utf-8 -*-

import json
import requests
import re
import pymssql
import datetime
import time
import os
import sys
import logging
import inspect
import servicemanager
import win32serviceutil 
import win32service 
import win32event 

class PythonService(win32serviceutil.ServiceFramework): 
    """
    Usage: 'PythonService.py [options] install|update|remove|start [...]|stop|restart [...]|debug [...]'
    Options for 'install' and 'update' commands only:
     --username domain\username : The Username the service is to run under
     --password password : The password for the username
     --startup [manual|auto|disabled|delayed] : How the service starts, default = manual
     --interactive : Allow the service to interact with the desktop.
     --perfmonini file: .ini file to use for registering performance monitor data
     --perfmondll file: .dll file to use when querying the service for
       performance data, default = perfmondata.dll
    Options for 'start' and 'stop' commands only:
     --wait seconds: Wait for the service to actually start or stop.
                     If you specify --wait with the 'stop' option, the service
                     and all dependent services will be stopped, each waiting
                     the specified period.
    """
    #服务名
    _svc_name_ = "PythonService"
    #服务显示名称
    _svc_display_name_ = "Python Service"
    #服务描述
    _svc_description_ = "Python Service"

    def __init__(self, args): 
        win32serviceutil.ServiceFramework.__init__(self, args) 
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.logger = self._getlogger()
        self.run = True


    def _getlogger(self):
        logger = logging.getLogger('[PythonService]')  
          
        this_file = inspect.getfile(inspect.currentframe())  
        dirpath = os.path.abspath(os.path.dirname(this_file))  
        handler = logging.FileHandler(os.path.join(dirpath, "service.log"))  
          
        formatter = logging.Formatter('%(asctime)s %(name)-12s %(levelname)-8s %(message)s')  
        handler.setFormatter(formatter)  
          
        logger.addHandler(handler)  
        logger.setLevel(logging.INFO)

        return logger


    def SvcDoRun(self):
        self.logger.info("service is running...")
        while self.run:
            self.logger.info("running..")
            data = getjson()
            writedata(data)
            time.sleep(290)
            
        # win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE) 
            
    def SvcStop(self):
        self.logger.info("service is stopper...")
        # 先告诉SCM停止这个过程 
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING) 
        # 设置事件 
        win32event.SetEvent(self.hWaitStop)
        self.run = False
        
          
def getjson():
    url = 'http://www.csi.ac.cn/publish/world/world2w.js'
    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_5) AppleWebKit/603.2.4 (KHTML, like Gecko)\
     Version/10.1.1 Safari/603.2.4', 'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'}
    r = requests.get(url, headers=headers)
    if r.status_code == requests.codes.ok:
        data = r.text
        temdata1 = data.replace("date=", "")  # 待优化
        temdata2 = temdata1.replace(";", "")
        re_item = re.compile(r'(?<=[{,])\w+')
        after = re_item.sub("\"\g<0>\"", temdata2)
        print('200')
        return after
    else:
        print('404')


def writedata(data):
    jsondata = json.loads(data)
    conn = pymssql.connect(host="222.18.158.202", user="Hugh", password='Yuson000', database='EESS', charset="utf8")
    cursor = conn.cursor()
    currentime = datetime.datetime.now()
    delta = datetime.timedelta(seconds=300)
    pretime = currentime + delta
    currentimestr = currentime.strftime("%Y-%m-%d %H:%M:%S.0")
    pretimestr = pretime.strftime("%Y-%m-%d %H:%M:%S.0")
    for item in jsondata:
        time = item['time']
        longitude = float(item['weidu'])
        latitude = float(item['jingdu'])
        depth = float(item['depth'])
        magnitude = float(item['Magnitude'])
        location = item['weizhi']
        if time >= pretimestr and time < currentimestr:
            insertsql = "INSERT INTO earinfo\
                        (time,longitude,latitude,depth,earthquakeMagnitude,location)\
                        VALUES ('%s','%f','%f','%f','%f','%s')" % \
                        (time, longitude, latitude, depth, magnitude, location)
            try:
                cursor.execute(insertsql)
                db.commit()
            except Exception as e:
                print(e)
                db.rollback()
    conn.close()
    # db = pymysql.connect("222.18.158.202", user='Hugh', passwd='Yuson000', db='EESS', charset='UTF8')


if __name__=='__main__':
    win32serviceutil.HandleCommandLine(PythonService)
        
    
        









