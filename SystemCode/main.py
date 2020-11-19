# -*- coding: utf-8 -*-
"""
Created on Sun Nov 15 14:53:15 2020

@author: 77433
"""
from PyQt5.uic import loadUi
import pandas as pd
from datetime import datetime
import sys
import os
from PyQt5.QtWidgets import QFileDialog,QApplication
from PyQt5.QtGui import QIcon
import RPA1
import RPA2
#from RPA2.RPA2 import getresult
import RPA3
from threading import Thread
from time import sleep
class Stats:

    def __init__(self):
        # 从文件中加载UI定义

        # 从 UI 定义中动态 创建一个相应的窗口对象
        # 注意：里面的控件对象也成为窗口对象的属性了
        # 比如 self.ui.button , self.ui.textEdit
        self.ui = loadUi("RPA1.ui")
        self.mes=''
        self.cwd = os.getcwd() # 获取当前程序文件位置
        self.path = self.cwd
        self.email = ''
        self.password = ''
        self.subject=''
        self.assign()
        self.autoset()
        self.prepare = False
        today = datetime.today()
        self.today = f'{today.year}-{today.month}-{today.day}'
        
    def autoset(self):
        pa = ''
        if os.path.exists('path.txt'):
            with open('path.txt','r') as f:
                pa = f.readline()
            self.path =pa
            # if pa!='':
            #     os.chdir(self.path)
            # else:
            #     return 
            self.ui.pathin.setText(self.path)
            if os.path.exists('email.txt'):
                with open('email.txt','r') as f:
                    tt = f.readlines()
            try:
                self.ui.emailin.setText(tt[1][0:-1])
                self.ui.pswin.setText(tt[2])
                sleep(0.1)
            except:
                pass
            if os.path.exists('title.txt'):
                with open('title.txt','r') as f:
                    ti = f.readline()
                self.ui.emailsbin.setText(ti)
            if os.path.exists('template.txt'):
                with open('template.txt','r') as f:
                    temp = f.read()
                self.ui.defaultresponse.setPlainText(temp)
        mes='please make sure your outlook account correct and you do choose the path(before you PRESS GET EMAIL INFO button)'
        self.printmsg(mes)
        mes='You can set your default response in the left text box.Dont change the part with parentheses(before you PRESS Reply E-mail button)'
        self.printmsg(mes)
    def assign(self):
        '''
        assign function to every button
        '''
        self.ui.comfirmbtn.clicked.connect(self.infocomfirm)
        self.ui.pathbtn.clicked.connect(self.choosepath)
        self.ui.SENDBTN.clicked.connect(self.sendemail)
        self.ui.GETBTN.clicked.connect(self.getemail)
    def infocomfirm(self):
        self.printmsg('Info Comfirming...')
        self.getinfo()
        
        
    def choosepath(self):
        
        dir_choose = QFileDialog.getExistingDirectory(None,  
                                    "choose file path",  
                                    self.cwd) # 起始路径
        if dir_choose != "":
            self.path = dir_choose
            self.printmsg('choosed dir:'+dir_choose)
            os.chdir(self.cwd)
            with open('path.txt','wt') as f:
                f.write(dir_choose)
            self.ui.pathin.setText(dir_choose)
            
    def getinfo(self):
        '''
        get configuration information
        '''
        url='https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=13&ct=1580788659&rver=7.0.6737.0&wp=MBI_SSL&wreply=https%3a%2f%2foutlook.live.com%2fowa%2f%3fnlp%3d1%26RpsCsrfState%3dd234420e-f55a-d62c-a8e6-c1c9a31e4e54&id=292841&aadredir=1&CBCXT=out&lw=1&fl=dob%2cflname%2cwld&cobrandid=90015'
        email = self.ui.emailin.text()
        psw = self.ui.pswin.text()
        sb = self.ui.emailsbin.text()
        self.email=email
        self.password = psw
        self.subject = sb
        # pa = self.ui.pathin.text()
        # if pa!='':
        #     self.path = pa
        # os.chdir(self.path)
        with open('email.txt','wt') as f:
            f.write(url+'\n')
            f.write(email+'\n')
            f.write(psw)
        #print(self.path)
        if email != '':
            self.printmsg('get account:'+email)
        info = self.ui.defaultresponse.toPlainText()
        if info!='':
            with open('template.txt','wt') as f:
                f.write(info)
            self.printmsg('set defualt response')
        #os.chdir(self.path)
        with open('title.txt','wt') as f:
            f.write(self.subject)
        self.ui.GETBTN.setEnabled(True)
        self.ui.SENDBTN.setEnabled(True)
        
    def printmsg(self,mes):
        time = datetime.now()
        st = f'[{time.hour}:{time.minute}:{time.second}]--'
        self.ui.messagebox.append(st+mes)
        
    def sendemail(self):
        self.ui.GETBTN.setEnabled(False)
        #os.chdir(self.path)
        if os.path.exists("email.txt"):
            self.printmsg('"email.txt" has been found')
            if os.path.exists("template.txt"):
                self.printmsg('"template.txt" has been found')
                TableName = "hotelinfo"+self.today+'.xlsx'
                if os.path.exists(TableName):
                    self.printmsg('start process...')
                    sleep(1)
                    self.printmsg('title:'+self.subject)
                    t=self.subject
                    sendthread = Thread(target=RPA3.sendemail,args=(t,0))
                    def detect():
                        filename = 'Report'+self.today+'.xlsx'
                        while(1):
                            if not os.path.exists(filename):
                                sleep(1)
                                pass
                            else:
                                self.transfer(filename)
                                self.printmsg(f'Send Email Done! go to {self.path} and check table3 for today')
                                break
                                
                            
                    #RPA3.sendemail(t)
                    detectthread = Thread(target=detect)
                    detectthread.start()
                    sendthread.start()
                    #sendthread.join()
                    #self.printmsg(f'Send Email Done! go to {self.path} and check table3.xls')
                else:
                    self.printmsg('[exception]make sure you have update table2 and rename it as "TABLE2UPDATE.xls" ')
            else: 
                self.printmsg("[exception]No response tempalte file found,please make sure you had correct file path and PRESSED 'OK' already")
                return 
                
    def getemail(self):
        self.ui.SENDBTN.setEnabled(False)
        self.printmsg('start to get email info...')
        email_info = {'email':self.email,'password':self.password}
        tq = RPA1.RPA1ExtractTravelRequests(email_info)
        getthread = Thread(target=tq.STARTRPA1)
        #os.chdir(self.path)
        output_file = "TravelRequest"+self.today+'.csv'
        def prepared():
            while(1):
                if not os.path.exists(output_file):
                    sleep(1)
                else:
                    self.printmsg('Already get info from email...')
                    self.prepare = True
                    self.transfer(output_file)
                    self.gethotel()
                    break
        preparedthread = Thread(target=prepared)
        preparedthread.start()
        if not os.path.exists(output_file):
            os.chdir(self.cwd)
            getthread.start()
            
    
    def gethotel(self):
        self.printmsg('start to grab hotel price info...')
        RPA2thread = Thread(target=RPA2.getresult)
        #os.chdir(self.path)
        
        newfile = "hotelinfo"+self.today+'.xlsx'
        def complete():
            while(1):
                if not os.path.exists(newfile):
                    sleep(1)
                else:
                    self.printmsg('Build hotel info...')
                    self.transfer(newfile)
                    self.printmsg('go check hotel info table and start booking!')
                    break
       
        comthread = Thread(target = complete)
        comthread.start()
        if not os.path.exists(newfile):
            os.chdir(self.cwd)
            RPA2thread.start()
        
    def transfer(self,filename):
        os.chdir(self.cwd)
        if 'csv' in filename:
            df = pd.read_csv(filename)
            os.chdir(self.path)
            df.to_excel(filename[0:-4]+'.xlsx')
        else:
            df = pd.read_excel(filename)
            os.chdir(self.path)
            df.to_excel(filename)
        os.chdir(self.cwd)
        self.printmsg(f"building {filename}")
        # os.remove(filename)
        
        
        
app = QApplication([])
app.setWindowIcon(QIcon('COFFEE.png'))
stats = Stats()
stats.ui.show()
app.exec_()