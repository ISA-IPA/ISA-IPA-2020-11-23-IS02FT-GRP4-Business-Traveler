# -*- coding: utf-8 -*-
"""
Created on Sat Nov 14 15:15:57 2020
RPA3 : use to send email automatically [] '' _
@author: 77433
"""
import tagui as t
import pandas as pd
import datetime
import os
PSW=''
ACCOUNT=''
LOGINURL=''
FileName = 'email.txt'
today = datetime.date.today()
TableName = "hotelinfo"+str(today)+'.xlsx'
TemplateName = 'template.txt'
def GetAccount():
    '''

    Returns  no return ,set PSW,ACCOUNT,LOGINURL according to filename
    -------
    None.

    '''
    with open (FileName,'r') as f:
        tt = f.readlines()
    global PSW,ACCOUNT,LOGINURL
    PSW=tt[2]
    ACCOUNT=tt[1]
    LOGINURL=tt[0]
    print('init account Done!')

def getinfo():
    '''
    get relevant info from table2upadate
    '''
    data = pd.read_excel(TableName)
    return data



'''
def buildmessage(name,destination):
    a = 'Dear '+ name + ','
    b = 'the hotel for Your travel to '+destination+' has been booked for you successfully!'

    c='if you dont receive the relevant message on your mobile ,please contact with me.'

    d='Regards,'
    e='xxx'
    message=a+'[enter]'+b+'[enter]'+c+'[enter]'+d+'[enter]'+e+'[enter]'
    return message
'''

def buildmessage(name='n',destination='hotel',indate='in',outdate='out',price='p'):
    with open (TemplateName,'r') as f:
        tt = f.readlines()
    message = ''
    for i in tt:
        if 'name' in i:
            i=i.replace('(name)',name)
        if 'destination' in i:
            i=i.replace('(destination)',destination)
        if 'indate' in i:
            i=i.replace('(indate)',indate)
        if 'outdate' in i:
            i=i.replace('(outdate)',outdate)
        if 'price' in i:
            i=i.replace('(price)',price)
        message = message+i
    print(message)
    return message

def sendemail(title,s=0):
    '''
    use the account to login and send email
    '''
    GetAccount()
    data = getinfo()
    t.init(visual_automation = True, chrome_browser = True)
    t.url(LOGINURL)
    t.type('//*[@id="i0116"]', ACCOUNT + '[enter]')
    t.wait(0.5)
    t.type('//*[@id="i0118"]',PSW + '[enter]')
    if t.present('//*[@id="idBtn_Back"]'):
        #t.click('//*[@id="idSIButton9"]')
        t.click('//*[@id="idBtn_Back"]')
    if t.present('//*[@id="id__5"]'):
        print('Email login sucessful')   
        t.click('//*[@id="id__5"]')
    # start send
    situation = []
    for i in range(data.shape[0]):
        
        if data.loc[i,'ISBook']=='YES'or data.loc[i,'ISBook']=='Yes':
            name = data.loc[i,'NAME']
            email = data.loc[i,'EMAIL']
            destination = data.loc[i,'Hotel']
            indate = str(data.loc[i,'CheckinMonth'])+'-'+str(data.loc[i,'CheckinDate'])
            outdate = str(data.loc[i,'CheckoutMonth'])+'-'+str(data.loc[i,'CheckoutDate'])
            price = data.loc[i,'Price']
            message = buildmessage(name,destination,indate,outdate,price)
            try:
                t.click('//*[@id="id__5"]')
                        
                t.type('//*[@aria-label="To"]',email)
                #t.type('//*[@id="app"]/div/div[2]/div[1]/div[1]/div[3]/div[2]/div/div[3]/div[1]/div/div/div/div[1]/div[1]/div[1]/div/div[1]/div/div/div/div/div[1]/div/div/input',email)
                t.type('//*[@maxlength="255"]',title)
                #//*[@id="app"]/div/div[2]/div[1]/div[1]/div[3]/div[2]/div/div[3]/div[1]/div/div/div/div[1]/div[1]/div[3]/div[2]/div/div/div
                t.type('//*[@role="textbox"]',message)
                #t.type('//*[@id="app"]/div/div[2]/div[1]/div[1]/div[3]/div[2]/div/div[3]/div[1]/div/div/div/div[1]/div[2]/div[1]',message)
                t.click('//*[starts-with(@class,"ms-Button-label")and contains(string(),"Send")]')
                situation.append('Done')
            except:
                situation.append('ERROR')
                
        else:
            situation.append('No BOOKED YET')
    data.loc[:,'SITUATION'] = situation
    today = datetime.date.today()
    filename = "Report"+str(today)+'.xlsx'
    data.to_excel(filename)
    print('DONE!')
    #logout()
    t.close()
    

if __name__ =='__main__':
    title = 'Hotel Booking Done'
    sendemail(title)

