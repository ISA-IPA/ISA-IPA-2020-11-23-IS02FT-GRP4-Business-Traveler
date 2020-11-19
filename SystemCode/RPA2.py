# -*- coding: utf-8 -*-
"""

RPA2 
@author: lijiayi
"""
import pandas as pd  
import numpy as np
import tagui as t
from final import getfinal
from compare import getcompare
from excel import getdivided
from datetime import datetime
from time import sleep
import os

def getrawdata():
	sleep(3)
	today = datetime.today()
	today = f'{today.year}-{today.month}-{today.day}'
	output_file = "TravelRequest"+today+'.csv'
	dff = pd.read_csv(output_file)
	#thismonth = input("enter this month:")
	thismonth = datetime.now().month
	n = dff.shape[0]
	i=0
	checkinDatehis  = 0
	checkoutDatehis = 0
	checkinMonthhis = 0
	for i in range(n):
		# Resinfo= dff.iloc[i,1]+ " " + dff.iloc[i,2]
		Resinfo = dff.iloc[i,2]
		print(Resinfo)
		dff2  = dff.iloc[i,4] 
		dff3  = dff.iloc[i,6]
		sno   = dff.iloc[i,0]
		name  = dff.iloc[i,1]
		email = dff.iloc[i,7]
		book  = "No"
		checkinDate = dff2
		checkoutDate = dff3
		checkinMonth  = dff.iloc[i,3]
		checkoutMonth = dff.iloc[i,5]

	# dff1= dff.iloc[0,1] + " " + dff.iloc[0,2]
	# print(dff1)
		t.init(visual_automation = True, chrome_browser = True)
		#t.init(visual_automation = False, chrome_browser = True)
		urlFirst = 'https://www.google.com/'

		t.url(urlFirst)

		
		item_list = []
		price_list = []
		hotelurl_list = []
		


		t.wait(4)
		t.type('//input[@name="q"]', Resinfo)
		t.click('//input[@name="btnK"]')
		t.wait(4)
		t.click('//div[contains(@class, "GJp0Vb")]')
		#UjxDhc
		t.wait(3)
		t.click('//div[contains(@class, "UjxDhc")]')
		t.wait(3)
		t.click('//a[contains(@class, "oDcTJb vpggTd Ubyxue ZdJdwb ed5F6c")]')
		t.wait(3)
		t.click('//div[contains(@class, "njjicd U0gh4b")]')
		t.click('//div[contains(@class, "njjicd U0gh4b")]')
		t.click('//div[contains(@class, "njjicd U0gh4b")]')
		t.click('//div[contains(@class, "njjicd U0gh4b")]')
		t.click('//div[contains(@class, "njjicd U0gh4b")]')
		t.click('//div[contains(@class, "njjicd U0gh4b")]')
		t.click('//div[contains(@class, "njjicd U0gh4b")]')
		t.click('//div[contains(@class, "njjicd U0gh4b")]')
		t.click('//div[contains(@class, "njjicd U0gh4b")]')
		t.click('//div[contains(@class, "njjicd U0gh4b")]')
		t.click('//div[contains(@class, "njjicd U0gh4b")]')
		t.click('//div[contains(@class, "njjicd U0gh4b")]')




		thismonth1= int(thismonth)
		if checkinMonth == thismonth1 + 1 or thismonth1 - checkinMonth == 11:
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			if checkinDate == checkinDatehis and checkoutDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
			if checkinDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('visual2/'+ str(checkinDate) + '.png')
				t.wait(3)
				t.click('visual/'+ str(checkoutDate) + '.png')
				t.wait(3)
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
		elif checkinMonth - thismonth1 == 2 or thismonth1 - checkinMonth == 10:
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			if checkinDate == checkinDatehis and checkoutDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
			if checkinDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('visual2/'+ str(checkinDate) + '.png')
				t.wait(3)
				t.click('visual/'+ str(checkoutDate) + '.png')
				t.wait(3)
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
		elif checkinMonth - thismonth1 == 3 or thismonth1 - checkinMonth == 9:
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			if checkinDate == checkinDatehis and checkoutDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
			if checkinDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('visual2/'+ str(checkinDate) + '.png')
				t.wait(3)
				t.click('visual/'+ str(checkoutDate) + '.png')
				t.wait(3)
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
		elif checkinMonth - thismonth1 == 4 or thismonth1 - checkinMonth == 8:
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			if checkinDate == checkinDatehis and checkoutDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
			if checkinDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('visual2/'+ str(checkinDate) + '.png')
				t.wait(3)
				t.click('visual/'+ str(checkoutDate) + '.png')
				t.wait(3)
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
		elif checkinMonth - thismonth1 == 5 or thismonth1 - checkinMonth == 7:
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			if checkinDate == checkinDatehis and checkoutDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
			if checkinDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('visual2/'+ str(checkinDate) + '.png')
				t.wait(3)
				t.click('visual/'+ str(checkoutDate) + '.png')
				t.wait(3)
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
		elif checkinMonth - thismonth1 == 6 or thismonth1 - checkinMonth == 6:
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			if checkinDate == checkinDatehis and checkoutDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
			if checkinDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('visual2/'+ str(checkinDate) + '.png')
				t.wait(3)
				t.click('visual/'+ str(checkoutDate) + '.png')
				t.wait(3)
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
		elif checkinMonth - thismonth1 == 7 or thismonth1 - checkinMonth == 5:
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			if checkinDate == checkinDatehis and checkoutDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
			if checkinDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('visual2/'+ str(checkinDate) + '.png')
				t.wait(3)
				t.click('visual/'+ str(checkoutDate) + '.png')
				t.wait(3)
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
		elif checkinMonth - thismonth1 == 8 or thismonth1 - checkinMonth == 4:
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			if checkinDate == checkinDatehis and checkoutDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
			if checkinDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('visual2/'+ str(checkinDate) + '.png')
				t.wait(3)
				t.click('visual/'+ str(checkoutDate) + '.png')
				t.wait(3)
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
		elif checkinMonth - thismonth1 == 9 or thismonth1 - checkinMonth == 3:
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			if checkinDate == checkinDatehis and checkoutDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
			if checkinDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('visual2/'+ str(checkinDate) + '.png')
				t.wait(3)
				t.click('visual/'+ str(checkoutDate) + '.png')
				t.wait(3)
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
		elif checkinMonth - thismonth1 == 10 or thismonth1 - checkinMonth == 2:
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			if checkinDate == checkinDatehis and checkoutDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
			if checkinDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('visual2/'+ str(checkinDate) + '.png')
				t.wait(3)
				t.click('visual/'+ str(checkoutDate) + '.png')
				t.wait(3)
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
		elif checkinMonth - thismonth1 == 11 or thismonth1 - checkinMonth == 1:
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(3)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			t.click('//div[contains(@class, "njjicd rSFy9b")]')
			t.wait(2)
			if checkinDate == checkinDatehis and checkoutDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
			if checkinDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('visual2/'+ str(checkinDate) + '.png')
				t.wait(3)
				t.click('visual/'+ str(checkoutDate) + '.png')
				t.wait(3)
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')



		else:
			t.wait(3)
			if checkinDate == checkinDatehis and checkoutDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')
			if checkinDate == checkoutDatehis and checkinMonth == checkinMonthhis:
				t.click('visual2/'+ str(checkinDate) + '.png')
				t.wait(3)
				t.click('visual/'+ str(checkoutDate) + '.png')
				t.wait(3)
				t.click('//div[contains(@class, "fSXkBc")]')
				t.wait(2)
				t.click('//a[contains(@class, "vJdf1c")]')

		t.click('visual/'+ str(checkinDate) + '.png')# num, f'{num}.png'
		t.wait(3)
		
		t.click('visual/'+ str(checkoutDate) + '.png')
		t.wait(3)
		
		t.click('//div[contains(@class, "fSXkBc")]')
		t.wait(3)
		t.click('//a[contains(@class, "vJdf1c")]')

		t.wait(4)
		#//*[@id="tsuidfoqvX7iCCNOP9QObnZqwDw1"]/div[1]/two-month-calendar/div/div/calendar-month[1]/calendar-week[3]/calendar-day[6]
		num_rows = t.count('//div[contains(@class, "t7H34")]')
		print(num_rows)
		for n in range(1, num_rows//2+1):
			item_selector = f'(//span[contains(@class, "NiGhzc")])[{n}]'
			t.hover(item_selector)
			item = t.read(item_selector)
			print(f"{n}. item = " + item)
			
			
			

			item_list.append(item)

			price_selector = f'(//span[contains(@class, "MW1oTb")])[{n}]'

			price = 0
			if t.present(price_selector):
				price = t.read(price_selector)
				print("    price = " + price)
			price_list.append(price)

			hotelurl_selector =  f'(//a[contains(@class, "hUGVEe jC2IFe")])[{n}]/@href'
			hotelurl = t.read(hotelurl_selector)
			print("    hotelurl = "  + hotelurl)
			
			hotelurl_list.append("https://www.google.com" + hotelurl)


		print(item_list)
		print(price_list)
		print(hotelurl_list)
		#df = pd.DataFrame({'ID':sno,'NAME':name,'EMAIL':email,'Hotel':Resinfo,'BookingPlatform': item_list, 'Price': price_list, 'BookingURL':hotelurl_list,'ISBook':book })
		df = pd.DataFrame({'ID':sno,'NAME':name,'CheckinMonth':checkinMonth, 'CheckinDate':checkinDate, 'CheckoutMonth':checkoutMonth, 'CheckoutDate':checkoutDate,'EMAIL':email,'Hotel':Resinfo,'BookingPlatform': item_list, 'Price': price_list, 'BookingURL':hotelurl_list,'ISBook':book })
		#df = pd.DataFrame({'Sno':sno, 'Item': item_list, 'Price': price_list, 'Hotelurl':hotelurl_list })

		# write to csv
		df.to_csv ('RequestHotel.csv',mode='a', index = None)
		checkinDatehis  = checkinDate
		checkoutDatehis = checkoutDate
		checkinMonthhis = checkinMonth

		t.close()


def getresult():
	today = datetime.today()
	today = f'{today.year}-{today.month}-{today.day}'
	output_file = "TravelRequest"+today+'.csv'
	try:
		test__s = pd.read_csv(output_file)
	except:
		return
	getrawdata()
	getdivided()
	getcompare()
	getfinal()


if __name__ =='__main__':
	getresult()

	

