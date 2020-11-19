
# -*- coding: utf-8 -*-

import pandas as pd
# import os
# import openpyxl
# import xlrd
def getdivided():
	writer = pd.ExcelWriter('text1.xlsx')
	data1 = pd.read_csv('RequestHotel.csv', encoding = "ISO-8859-1", engine='python')
	print(data1)
	df2 =data1[~data1['ID'].isin(['ID'])]

	print(df2)

	df2.to_excel(writer, sheet_name='staff1',index = False)
	writer.save()


	#第二步：读取数据
	iris = pd.read_excel('text1.xlsx')#读入数据文件
	class_list = list(iris['ID'].drop_duplicates())#获取数据class列，去重并放入列表
	# 第三步：按照类别分sheet存放数据
	for i in class_list:
		iris1 = iris[iris['ID']==i]
		iris1.to_excel('Assem/%s.xlsx'%(i))
		#iris1.to_excel('D:/RPA/Project3/Assem/%s.xlsx'%(i))
	writer.save()#文件保存
	writer.close()#文件关闭
	print("getdivided done")

if __name__ =="__main__":
	getdivided()