 #_*_ coding:utf-8 _*
import xlwt
import xlrd
import numpy as np
import pandas as pd  
import os
def getcompare():
	i=0
	#path = r'D:\RPA\Project3\Assem'      # 输入文件夹地址
	path = r'Assem'      # 输入文件夹地址
	files = os.listdir(path)   # 读入文件夹
	num_png = len(files)       # 统计文件夹中的文件个数
	print(num_png)             # 打印文件个数
	# 输出所有文件名
	print("所有文件名:")
	for file in files:
	    print(file) 
	for i in range(num_png):
		#df = pd.read_excel('D:/RPA/Project3/Assem/%s'%(files[i]))
		df = pd.read_excel('Assem/%s'%(files[i]))
		df=df.sort_values(by="Price" , ascending=True)
		df = df.drop(labels=['Unnamed: 0'],axis = 1)
		df.to_excel('Assem/%s'%(files[i]))
	# print(df)









                                        
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             









