import numpy as np
import pandas as pd
import glob
import csv
from datetime import datetime, date
today = date.today()
import shutil
import os
def getfinal():
	all_data = pd.DataFrame()
#	for f in glob.glob("D:/RPA/Project3/Assem/*.xlsx"):
	for f in glob.glob("Assem/*.xlsx"):
		df = pd.read_excel(f).head(1)
		all_data = all_data.append(df,ignore_index=True)

	# now save the data frame
	writer = pd.ExcelWriter("hotelinfo"+str(today)+'.xlsx')
	all_data = all_data.drop(labels=['Unnamed: 0'],axis = 1)

	all_data.to_excel(writer,'sheet1')
	writer.save()
	writer.close()
	df1=pd.read_excel("hotelinfo"+str(today)+'.xlsx')
	df1 = df1.drop(labels=['Unnamed: 0'],axis = 1)
	df1.to_csv("hotelinfo"+str(today)+'.csv')
	os.remove('RequestHotel.csv')
	os.remove('test1.csv')
	try:
		shutil.rmtree('Assem')
		print('buffer clear')
	except:
		pass
	os.mkdir('Assem')
