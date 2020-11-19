import tagui as t
import json
import os
import pandas as pd
import re
import datetime
from google.cloud import language_v1
from google.oauth2 import service_account

class RPA1ExtractTravelRequests(object):
	"""docstring for RPA1TravelRequests"""
	def __init__(self, email_info):
		self.email_info=email_info
		self.output_dict={"ID":[],"Name":[],"Hotel":[],"CheckInMonth":[],"CheckInDate":[],"CheckOutMonth":[],"CheckOutDate":[],"Email":[],"Booked":[]}
		self.emails=[]

	def login_email(self):
		t.init()
		t.url('https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=13&ct=1580788659&rver=7.0.6737.0&wp=MBI_SSL&wreply=https%3a%2f%2foutlook.live.com%2fowa%2f%3fnlp%3d1%26RpsCsrfState%3dd234420e-f55a-d62c-a8e6-c1c9a31e4e54&id=292841&aadredir=1&CBCXT=out&lw=1&fl=dob%2cflname%2cwld&cobrandid=90015')
		t.wait(2)
		if not t.present('//*[@id="o365header"]'):
			t.type('//*[@id="i0116"]', self.email_info['email'] + '[enter]')
			t.wait(0.5)
			t.type('//*[@id="i0118"]', self.email_info['password'] + '[enter]')
			t.wait(1)
			if t.present(f"(//div[contains(string(),'Stay signed in')])"):
				print('find stay signed in')
				t.click('//*[@id="idSIButton9"]')
			count = 0
			while True:
				if count > 3:
					print( "Fail.Please Check your Internet Connection.")
					break
				if t.present('//*[@id="O365_NavHeader"]'):
					t.wait(2)
					break
				else:
					t.wait(1)
					count = count + 1

	def search_Unread_TravelRequest(self):
		t.type(f'//input[@aria-label="Search"]','[clear]'+'travel request OR travelrequest')
		t.click(f'//button[@aria-label="Search"]')
		t.wait(1)
		t.click(f'//div[@class="_3lIR1LfENBYPLTyCDNhq5k" and contains(string(),"Filter")]')
		t.wait(1)
		t.click(f'//span[contains(text(),"Unread")]')
		t.wait(2)
		unread_num = t.count(f'//div[@class="_1xP-XmXM1GGHpRKCCeOKjP"]')

		print("unread_num:",unread_num)
		for i in range(1,unread_num+1):
			if t.present(f'//div[contains(string(),"All results")]') and i<4: continue
			else:
				temp={'email_address':"",'text':"",'time':"",'name':"","ID":""}
				print(i)
				t.click(f'(//div[@class="_1xP-XmXM1GGHpRKCCeOKjP"])[{i}]')
				t.wait(1)
				temp['text'] = t.read(f'//div[@class="_3U2q6dcdZCrTrR_42Nxby JWNdg1hee9_Rz6bIGvG1c allowTextSelection"]')
				temp['email_address'] = t.read(f'//div[starts-with(@class,"VHquDtYElxQkNvZKxCJct")]')
				temp['time'] = t.read(f'//div[@class="DWrY3hKxZTZNTwt3mx095"]')
				email = re.search('<(.*)>',temp['email_address'])
				name = re.search('(.*)<',temp['email_address'])
				# employee_id = re.search('employee id is (\d+)|employee id : (\d+)',temp['text'])
				employee_id = re.search(
					re.compile('employee id is ([a-zA-Z0-9]+)|employee id: ?([a-zA-Z0-9]+)', re.IGNORECASE),
					temp['text'])
				if employee_id:
					temp['ID'] = employee_id.group(1)
				if email:
					temp['email_address'] = email.group(1)
				if name:
					temp['name'] = name.group(1).strip()

				print(temp['text'])
				print(temp['email_address'])
				print(temp['time'])
				print(temp['name'])
				t.click(f'//button[@aria-label="More mail actions" and @role="button"]')
				t.wait(0.5)
				t.click(f'//button[@aria-label="Mark as read" and @role="menuitem"]')
				self.emails.append(temp.copy())

				json_file = "emails_json_file.json"
				if self.emails:
					emails_json_file = json.dumps(self.emails)
					with open(json_file, 'w') as fw:
						fw.write(emails_json_file)


	def analyze_entities(self):
		"""
		Analyzing Entities in a String
		Args:
		  text_content The text content to analyze
		"""
		credentials = service_account.Credentials.from_service_account_file(
		'proj3NER-0fda29a98082.json')
		client = language_v1.LanguageServiceClient(credentials = credentials)

		# text_content = 'California is a state.'

		# Available types: PLAIN_TEXT, HTML
		type_ = language_v1.Document.Type.PLAIN_TEXT

		# Optional. If not specified, the language is automatically detected.
		# For list of supported languages:
		# https://cloud.google.com/natural-language/docs/languages


		# Available values: NONE, UTF8, UTF16, UTF32
		encoding_type = language_v1.EncodingType.UTF8
		Month_list=['jan','feb','Mar','apr','may','jun','jul','aug','sep','oct','nov','dec']

		def send(email):
			language = "en"
			document = {"content": email['text'], "type_": type_, "language": language}
			response = client.analyze_entities(request = {'document': document, 'encoding_type': encoding_type})

			slots={"ID":email['ID'],"Name":email['name'],"Hotel":"","CheckInMonth":"","CheckInDate":"","CheckOutMonth":"","CheckOutDate":"","Email":email['email_address'],"Booked":""}
			# Loop through entitites returned from the API
			for entity in response.entities:
				if language_v1.Entity.Type(entity.type_).name == "LOCATION":
					slots["Hotel"] = slots["Hotel"] + entity.name if not slots["Hotel"] else slots["Hotel"] + " "+ entity.name
				elif language_v1.Entity.Type(entity.type_).name == "DATE":
					month,date = 0,0
					date_match = re.search('\d+',entity.name)
					if date_match:
						date = str(int(date_match.group()))
					for i,m in enumerate (Month_list):
						if m in entity.name.lower():
							month = str (i+1)
							break
					if not slots["CheckInMonth"] and not slots["CheckInDate"] :
						slots["CheckInMonth"] = month
						slots["CheckInDate"] = date
					else:
						slots["CheckOutMonth"] = month
						slots["CheckOutDate"] = date

			return slots

		for email in self.emails:
			slots = send(email)
			print("slots:",slots)
			for k,v in slots.items():
				self.output_dict[k].append(v)

	def write_to_csv(self):
		today = datetime.date.today()
		df_output = pd.DataFrame(self.output_dict)
		output_file = "TravelRequest"+str(today)+'.csv'
		df_output.to_csv(output_file,index=False)
		print("[RPA]Write to {} Done".format(output_file))
		t.close()

	def STARTRPA1(self):
		self.login_email()
		self.search_Unread_TravelRequest()
		self.analyze_entities()
		self.write_to_csv()

if __name__ == '__main__':
	email_info = {'email':"businesstravelrpa@outlook.com",'password':"businesstravel"}
	tq = RPA1ExtractTravelRequests(email_info)
	tq.STARTRPA1()
# 	tq.login_email()
# 	tq.search_Unread_TravelRequest()
# 	tq.analyze_entities()
# 	tq.write_to_csv()

	


