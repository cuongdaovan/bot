from datetime import time
from datetime import datetime
import time
from threading import Thread
import threading

from skpy import Skype, SkypeChats
from skpy import SkypeContact
import gspread
from oauth2client.service_account import ServiceAccountCredentials


class Bot:
	fetch_admin = Skype('cuongdaovan262@gmail.com', 'developer26297@')
	# ch = sk.chats["19:6fc079db903d48babebfa404621b1457@thread.skype"]
	fetch_group = fetch_admin.chats["19:b098950b91e14699b97011b951b2b3aa@thread.skype"]
	print('run')
	username = ''
	sheet = None
	all_key = ''
	scope = [
		'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive',
    ]
	creds = ServiceAccountCredentials.from_json_keyfile_name('fetch_sheet.json', scope)
	client = gspread.authorize(creds)
	def sheet(self):

		s = self.client.open('fetch').sheet1
		self.sheet = s.get_all_records()
		for dic in self.sheet:
			self.all_key += '- '+ dic['key']+'\n'

	def notify(self):
		while True:
			if datetime.now().strftime("%H:%M:%S") == "19:11:01":
				print('true')
				m = "chúc mọi người một ngày làm việc vui vẻ và hiệu quả (flex)(flex) \n"
				m += "xem các option: @Fetch_admin --help"
				self.fetch_group.sendMsg(m)
			if datetime.now().strftime("%H:%M:%S") == "20:45:01":
				for dic in self.sheet:
					if dic['key'].find('đặt cơm') >= 0:
						self.fetch_group.sendMsg(dic['answer'])
			if datetime.now().strftime("%H:%M:%S") == "21:41:01":
				sheet_order = self.client.open('Order cơm trưa').sheet1
				s = {}
				p = {}
				# print(sheet.get_all_records())
				for dic in sheet_order.get_all_records()[1:]:
					# print(dic['Các món khác'])
					if dic['Các món khác'] != '':
						if dic['Các món khác'] in s:
							s[dic['Các món khác']] += dic['Tên'] + ', '
						else:
							s.update({dic['Các món khác']: dic['Tên']+', '})
					if dic['Loại'] != '':
						if dic['Loại'] in p:
							p[dic['Các món khác']] += dic['Tên'] + ', '
						else:
							p.update({dic['Loại']: dic['Tên']+', '})
				msg = 'fetch_admin xác nhận: \n'
				for key in s:
					msg += key + ': ' + s[key]+'\n'
				for key in p:
					msg += key + ': ' + p[key]+'\n'
				for dic in self.sheet:
					if dic['key'].find('link order cơm') >= 0:
						msg += dic['answer']
				self.fetch_group.sendMsg(msg)
			if datetime.now().strftime("%H:%M:%S") == "11:58:01":
				self.fetch_group.sendMsg('mọi người nghỉ tay đi ăn cơm đi ạ (sun)(sun)')
			time.sleep(1)

	def msg(self):
		while True:
			events = self.fetch_admin.getEvents()
			if events != []:
				event = events[0]
				if event.type == "NewMessage":
					
					if hasattr(event, 'user'):
						self.username = str(event.user.name)

					if hasattr(event, 'msg'):
						# mess = 'hello '+ event.user
						msg = event.msg
						content = msg.content
						if content.startswith('<at id="8:live:cuongdaovan262">Fetch_admin</at>'):

							# print(self.username)
							mess = content[len('<at id="8:live:cuongdaovan262">Fetch_admin</at> '):]
							for dic in self.sheet:
								if mess.find(dic['key']) >= 0:
									self.fetch_group.sendMsg(dic['answer'])
							if mess.find('--help') >= 0:
								msg = 'dưới đây là những option: \n'+self.all_key
								self.fetch_group.sendMsg(msg)
			time.sleep(0.1)

bot = Bot()
bot.sheet()
t1 = threading.Thread(target=bot.notify)
t2 = threading.Thread(target=bot.msg)
t1.start()
t2.start()

