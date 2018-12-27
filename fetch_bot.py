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
	print('run')
	fetch_admin = Skype('cuongdaovan262@gmail.com', 'developer26297@') # skype account
	fetch_id = "19:6fc079db903d48babebfa404621b1457@thread.skype"
	test_id = "19:b098950b91e14699b97011b951b2b3aa@thread.skype"
	fetch_group = fetch_admin.chats[fetch_id]
	# fetch_group = fetch_admin.chats[test_id] # group chat
	fetch_group.setTyping(active=True)
	username = '' # lay ra ten user da gui
	sheet = None
	all_key = ''
	scope = [
		'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive',
    ]
	creds = ServiceAccountCredentials.from_json_keyfile_name('fetch_sheet.json', scope)
	client = gspread.authorize(creds)

	def sheet(self):
		"""
		doc sheet lien tuc
		# need_to_do
		dang co bug
		"""
		while True:
			s = self.client.open('fetch').sheet1
			self.sheet = s.get_all_records()
			time.sleep(10)

	def all_key(self):
		"""
		get all key of key-answer
		display like:
		- key1
		- key2
		"""
		keys = ''
		for dic in self.sheet:
			keys += '- '+ dic['key']+'\n'
		return keys

	def update_order(self):
		sheet_order = self.client.open('Order cơm trưa').sheet1
		s = {} # dictionary of other order
		p = {} # dictionary of rice order
		for dic in sheet_order.get_all_records()[1:]:
			"""
			kiem tra xem co mon trong dict chua 
			neu chua co thi them vao dict, neu co roi thi cap nhat dict
			"""
			if dic['Các món khác'] != '':
				if dic['Các món khác'] in s: 
					s[dic['Các món khác']] += dic['Tên'] + ', '
				else:
					s.update({dic['Các món khác']: dic['Tên']+', '})
			if dic['Loại'] != '':
				if dic['Loại'] in p:
					p[dic['Loại']] += dic['Tên'] + ', '
				else:
					p.update({dic['Loại']: dic['Tên']+', '})
		"""
		gui thong tin dat com len group
		"""
		msg = 'fetch_admin xác nhận: \n'
		for key in s:
			msg += str(key) + ': ' + str(s[key])+'\n'
		for key in p:
			msg += str(key) + ': ' + str(p[key])+'\n'
		for dic in self.sheet:
			if dic['key'].find('link order cơm') >= 0:
				msg += dic['answer']
		self.fetch_group.sendMsg(msg, rich=True)

	def notify(self):
		"""
		gui thong bao theo gio
		#need_to_do
		check theo thứ
		"""
		while True:
			now = datetime.now()
			days= ["monday","tuesday","wednesday","thursday","friday"]
			if now.strftime("%A").lower() in days:
				if datetime.now().strftime("%H:%M:%S") == "09:00:01":
					print('true')
					m = 'chúc <at id="*">all</at> một ngày làm việc vui vẻ và hiệu quả (flex)(flex) \n'
					m += "xem các option: @Fetch_admin --help"
					self.fetch_group.sendMsg(m, rich=True)
				if datetime.now().strftime("%H:%M:%S") == "10:00:01":
					for dic in self.sheet:
						if dic['key'].find('đặt cơm') >= 0:
							self.fetch_group.sendMsg(dic['answer'])
				if datetime.now().strftime("%H:%M:%S") == "10:55:01":
					self.update_order()
				if datetime.now().strftime("%H:%M:%S") == "11:58:01":
					self.fetch_group.sendMsg('mọi người nghỉ tay đi ăn cơm đi ạ (sun)(sun)')
				time.sleep(1)

	def msg(self):
		while True:
			events = self.fetch_admin.getEvents() # get tất cả các event trên skype
			if events != []:
				event = events[0]
				if event.type == "NewMessage": # kiểm tra xe có message mới không
					
					if hasattr(event, 'user'): # kiểm tra xem có user variable không
						self.username = str(event.user.name)

					if hasattr(event, 'msg'): # kiêm tra xem tồn tại msg variable không
						msg = event.msg
						content = msg.content
						if content != '':
							# kiểm tra xem có mark @Fetch_admin không
							if content.startswith('<at id="8:live:cuongdaovan262">Fetch_admin</at>'):
								# cut lấy msg
								mess = content[len('<at id="8:live:cuongdaovan262">Fetch_admin</at> '):]
								for dic in self.sheet:
									if mess.find(dic['key']) >= 0: # kiểm tra xem có key trong list key không
										self.fetch_group.sendMsg(dic['answer'], rich=True)
								if mess.find('--help') >= 0:
									msg = 'dưới đây là những option: \n'+self.all_key()
									self.fetch_group.sendMsg(msg, rich=True)
								if mess.find('update order') >= 0:
									self.update_order()
			time.sleep(0.1) # không được sleep trên 1s 

bot = Bot()
t1 = threading.Thread(target=bot.sheet)
t2 = threading.Thread(target=bot.notify)
t3 = threading.Thread(target=bot.msg)

t1.start()
t2.start()
t3.start()