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
    fetch_admin = Skype(
        'cuongdaovan262@gmail.com',
        'developer26297@',
        tokenFile='token.txt'
    )  # skype account
    # fetch_id = "19:6fc079db903d48babebfa404621b1457@thread.skype"
    test_id = "19:b098950b91e14699b97011b951b2b3aa@thread.skype"
    # test1 = "19:1e00439c6b8c428f9a42c79bc84ee641@thread.skype"
    # fetch_group = fetch_admin.chats[fetch_id]
    fetch_group = fetch_admin.chats[test_id]  # group chat
    option = 1000
    username = ''  # lay ra ten user da gui
    sheet = None
    all_key = ''
    spams = {}
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive',
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        'fetch_sheet.json', scope)
    client = gspread.authorize(creds)

    def sheet(self):
        """
        doc sheet lien tuc
        # need_to_do
        bug: bi du dot ngot
        """
        while True:
            s = self.client.open('fetch').sheet1
            self.sheet = s.get_all_records()
            time.sleep(60)

    def all_key(self):
        """
        get all key of key-answer
        display like:
        - key1
        - key2
        """
        keys = []
        for dic in self.sheet:
            keys.append(dic['key'])
        return keys

    def sendMsg(self, group, msg='', rich=False, typing=False):
        group.setTyping(typing)
        group.sendMsg(msg, rich=rich)

    def check_time(self, now, time=''):
        if now == time:
            return True
        else:
            return False

    def update_order(self):
        sheet_order = self.client.open('Order cơm trưa').sheet1
        s = {}  # dictionary of other order
        p = {}  # dictionary of rice order
        for dic in sheet_order.get_all_records()[1:]:
            """
            kiem tra xem co mon trong dict chua 
            neu chua co thi them vao dict, neu co roi thi cap nhat dict
            """
            if dic['Các món khác'] != '':
                if dic['Các món khác'] in s:
                    s[dic['Các món khác']] += dic['Tên'] + ', '
                else:
                    s.update({dic['Các món khác']: dic['Tên'] + ', '})
            if dic['Loại'] != '':
                if dic['Loại'] in p:
                    p[dic['Loại']] += dic['Tên'] + ', '
                else:
                    p.update({dic['Loại']: dic['Tên'] + ', '})
        """
        gui thong tin dat com len group
        """
        msg = '<at id="*">all</at> fetch_admin xác nhận: \n'
        for key in s:
            msg += str(key) + ': ' + str(s[key]) + '\n'
        for key in p:
            msg += str(key) + ': ' + str(p[key]) + '\n'
        for dic in self.sheet:
            if dic['key'].find('link order') >= 0:
                msg += dic['answer']
        self.sendMsg(self.fetch_group, msg=msg, rich=True, typing=True)

    def notify(self):
        """
        gui thong bao theo gio
        #need_to_do
        check theo thứ
        """
        while True:
            now = datetime.now()
            days = ["monday", "tuesday", "wednesday", "thursday", "friday"]
            if now.strftime("%A").lower() in days:
                if self.check_time(datetime.now().strftime("%H:%M:%S"), "00:18:01"):
                    print('true')
                    msg = 'chúc <at id="*">all</at> một ngày làm việc vui vẻ và hiệu quả\n'
                    msg += "xem các option: @Fetch_admin --help"
                    self.sendMsg(self.fetch_group, msg=msg,
                                 rich=True, typing=True)
                    time.sleep(1)
                if self.check_time(datetime.now().strftime("%H:%M:%S"), "10:07:05"):
                    for dic in self.sheet:
                        if dic['key'].find('order') >= 0:
                            self.sendMsg(
                                self.fetch_group, msg=dic['answer'], rich=True, typing=True)
                    time.sleep(1)
                if self.check_time(datetime.now().strftime("%H:%M:%S"), "11:03:01"):
                    self.update_order()
                    time.sleep(1)
                if self.check_time(datetime.now().strftime("%H:%M:%S"), "11:58:01"):
                    msg = "mọi người nghỉ tay đi ăn cơm đi ạ (sun)(sun)"
                    self.sendMsg(self.fetch_group, msg=msg,
                                 rich=False, typing=True)
                    time.sleep(1)
                if self.check_time(datetime.now().strftime("%H:%M"), "00:45"):
                    worksheet = self.client.open('Order cơm trưa').sheet1
                    cell_list = worksheet.range('C3:H200')
                    for cell in cell_list:
                        cell.value = ''
                    worksheet.update_cells(cell_list)
                    time.sleep(60)
                time.sleep(0.5)

    def msg(self):
        while True:
            events = self.fetch_admin.getEvents()  # get tất cả các event trên skype
            if events != []:
                event = events[0]
                if event.type == "NewMessage":  # kiểm tra xe có message mới không
                    if hasattr(event, 'user'):  # kiểm tra xem có user variable không
                        self.username = str(event.user.name)
                    if hasattr(event, 'msg'):  # kiêm tra xem tồn tại msg variable không
                        msg = event.msg
                        content = msg.content
                        if msg.chat.id == self.test_id:
                            if content != None:
                                # kiểm tra xem có mark @Fetch_admin không
                                if content.startswith('<at id="8:live:cuongdaovan262">Fetch_admin</at>'):
                                    # cut lấy msg
                                    mess = content[len(
                                        '<at id="8:live:cuongdaovan262">Fetch_admin</at>'):]
                                    try:
                                        self.option = int(mess)
                                        print('int')
                                    except ValueError:
                                        print('not int')
                                    print('cuong')
                                    for dic in self.sheet:
                                        if self.option <= len(self.all_key()) and self.option > 0:
                                            if dic['key'] == self.all_key()[self.option - 1]:
                                                msg = dic['answer']
                                                self.sendMsg(
                                                    self.fetch_group, msg=msg, rich=True, typing=True)
                                                if self.all_key()[self.option - 1].find('update order') >= 0:
                                                    self.update_order()
                                                self.option = 10000
                                        # kiểm tra xem có key trong list key khong
                                        if mess.lower().find(dic['key'].lower()) >= 0:
                                            self.fetch_group.sendMsg(
                                                dic['answer'], rich=True)
                                    if mess.find('--help') >= 0:
                                        msg = 'dưới đây là những option: \n'
                                        for key in self.all_key()[0:]:
                                            msg += str(self.all_key().index(key) + 1) + ". " + key + "\n"
                                        self.sendMsg(self.fetch_group, msg=msg, rich=True, typing=True)
                                    if mess.find('update order') >= 0:
                                        self.update_order()
                events = []
            time.sleep(0.5)  # không được sleep trên 1s

    # def spam(self):
    #     while True:
    #     	# print('spam')
    #         events = self.fetch_admin.getEvents()
    #         # print(events)
    #         if events != []:
    #             event = events[0]
    #             if event.type == "NewMessage":
    #                 if hasattr(event, 'msg'):
    #                 	msg = event.msg
    #                 	print(msg.user)
    #                 	if str(msg.user.name) in self.spams:
    #                 		self.spams[str(msg.user.name)] += 1
    #                 	else:
    #                 		self.spams.update({str(msg.user.name): 1})
    #                 	# print(msg.user)
    #                 	time.sleep(0.1)


# # print(self.spams)


bot = Bot()
t1 = threading.Thread(target=bot.sheet)
t2 = threading.Thread(target=bot.notify)
t3 = threading.Thread(target=bot.msg)
# t4 = threading.Thread(target=bot.spam)

t1.start()
t2.start()
t3.start()
# t4.start()
