from datetime import time
from datetime import datetime
import time
from threading import Thread
import threading
import random
import os
from collections import deque

from skpy import Skype, SkypeChats
from skpy import SkypeContact, SkypeAuthException, SkypeApiException
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pformat
import icu


class Bot:
    print('run')
    # fetch_id = "19:6fc079db903d48babebfa404621b1457@thread.skype"  # real
    fetch_id = "19:b098950b91e14699b97011b951b2b3aa@thread.skype"  # test
    # fetch_id = "19:87c093e9ad0443ba97b7f2d756ace3a5@thread.skype"
    error_id = "19:87c093e9ad0443ba97b7f2d756ace3a5@thread.skype"  # error
    option = 1000
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive',
    ]

    def __init__(self):
        self.fetch_admin = None
        self.fetch_group = None
        self.fetch_error = None
        self.sheet = None  # get data from sheet
        self.creds = None
        self.client = None
        self.list_order = None
        self.connect()

    def connect(self):
        try:
            # self.fetch_admin = Skype(
            #     'cuongdaovan262@gmail.com', 'developer26297@')
            # f = open('token.txt', 'a')
            # f.write(' ')
            self.fetch_admin = Skype('cuongdaovan262@gmail.com', 'developer26297@', tokenFile='token.txt')  # skype account
            self.fetch_group = self.fetch_admin.chats[self.fetch_id]
            self.fetch_error = self.fetch_admin.chats[self.error_id]
        except (SkypeApiException, SkypeAuthException):
            self.refreshToken()
            self.fetch_group = self.fetch_admin.chats[self.fetch_id]
            self.fetch_error = self.fetch_admin.chats[self.error_id]
    def sheet_update(self):
        """
        doc sheet lien tuc
        # need_to_do
        """
        while True:
            self.creds = ServiceAccountCredentials.from_json_keyfile_name(
                'fetch_sheet.json', self.scope)
            self.client = gspread.authorize(self.creds)
            s = self.client.open('fetch').sheet1
            self.sheet = s.get_all_records()
            self.list_dish()
            time.sleep(60)

    def all_key(self):
        """
        get all key of key-answer onto array
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
    
    def updateSheetOrder(self):
        while True:
            try:
                users = []
                for user in self.fetch_group.users:
                    users.append(str(user.name))
                sheet = self.client.open('Order cơm trưa').sheet1
                collator = icu.Collator.createInstance(icu.Locale('de_DE.UTF-8'))
                users.sort(key=collator.getSortKey)
                print(users)
                cell_list = sheet.range('B3:B'+str(len(users)+2))
                usr = deque(users)
                try:
                    for cell in cell_list:
                        if usr != []:
                            cell.value = str(usr.popleft())
                        else:
                            cell.value = ''
                except IndexError as e:
                    print('IndexError')
                # Update in batch
                sheet.update_cells(cell_list)
                time.sleep(1000)
            except Exception as e:
                self.sendMsg(self.fetch_group, msg='update order error: '+e, rich=False, typing=False)

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

    def list_dish(self):
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(
            'fetch_sheet.json', self.scope)
        self.client = gspread.authorize(self.creds)
        sheet_menu = self.client.open('menu')
        worksheet = sheet_menu.get_worksheet(0)
        rice = {}
        for dic in worksheet.get_all_records():
            if dic['key'].strip('\n') == 'rice food':
                rice.update({'rice food': dic['dish'].split('\n')})
            if dic['key'].strip('\n') == 'vegetables':
                dish = dic['dish'].split('\n')
                rice.update({'vegetables': dish})
            if dic['key'].strip('\n') == 'cost':
                rice.update({'cost': dic['dish'].split('\n')})
        worksheet1 = sheet_menu.get_worksheet(1)
        other = {}
        for dic in worksheet1.get_all_records():
            other.update({dic['other']: dic['cost']})
        result = {}
        result.update({'rice': rice})
        result.update({'other': other})
        self.list_order = result

    def order(self, msg, user=''):
        if msg.lower().find(',') >= 0:
            check = True
            sheet_order = self.client.open('Order cơm trưa').sheet1
            try:
                ms = msg[len('-order'):]
                order_rice = ms.split(',')
                order_rice = [str(i).lstrip(' ').rstrip(' ') for i in order_rice]
                v = self.list_order['rice']['vegetables']
                r = self.list_order['rice']['rice food']
                for i in order_rice[0].split(' '):
                    r = [s for s in r if i.lower() in s.lower()]
                for i in order_rice[1].split(' '):
                    v = [s for s in v if i.lower() in s.lower()]
                vegetable = random.choice(v)
                try:
                    cost = order_rice[2]
                    if cost in self.list_order['rice']['cost']:
                        pass
                    else:
                        cost = '30'
                except IndexError as e:
                    cost = '30'
                    print('cost error')
                rice = random.choice(r)
                if vegetable == '':
                    check = False
                if rice == '':
                    check = False
                if check == True:
                    cell = sheet_order.find(user)
                    sheet_order.update_acell('C'+str(cell.row), str(rice))
                    sheet_order.update_acell('D'+str(cell.row), str(vegetable))
                    sheet_order.update_acell('E'+str(cell.row), 'Cơm hộp')
                    sheet_order.update_acell('F'+str(cell.row), '')
                    sheet_order.update_acell('G'+str(cell.row), cost)
                    sheet_order.update_acell('H'+str(cell.row), 'Chưa Thanh toán')
            except ValueError as e:
                self.sendMsg(self.fetch_error, msg='order error: ' + e, rich=False, typing=True)
        else:
            try:
                result = self.list_order['other'].keys()
                ms = msg[len('-order'):]
                order = ms.split(' ')
                for i in order:
                    result = [s for s in result if i.lower() in s.lower()]
                other = random.choice(result)
                if other != '':
                    sheet_order = self.client.open('Order cơm trưa').sheet1
                    cell = sheet_order.find(user)
                    sheet_order.update_acell('C'+str(cell.row), '')
                    sheet_order.update_acell('D'+str(cell.row), '')
                    sheet_order.update_acell('E'+str(cell.row), '')
                    sheet_order.update_acell('F'+str(cell.row), other)
                    sheet_order.update_acell('G'+str(cell.row), self.list_order['other'][other])
                    sheet_order.update_acell('H'+str(cell.row), 'Chưa Thanh toán')
                else:
                    pass
            except Exception as e:
                self.sendMsg(self.fetch_error, msg='order other error: '+e, rich=False, typing=True)

    def refreshToken(self):
        f = open('token.txt','a')
        f.write(' ')
        self.fetch_admin = Skype('cuongdaovan262@gmail.com', 'developer26297@', tokenFile='token.txt')
        print(self.fetch_admin.conn.connected)

    def notify(self):
        """
        gui thong bao theo gio
        #need_to_do
        """
        while True:
            try:
                now = datetime.now()
                days = ["monday", "tuesday", "wednesday", "thursday", "friday"]
                if now.strftime("%A").lower() in days:
                    if self.check_time(datetime.now().strftime("%H:%M"), "09:09"):
                        msg = 'chúc <at id="*">all</at> một ngày làm việc vui vẻ và hiệu quả\n'
                        msg += "xem các option: -admin -help"
                        self.sendMsg(self.fetch_group, msg=msg,
                                     rich=True, typing=True)
                        time.sleep(100)
                    if self.check_time(datetime.now().strftime("%H:%M"), "10:00"):
                        for dic in self.sheet:
                            if dic['key'].find('đặt cơm') >= 0:
                                self.sendMsg(
                                    self.fetch_group, msg=dic['answer'], rich=True, typing=True)
                        time.sleep(100)
                    if self.check_time(datetime.now().strftime("%H:%M"), "10:45"):
                        self.update_order()
                        time.sleep(100)
                    if self.check_time(datetime.now().strftime("%H:%M"), "11:58"):
                        msg = "mọi người nghỉ tay đi ăn cơm đi ạ (sun)(sun)"
                        self.sendMsg(self.fetch_group, msg=msg,
                                     rich=False, typing=True)
                        time.sleep(100)
                    if self.check_time(datetime.now().strftime("%H:%M"), "17:00"):
                        worksheet = self.client.open('Order cơm trưa').sheet1
                        cell_list = worksheet.range('C3:H200')
                        for cell in cell_list:
                            cell.value = ''
                        worksheet.update_cells(cell_list)
                        time.sleep(60)
                    time.sleep(0.5)
            except (SkypeAuthException, SkypeApiException):
                self.refreshToken()
            except Exception as e:
                self.sendMsg(self.fetch_error, msg='error notify'+str(e),
                             rich=False, typing=True)

    def msg(self):
        while True:
            try:
                events = self.fetch_admin.getEvents()  # get tất cả các event trên skype
                if events != []:
                    event = events[0]
                    if event.type == "NewMessage":  # kiểm tra xe có message mới không
                        if hasattr(event, 'msg'):  # kiêm tra xem tồn tại msg variable không
                            msg = event.msg
                            content = msg.content
                            if msg.chat.id == self.fetch_id:
                                if content != None:
                                    # kiểm tra xem có mark @Fetch_admin không
                                    if content.startswith('-admin'):
                                        # cut take msg
                                        mess = content[len(
                                            '-admin '):]
                                        try:
                                            self.option = int(mess)
                                            print('int')
                                        except ValueError:
                                            print('not int')
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
                                            if mess.lower().find(dic['key'].lower().strip('\n')) >= 0:
                                                self.fetch_group.sendMsg(
                                                    dic['answer'], rich=True)
                                        if mess.find('-help') >= 0:
                                            msg = 'dưới đây là những option: \n'
                                            for key in self.all_key()[0:]:
                                                msg += str(self.all_key().index(key) +
                                                           1) + ". " + key + "\n"
                                            self.sendMsg(
                                                self.fetch_group, msg=msg, rich=True, typing=True)
                                        if mess.find('update order') >= 0:
                                            self.update_order()
                                        if mess.find('món ăn') >= 0:
                                            if self.list_order != None:
                                                v = 'vegetables: \n'
                                                for dish in self.list_order['rice']['vegetables']:
                                                    v += dish + '\n'
                                                r = 'rice: \n'
                                                for dish in self.list_order['rice']['rice food']:
                                                    r += dish + '\n'
                                                other = 'other: \n'
                                                for k in self.list_order['other']:
                                                    other += str(k) + ': ' + str(self.list_order['other'][k]) + '\n'
                                                self.sendMsg(
                                                    self.fetch_group, msg=v+r+other, rich=False, typing=True)
                                    if content.startswith('-order'):
                                        self.order(content, str(msg.user.name))
                    events = []
                time.sleep(0.5)
            except (SkypeAuthException, SkypeApiException):
                self.refreshToken()
            except Exception as e:
                self.sendMsg(self.fetch_error, msg='msg error: ' + str(e),
                             rich=False, typing=False)


bot = Bot()
t1 = threading.Thread(target=bot.sheet_update)
t2 = threading.Thread(target=bot.notify)
t3 = threading.Thread(target=bot.msg)
# t4 = threading.Thread(target=bot.updateSheetOrder)
t1.start()
t2.start()
t3.start()
# t4.start()
