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
from google.auth.transport.requests import Request


class Bot:
    print('run')
    fetch_id = "19:6fc079db903d48babebfa404621b1457@thread.skype"  # real
    # fetch_id = "19:b098950b91e14699b97011b951b2b3aa@thread.skype"  # test
    # fetch_id = "19:87c093e9ad0443ba97b7f2d756ace3a5@thread.skype"
    error_id = "19:87c093e9ad0443ba97b7f2d756ace3a5@thread.skype"  # error
    option = 1000
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive',
    ]

    def __init__(self):
        self.fetch_admin = Skype('cuongdaovan262@gmail.com', 'developer26297@', tokenFile='.tokens')
        self.fetch_group = None
        self.fetch_error = None
        self.sheet = None  # get data from sheet
        self.creds = None
        self.client = None
        self.list_order = None
        self.dishes = ''
        self.q_a = None
        self.winner = []
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(
                    'fetch_sheet.json', self.scope)
        self.client = gspread.authorize(self.creds)
        self.connect()

    def connect(self):
        try:
            self.fetch_admin.conn.readToken()
        except SkypeAuthException:
            self.fetch_admin.conn.setUserPwd('cuongdaovan262@gmail.com', 'developer26297@')
            self.fetch_admin.conn.getSkypeToken()
            self.fetch_admin.conn.writeToken()
        print(self.fetch_admin)
        self.fetch_group = self.fetch_admin.chats[self.fetch_id]
        self.fetch_error = self.fetch_admin.chats[self.error_id]

    def sheet_update(self):
        """
        doc sheet lien tuc
        # need_to_do
        """
        try:
            s = self.client.open('fetch').sheet1
            self.sheet = s.get_all_records()
            self.list_dish()
            self.question_answer()
        except gspread.exceptions.GSpreadException as e:
            self.client.auth.refresh(Request())                
        except Exception as e:
            self.sendMsg(self.fetch_error, msg='sheet update error: '+ str(e))
         
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
    
    def question_answer(self):
        sheet_qa = self.client.open('fetch')
        qa = sheet_qa.get_worksheet(1)
        self.q_a = qa.get_all_records()
    
    def updateSheetOrder(self):
        try:
            users = []
            for user in self.fetch_group.users:
                users.append(str(user.name))
            sheet = self.client.open('Order cơm trưa').sheet1
            collator = icu.Collator.createInstance(icu.Locale('de_DE.UTF-8'))
            users.sort(key=collator.getSortKey)
            print(users)
            cell_list = sheet.range('B2:B'+str(len(users)+100))
            usr = deque(users)
            try:
                for cell in cell_list:
                    cell.value = ''
                    if usr != deque([]):
                        cell.value = str(usr.popleft())
            except IndexError as e:
                print('IndexError')
            # Update in batch
            sheet.update_cells(cell_list)
        except Exception as e:
            self.sendMsg(self.fetch_error, msg='update order error: '+e, rich=False, typing=False)

    def update_order(self):
        sheet_order = self.client.open('Order cơm trưa').sheet1
        s = {}  # dictionary of other order
        p = {}  # dictionary of rice order
        for dic in sheet_order.get_all_records()[1:]:
            if dic['Món khác'] != '':
                if dic['Món khác'] in s:
                    s[dic['Món khác']] += dic['Tên'] + ', '
                else:
                    s.update({dic['Món khác']: dic['Tên'] + ', '})
            if dic['Cơm suất'] != '':
                if dic['Cơm suất'] in p:
                    p[dic['Cơm suất']] += dic['Tên'] + ', '
                else:
                    p.update({dic['Cơm suất']: dic['Tên'] + ', '})
        msg = '<at id="*">all</at> fetch_admin xác nhận: \n'
        for key in s:
            msg += str(key) + ': ' + str(s[key]) + '\n'
        for key in p:
            msg += str(key) + ': ' + str(p[key]) + '\n'
        for dic in self.sheet:
            if dic['key'].find('link order') >= 0:
                msg += dic['answer']
        self.sendMsg(self.fetch_group, msg=msg, rich=True, typing=False)

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

    def order(self, msg, user):
        error = 'order sai\nmuốn order cơm: -order món thịt, món rau\nmuốn order món khác: -order tên món'
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
                vegetable = ''
                if v != []:
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
                rice = ''
                if r != []:
                    rice = random.choice(r)
                if vegetable == '':
                    check = False
                if rice == '':
                    check = False
                if check == True:
                    cell = sheet_order.find(str(user.name))
                    sheet_order.update_acell('C'+str(cell.row), str(rice))
                    # sheet_order.update_acell('D'+str(cell.row), str(vegetable))
                    # sheet_order.update_acell('E'+str(cell.row), 'Cơm hộp')
                    sheet_order.update_acell('D'+str(cell.row), '')
                    # sheet_order.update_acell('G'+str(cell.row), cost)
                    sheet_order.update_acell('E'+str(cell.row), 'Chưa thanh toán')
                else:
                    ch = self.fetch_admin.contacts[user.id].chat
                    ch.sendMsg(error)
            except ValueError as e:
                self.sendMsg(self.fetch_error, msg='order error: ' + e, rich=False, typing=True)
        else:
            try:
                result = self.list_order['other'].keys()
                note = ''
                ms = msg[len('-order'):]
                if ms.find('|') >= 0:
                    #do something
                    note = ms[ms.index('|') + 1: len(ms)]
                    ms = ms[0 : ms.index('|')]
                    print(ms)
                order = ms.split(' ')           
                for i in order:
                    result = [s for s in result if i.lower() in s.lower()]
                other = ''
                if result != []:
                    other = random.choice(result)
                print(other)
                print(note)
                if other != '':
                    sheet_order = self.client.open('Order cơm trưa').sheet1
                    cell = sheet_order.find(str(user.name))
                    sheet_order.update_acell('C'+str(cell.row), '')
                    sheet_order.update_acell('D'+str(cell.row), '')
                    sheet_order.update_acell('E'+str(cell.row), '')
                    sheet_order.update_acell('F'+str(cell.row), other)
                    sheet_order.update_acell('G'+str(cell.row), self.list_order['other'][other])
                    sheet_order.update_acell('H'+str(cell.row), 'Chưa thanh toán')
                    sheet_order.update_acell('I'+str(cell.row), note)
                    self.sendMsg(self.fetch_group,msg='(like)(like)(like)(like)(like)', rich=False, typing=True)
                else:
                    ch = self.fetch_admin.contacts[user.id].chat
                    ch.sendMsg(error)
            except Exception as e:
                self.sendMsg(self.fetch_error, msg='order other error: '+e, rich=False, typing=True)

    def refreshToken(self):
        self.fetch_admin.conn.setUserPwd('cuongdaovan262@gmail.com', 'developer26297@')
        self.fetch_admin.getSkypeToken()
        self.fetch_admin.conn.writeToken()
        self.fetch_group = self.fetch_admin.chats[self.fetch_id]
        self.fetch_error = self.fetch_admin.chats[self.error_id]

    def notify(self):
        """
        gui thong bao theo gio
        #need_to_do
        """
        # self.updateSheetOrder()
        while True:
            try:
                now = datetime.now()
                days = ["monday", "tuesday", "wednesday", "thursday", "friday"]
                if now.strftime("%A").lower() in days:
                    if self.check_time(datetime.now().strftime("%H:%M"), "09:00"):
                        msg = 'chúc <at id="*">all</at> một ngày làm việc vui vẻ và hiệu quả\n'
                        msg += "xem các option: -admin -help"
                        self.sendMsg(self.fetch_group, msg=msg,
                                     rich=True, typing=True)
                        time.sleep(60)
                    if self.check_time(datetime.now().strftime("%H:%M"), "09:15"):
                        for dic in self.sheet:
                            if dic['key'].find('đặt cơm') >= 0:
                                self.sendMsg(
                                    self.fetch_group, msg=dic['answer'], rich=True, typing=True)
                        time.sleep(60)
                    if self.check_time(datetime.now().strftime("%H:%M"), "10:45"):
                        self.update_order()
                        time.sleep(60)
                    if self.check_time(datetime.now().strftime("%H:%M"), "09:45"):
                        self.update_order()
                        time.sleep(60)
                    if self.check_time(datetime.now().strftime("%H:%M"), "11:58"):
                        msg = "mọi người nghỉ tay đi ăn cơm đi ạ (sun)(sun)"
                        self.sendMsg(self.fetch_group, msg=msg,
                                     rich=False, typing=True)
                        time.sleep(60)
                    if self.check_time(datetime.now().strftime("%H:%M"), "17:00"):
                        worksheet = self.client.open('Order cơm trưa').sheet1
                        cell_list = worksheet.range('C3:F200')
                        for cell in cell_list:
                            cell.value = ''
                        worksheet.update_cells(cell_list)
                        self.updateSheetOrder()
                        time.sleep(60)
                    if self.check_time(datetime.now().strftime("%H:%M"), "09:10"):
                        worksheet = self.client.open('Order cơm trưa').sheet1
                        cell_list = worksheet.range('C3:F200')
                        for cell in cell_list:
                            cell.value = ''
                        worksheet.update_cells(cell_list)
                        self.updateSheetOrder()
                        time.sleep(60)
                    # if self.check_time(datetime.now().strftime("%H:%M"), "16:32"):
                    #     msg = 'Hi <at id="*">all</at>, Xin thông báo, fetch_admin đưa ra mỗi ngày một câu hỏi để cho các ace giải trí\n'
                    #     msg += "Thể lệ cuộc thi là tìm ra ai là người trả lời câu hỏi đúng nhất chính xác nhất và nhanh nhất " 
                    #     msg += "và người chiến thắng sẽ được ra câu hỏi cho ngày tiếp theo\n"
                    #     msg += "mỗi người chỉ được trả lời 1 lần"
                    #     msg += "fetch_admin sẽ công bố vào lúc 17:15\n"
                    #     msg += "Câu hỏi ngày hôm nay là:\n"
                    #     if self.q_a != None:
                    #         msg += self.q_a[0]['question']
                    #         print(self.q_a[0]['question'])
                    #         self.sendMsg(self.fetch_group, msg=msg, rich=True)
                    #     self.winner = []
                    #     time.sleep(60)
                    # if self.check_time(datetime.now().strftime("%H:%M"), "17:15"):
                    #     msg = "fetch_admin xin công bố người chiến thắng hôm nay là: "
                    #     if self.winner != []:
                    #         msg += self.winner[0]
                    #     else:
                    #         msg += "không có ai chiến thắng" 
                    #     self.sendMsg(self.fetch_group, msg=msg)
                    #     time.sleep(60)
                time.sleep(5)
            except SkypeAuthException as e:
                self.refreshToken()
            except SkypeApiException as e:
                self.sendMsg(self.fetch_error, msg='error notify'+str(e), rich=False, typing=False)
            except Exception as e:
                print('error')
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
                                    print(content)
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
                                        self.order(content, msg.user)
                                    # if self.q_a != None:
                                    #     if content.lower().startswith(str(self.q_a[0]['answer']).lower()):
                                    #         self.winner.append(str(msg.user.name))
                                    #         print('chính xác')
                    events = []
                time.sleep(1)
            except SkypeAuthException as e:
                self.refreshToken()
            except Exception as e:
                self.sendMsg(self.fetch_error, msg='msg error: ' + str(e),
                             rich=False, typing=False)

bot = Bot()
bot.sheet_update()
# t1 = threading.Thread(target=bot.sheet_update)
t2 = threading.Thread(target=bot.msg)
t3 = threading.Thread(target=bot.notify)
# t1.start()
t2.start()
t3.start()