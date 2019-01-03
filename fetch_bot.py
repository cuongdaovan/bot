from datetime import time
from datetime import datetime
import time
from threading import Thread
import threading

from skpy import Skype, SkypeChats
from skpy import SkypeContact, SkypeAuthException
import gspread
from oauth2client.service_account import ServiceAccountCredentials


class Bot:
    print('run')
    fetch_admin = None
    # fetch_id = "19:6fc079db903d48babebfa404621b1457@thread.skype"  # real
    fetch_id = "19:b098950b91e14699b97011b951b2b3aa@thread.skype"  # test
    error_id = "19:87c093e9ad0443ba97b7f2d756ace3a5@thread.skype"  # error
    fetch_group = None
    fetch_error = None
    option = 1000
    username = ''  # lay ra ten user da gui
    sheet = None  # get data from sheet
    all_key = ''  # get all key from spreadsheet
    spams = {}
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive',
    ]
    creds = None
    client = None
    list_order = None

    def connect(self):
        try:
            self.fetch_admin = Skype(
                tokenFile='token.txt',
                connect=True
            )  # skype account
            # group chat
            self.fetch_group = self.fetch_admin.chats[self.fetch_id]
            # group error
            self.fetch_error = self.fetch_admin.chats[self.error_id]
        except Exception:
            self.sendMsg(self.fetch_error, msg='connect error',
                         rich=False, typing=True)

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

    def order(self, msg, user=''):
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(
            'fetch_sheet.json', self.scope)
        self.client = gspread.authorize(self.creds)
        sheet_menu = self.client.open('menu')
        worksheet = sheet_menu.get_worksheet(0)
        sheet_order = self.client.open('order').sheet1
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
        order_rice = msg.split(' ')
        print(order_rice)
        try:
            if order_rice[0] == 'rice':
                order = 'cơm hộp: '
                cost = 0
                right = True
                if int(order_rice[1]) < len(result['rice']['vegetables']) and int(order_rice[1]) > 0:
                    order += result['rice']['vegetables'][int(
                        order_rice[1])] + " + "
                else:
                    right = False
                if int(order_rice[2]) < len(result['rice']['rice food']) and int(order_rice[2]) > 0:
                    order += result['rice']['rice food'][int(order_rice[2])]
                else:
                    right = False
                if order_rice[3] in result['rice']['cost']:
                    cost = int(order_rice[3])
                else:
                    right = True
                if right == True:
                    cell = sheet_order.find(user)
                    sheet_order.update_acell('C'+str(cell.row), order)
                    sheet_order.update_acell('D'+str(cell.row), '')
                    sheet_order.update_acell('E'+str(cell.row), cost)
                    sheet_order.update_acell(
                        'F'+str(cell.row), 'chưa thanh toán')
            if order_rice[0] == 'other':
                if int(order_rice[1]) < len(result['other']):
                    cell = sheet_order.find(user)
                    sheet_order.update_acell('C'+str(cell.row), '')
                    sheet_order.update_acell(
                        'D'+str(cell.row), list(result['other'].keys())[int(order_rice[1])])
                    sheet_order.update_acell(
                        'E'+str(cell.row), list(result['other'].values())[int(order_rice[1])])
                    sheet_order.update_acell(
                        'F'+str(cell.row), 'chưa thanh toán')
        except Exception:
            print('error value')

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
                    if self.check_time(datetime.now().strftime("%H:%M:%S"), "09:00:01"):
                        print('true')
                        msg = 'chúc <at id="*">all</at> một ngày làm việc vui vẻ và hiệu quả\n'
                        msg += "xem các option: @Fetch_admin --help"
                        self.sendMsg(self.fetch_group, msg=msg,
                                     rich=True, typing=True)
                        time.sleep(1)
                    if self.check_time(datetime.now().strftime("%H:%M:%S"), "10:00:05"):
                        for dic in self.sheet:
                            if dic['key'].find('order') >= 0:
                                self.sendMsg(
                                    self.fetch_group, msg=dic['answer'], rich=True, typing=True)
                        time.sleep(1)
                    if self.check_time(datetime.now().strftime("%H:%M:%S"), "10:45:01"):
                        self.update_order()
                        time.sleep(1)
                    if self.check_time(datetime.now().strftime("%H:%M:%S"), "11:58:01"):
                        msg = "mọi người nghỉ tay đi ăn cơm đi ạ (sun)(sun)"
                        self.sendMsg(self.fetch_group, msg=msg,
                                     rich=False, typing=True)
                        time.sleep(1)
                    if self.check_time(datetime.now().strftime("%H:%M"), "17:00"):
                        worksheet = self.client.open('Order cơm trưa').sheet1
                        cell_list = worksheet.range('C3:H200')
                        for cell in cell_list:
                            cell.value = ''
                        worksheet.update_cells(cell_list)
                        time.sleep(60)
                    time.sleep(0.5)
            except Exception:
                self.sendMsg(self.fetch_error, msg='error notify',
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
                            print(type(msg.user.name.__str__()))
                            if msg.chat.id == self.fetch_id:
                                if content != None:
                                    # kiểm tra xem có mark @Fetch_admin không
                                    if content.startswith('<at id="8:live:cuongdaovan262">Fetch_admin</at>'):
                                        # cut take msg
                                        mess = content[len(
                                            '<at id="8:live:cuongdaovan262">Fetch_admin</at> '):]
                                        print(mess)
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
                                            if mess.lower().find(dic['key'].lower()) >= 0:
                                                self.fetch_group.sendMsg(
                                                    dic['answer'], rich=True)
                                        if mess.find('--help') >= 0:
                                            msg = 'dưới đây là những option: \n'
                                            for key in self.all_key()[0:]:
                                                msg += str(self.all_key().index(key) +
                                                           1) + ". " + key + "\n"
                                            self.sendMsg(
                                                self.fetch_group, msg=msg, rich=True, typing=True)
                                        if mess.find('update order') >= 0:
                                            self.update_order()
                                        if mess.find('rice') or mess.find('other'):
                                            # self.order(
                                            #     mess, msg.user.name.__str__())
                                            pass
                                        if mess.find('món ăn'):
                                            v = 'vegetables: \n'
                                            for dish in self.list_order['rice']['vegetables']:
                                                v += str(self.list_order['rice']['vegetables'].index(
                                                    dish)) + ".  " + dish + '\n'
                                            print(v)
                                            r = 'rice: \n'
                                            for dish in self.list_order['rice']['rice food']:
                                                r += str(self.list_order['rice']['rice food'].index(
                                                    dish)) + ".  " + dish + '\n'
                                            print(r)
                                            self.sendMsg(
                                                self.fetch_group, msg=v+r, rich=False, typing=True)

                    events = []
                time.sleep(0.5)
            except Exception as e:
                self.sendMsg(self.fetch_error, msg=e,
                             rich=False, typing=False)


bot = Bot()
bot.connect()
t1 = threading.Thread(target=bot.sheet_update)
t2 = threading.Thread(target=bot.notify)
t3 = threading.Thread(target=bot.msg)

t1.start()
t2.start()
t3.start()
