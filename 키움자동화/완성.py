import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtMultimedia import QSound , QMediaPlayer , QMediaContent
from PyQt5.QtCore import QUrl

from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtGui import QClipboard  # 복사

from collections import deque
import numpy as np
from datetime import datetime
import time

import atexit
import csv

now = datetime.now()


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setGeometry(300, 300, 800, 600)
        self.setWindowTitle("Kiwoom 조건식 완성 및 알림")

        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)
        self.ocx.OnReceiveConditionVer.connect(self._handler_condition_load)
        self.ocx.OnReceiveRealCondition.connect(self._handler_real_condition)
        self.ocx.OnReceiveTrCondition.connect(self._handler_tr_condition)
        self.ocx.OnReceiveTrData.connect(self._handler_tr_data)
        self.ocx.OnReceiveRealData.connect(self._handler_real_data)
        self.CommConnect()

        btn1 = QPushButton("Condition Down")
        btn2 = QPushButton("Condition List")
        btn3 = QPushButton("Condition Send")

        self.listWidget = QListWidget()
        self.price_data = {}
        self.rsiListWidget = QListWidget()
        self.rsiList30under  =  QListWidget()
        self.rsiList30today = QListWidget()

        # 알람 표시
        self.last_shown_message = None  # 마지막으로 표시된 메시지를 기록
        self.message_timestamp = 0      # 마지막 메시지의 타임스탬프를 기록
        self.alert_timer = QTimer(self)  # 타이머 설정
        self.alert_timer.setSingleShot(True)  # 타이머가 한 번만 실행되도록 설정
        self.last_saved_time = 0 
        
 
        # Layouts
        main_layout = QVBoxLayout()
        button_layout = QHBoxLayout()
        list_layout = QHBoxLayout()


        # Add buttons to button layout
        button_layout.addWidget(btn1)
        button_layout.addWidget(btn2)
        button_layout.addWidget(btn3)

         # Add lists to list layout with titles
        vertical_layout_left = QVBoxLayout()
        vertical_layout_left.addWidget(QLabel("Stock Codes"))
        vertical_layout_left.addWidget(self.listWidget)

        vertical_layout_right = QVBoxLayout()
        vertical_layout_right.addWidget(QLabel("RSI Values"))
        vertical_layout_right.addWidget(self.rsiListWidget)

        vertical_layout_right1 = QVBoxLayout()
        vertical_layout_right1.addWidget(QLabel("RSI stock and time"))
        vertical_layout_right1.addWidget(self.rsiList30under)

        vertical_layout_right2 = QVBoxLayout()
        vertical_layout_right2.addWidget(QLabel("today RSI"))
        vertical_layout_right2.addWidget(self.rsiList30today)

        list_layout.addLayout(vertical_layout_left)
        list_layout.addLayout(vertical_layout_right)
        list_layout.addLayout(vertical_layout_right1)
        list_layout.addLayout(vertical_layout_right2)

        # Add sub-layouts to main layout
        main_layout.addLayout(button_layout)
        main_layout.addLayout(list_layout)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        # event
        btn1.clicked.connect(self.GetConditionLoad)
        btn2.clicked.connect(self.GetConditionNameList)
        btn3.clicked.connect(self.send_condition)
        
        self.media_player = QMediaPlayer()

        # Queue for managing stock price requests
        self.stock_queue = deque()
        self.price_data = {}  # Dictionary to store price data for each stock
        self.price_data120 = {}
        self.timer = QTimer()
        self.timer.timeout.connect(self.process_queue)
        self.timer.start(200)  # Process queue every 200ms

        #메세지
        self.msg_boxes = []

        # 시스템 종료 했을 떄 기록
        self.today_ris = []

        # Connect the itemClicked signal to the slot
        self.rsiList30under.itemClicked.connect(self.display_rsi_history)

        # 프로그램 종료 시 save_data_to_csv 함수 실행 등록
        atexit.register(self.save_data)



    def CommConnect(self):
        self.ocx.dynamicCall("CommConnect()")

    def _handler_login(self, err_code):  # 로그인
        print("handler login", err_code)

    def _handler_condition_load(self, ret, msg): # 조건 로드
        print("handler condition load", ret, msg)

    def _handler_real_condition(self, code, type, cond_name, cond_index): # 실시간
        print(cond_name, code, type)

    def _handler_tr_condition(self, sCrNo, strCodeList, strConditionName, nIndex, nNext):
        print(strCodeList, strConditionName)
        self.update_list_widget(strCodeList)

    def _handler_tr_data(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext, data_len, err_code, msg1, msg2):
        if sRQName == "opt10001_req":
            current_price = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "현재가").strip()
            stock_code = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "종목코드").strip()
            self.price_data[stock_code]= [[float(current_price) ,0]]
            self.price_data120[stock_code]= [[float(current_price) ,0]]


    def _handler_real_data(self, code, real_type, real_data):
        if real_type == "주식체결":
            tick_price = self.ocx.dynamicCall("GetCommRealData(QString, int)", code, 10).strip()
            volume = self.ocx.dynamicCall("GetCommRealData(QString, int)",code , 15)  # 거래량

            self.price_data[code].append([float(tick_price), volume])

            if len(self.price_data[code]) >= 120:
                r = self.price_data[code].pop()
                self.price_data120[code].append(r)
                self.price_data[code] = deque(maxlen=121)
 
            if len(self.price_data120[code]) >= 14:
                rsi = self.calculate_rsi_one(code)
                self.update_rsi_list_widget(code, rsi)
                self.price_data120[code].popleft()

    def calculate_rsi_one(self, code):
        prices = [int(price) for price, volume in self.price_data120[code]]
        if len(prices) < 14:
            return 100
        
        deltas = [prices[i] - prices[i-1] for i in range(1, len(prices))]
        print(deltas)

        gains = [delta for delta in deltas if delta > 0]
        losses = [-delta for delta in deltas if delta < 0]

        average_gain = sum(gains) / 14 if len(gains) > 0 else 0
        average_loss = sum(losses) / 14 if len(losses) > 0 else 0

        # RSI 계산
        for i in range(14, len(deltas)):
            delta = deltas[i]
            gain = max(delta, 0)
            loss = -min(delta, 0)
            
            average_gain = (average_gain * 13 + gain) / 14
            average_loss = (average_loss * 13 + loss) / 14

        if average_loss == 0:
            rsi = 100
        else:
            rs = average_gain / average_loss
            rsi = 100 - (100 / (1 + rs))

        return rsi

    def calculate_rsi(self, prices):
        deltas = np.diff(prices)
        seed = deltas[:13]  # First 13
        up = seed[seed >= 0].sum() / 14
        down = -seed[seed < 0].sum() / 14
        rs = up / down
        rsi = 100 - (100 / (1 + rs))
        return rsi

    def update_list_widget(self, strCodeList):  # 업데이트 리스트 
        self.listWidget.clear()
        self.rsiListWidget.clear()
        self.rsiList30under.clear()
        codes = strCodeList.split(';')
        print("개수" , len(codes))
        for code in codes:
            if code:
                item = QListWidgetItem(code)
                self.listWidget.addItem(item)  # ui 목록에 추가하는 법
                self.request_real_time_data(code)
                self.price_data[code] = deque(maxlen=120)
                self.price_data120[code] = deque(maxlen=120)

    def update_rsi_list_widget(self, stock_code, rsi):
        if 0 < rsi < 30:
            background_color = Qt.red 
            # Play the sound
            url = QUrl.fromLocalFile("alert.mp3")  # Replace with the path to your sound file
            self.media_player.setMedia(QMediaContent(url))
            self.media_player.play()

            # Show popup message 
            self.handle_alert(stock_code, rsi)

            for i in range(self.rsiList30under.count()):
                item = self.rsiList30under.item(i)
                if stock_code in item.text():
                    #current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    item.setText(f"{stock_code}")
                    item.setBackground(background_color)
                    return

            #current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            item = QListWidgetItem(f"{stock_code}")
            self.rsiList30under.addItem(item)

        else:
            background_color = Qt.magenta

        for i in range(self.rsiListWidget.count()):
            item = self.rsiListWidget.item(i)
            if stock_code in item.text():
                item.setText(f"{stock_code}: {rsi:.2f}")
                item.setBackground(background_color)
                return

        item = QListWidgetItem(f"{stock_code}: {rsi:.2f}")
        item.setBackground(background_color)
        self.rsiListWidget.addItem(item)

    def handle_alert(self, stock_code, rsi):
        current_time = time.time()  # 현재 시간을 초 단위로 얻기
        time_diff = current_time - self.last_saved_time
        print(self.last_saved_time)
        # 마지막 저장이 3초 이내이면 데이터를 저장하지 않음
        if time_diff >= 3:
            # 현재 시간 기록
            self.last_saved_time = current_time

            current_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.today_ris.append([current_time_str, stock_code, f"{rsi:.2f}"])
            self.save_data()  # Save data to CSV whenever an alert is triggered

            # Show popup message
            self.show_popup( f"The RSI for {stock_code} , Current RSI: {rsi:.2f}" , f"{stock_code}",)


    def display_rsi_history(self, item):
        self.rsiList30today.clear()
        stock_code = item.text()
        history = [entry for entry in self.today_ris if entry[1] == stock_code]
        print("가져옴 ",history )
        for entry in history:
            self.rsiList30today.addItem(f"{entry[0]} : {entry[1]} : {entry[2]}")

    def show_popup(self, title, message):

        # 팝업 창 표시
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Warning)
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.setModal(True)  # 비모달로 설정

        # 팝업이 닫힐 때 리스트에서 제거
        msg_box.finished.connect(lambda: self.msg_boxes.remove(msg_box))
        msg_box.buttonClicked.connect(lambda button: self.copy_to_clipboard(button, message, msg_box))
        # 팝업을 표시하고 리스트에 저장
        msg_box.show()
        self.msg_boxes.append(msg_box)

        #msg_box.exec_()
        #msg_box.exec_()
        #일정 시간 후에 팝업 창을 닫도록 타이머 설정
        #QTimer.singleShot(5000, msg_box.close)  # 5000ms = 5초 동안 팝업 유지

    def copy_to_clipboard(self, button, message, msg_box):
        if button.text() == "OK":
            # 클립보드에 메시지 복사
            clipboard = QApplication.clipboard()
            clipboard.setText(message)

    def request_stock_price(self, code):
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.ocx.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10001_req", "opt10001", 0, "2000")
    
    def request_real_time_data(self, code):
        self.ocx.dynamicCall("SetRealReg(QString, QString, QString, QString)", "0101", code, "20", "1")

    def process_queue(self):
        if self.stock_queue and not self.ocx.dynamicCall("GetCommDataEx()"):  # Ensure no current request is being processed
            code = self.stock_queue.popleft()
            self.request_stock_price(code)

    def GetConditionLoad(self):
        self.ocx.dynamicCall("GetConditionLoad()")
        # Set the media content
        url = QUrl.fromLocalFile("alert.mp3")  # Replace with the path to your sound file
        self.media_player.setMedia(QMediaContent(url))
        
        # Play the sound
        self.media_player.play()

    def GetConditionNameList(self):
        data = self.ocx.dynamicCall("GetConditionNameList()")
        conditions = data.split(";")[:-1]
        for condition in conditions:
            index, name = condition.split('^')
            print(index, name)
        self.index = index
        self.name = name

    def SendCondition(self, screen, cond_name, cond_index, search):
        ret = self.ocx.dynamicCall("SendCondition(QString, QString, int, int)", screen, cond_name, cond_index, search)
        print(ret)

    def send_condition(self):
        self.SendCondition("100", self.name, self.index,1)

    def save_data(self):
        print(self.today_ris)
        save_data_to_csv(self.today_ris)

def save_data_to_csv(data, filename="price_data.csv"):
    with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ["시간" , "종목" , "RSI"]
        writer = csv.writer(csvfile)
        writer.writerow(fieldnames)
        for i in data:
            writer.writerow(i)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()
