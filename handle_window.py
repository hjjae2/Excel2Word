from PyQt5 import QtWidgets
from PyQt5.QtGui import *
from PyQt5.QtCore import Qt, QBasicTimer
from PyQt5.QtWidgets import *

import handle_word

'''
Reference
1. Python-docx 사용해보기, https://nittaku.tistory.com/253
'''

class App(QWidget):
    def __init__(self):

        super().__init__()

        self.setWindowTitle("V1")
        self.resize(860, 640)   # for size
        # _palette = QPalette() # for background-image
        # _palette.setBrush(10, QBrush(QImage('./image/back.jpg').scaled(800, 600)))
        # self.setPalette(_palette)
        self.setStyleSheet("background: #708090;")   # for style-sheet
        self.setStyleSheet("background-image: linear-gradient(to bottom, #708090, #4d627e, #2f446b, #172656, #06063f);")
        _palette = QPalette()
        gradient = QLinearGradient(0, 0, 800, 600)    # start point, final point
        gradient.setColorAt(0.0, QColor(110, 130, 145))
        gradient.setColorAt(1.0, QColor(55, 65, 75))
        _palette.setBrush(QPalette.Window, QBrush(gradient))
        self.setPalette(_palette)
        self.file_name = "Files..."
        self.initUI()

    def initUI(self):

        ''' Widgets for sender '''
        self.setAcceptDrops(True)   # for drag & drop
        self.sender_number = QLineEdit()
        self.sender_number.setPlaceholderText("123-45-67890")
        self.sender_company = QLineEdit()
        self.sender_company.setPlaceholderText("상호")
        self.sender_name = QLineEdit()
        self.sender_name.setPlaceholderText("대표자명")
        self.sender_address = QLineEdit(self)
        self.sender_address.setPlaceholderText("사업장 주소")
        self.sender_call = QLineEdit(self)
        self.sender_call.setPlaceholderText("123)456-7890")
        self.sender_fax = QLineEdit(self)
        self.sender_fax.setPlaceholderText("123)456-7890")
        self.sender_business = QLineEdit(self)
        self.sender_business.setPlaceholderText("업태")
        self.sender_category = QLineEdit(self)
        self.sender_category.setPlaceholderText("종목")
        self.receiver = QLineEdit(self)
        self.receiver.setPlaceholderText("받는 회사명")
        self.register_date = QLineEdit(self)
        self.register_date.setPlaceholderText("2019-12-31")
        # self.register_date_year = QComboBox()
        # self.register_date_month = QComboBox()
        # self.register_date_day = QComboBox()
        # for year in range(2019,2023):
        #     self.register_date_year.addItem(str(year))
        # for month in range(1,13):
        #     self.register_date_month.addItem(str(month))
        # for day in range(1,32):
        #     self.register_date_day.addItem(str(day))

        self.sender_number_label = QLabel("등록번호", self)
        self.sender_company_label = QLabel("상호", self)
        self.sender_name_label = QLabel("대표자", self)
        self.sender_address_label = QLabel("사업장주소", self)
        self.sender_call_label = QLabel("전화", self)
        self.sender_fax_label = QLabel("팩스", self)
        self.sender_business_label = QLabel("업태", self)
        self.sender_category_label = QLabel("종목", self)
        self.receiver_label = QLabel("받는회사", self)
        self.register_date_label = QLabel("등록날짜", self)
        # self.register_date_year_label = QLabel("년", self)
        # self.register_date_month_label = QLabel("월", self)
        # self.register_date_day_label = QLabel("일", self)

        # self.dnd_label = QLabel('Drag n Drop, Here!', self)
        self.dnd_label = FileEdit(self)
        self.dnd_label.setAlignment(Qt.AlignCenter)
        self.dnd_label.setStyleSheet("background:white url('./image/here2.jpg') no-repeat center center; border:2px solid black; border-radius: 30px;")

        self.add_button = QPushButton(self)
        self.add_button.setIcon(QIcon('./image/plus.png'))
        self.add_button.clicked.connect(self.filePushButtonClicked)
        self.add_button.setStyleSheet("margin:0px; ")

        self.sub_button = QPushButton(self)
        self.sub_button.setIcon(QIcon('./image/minus.png'))
        self.sub_button.clicked.connect(self.fileDownButtonClicked)
        self.sub_button.setStyleSheet("margin:0px")

        self.go_button = QPushButton(self)
        self.go_button.setIcon(QIcon('./image/check.png'))
        self.go_button.clicked.connect(self.fileGoButtonClicked)
        self.go_button.setStyleSheet("margin:0px 50px 0px 50px")

        self.files_label = QLabel('Files...', self)
        self.files_label.setAlignment(Qt.AlignCenter)
        self.files_label.setStyleSheet("font:bold 15px 나눔고딕,NanumGothic,돋움,Dotum")

        self.dnd_label.setTargetWidget(self.files_label)


        self.status_bar = QProgressBar(self)
        self.status_bar.setAlignment(Qt.AlignCenter)
        self.timer = QBasicTimer()
        self.step = 0

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.sub_button)
        button_layout.setContentsMargins(50,0,50,0)

        main_layout = QVBoxLayout()
        sub_layout = QGridLayout()
        sub_layout.setContentsMargins(20,5,20,5)
        sub_layout.addWidget(self.sender_number_label, 0, 0); sub_layout.addWidget(self.sender_number, 0, 1); sub_layout.addWidget(self.register_date_label, 0, 2); sub_layout.addWidget(self.register_date, 0, 3)
        sub_layout.addWidget(self.sender_company_label, 1, 0); sub_layout.addWidget(self.sender_company, 1, 1); sub_layout.addWidget(self.receiver_label, 1, 2); sub_layout.addWidget(self.receiver, 1, 3);
        sub_layout.addWidget(self.sender_name_label, 2, 0); sub_layout.addWidget(self.sender_name, 2, 1)
        sub_layout.addWidget(self.sender_address_label, 3, 0); sub_layout.addWidget(self.sender_address, 3, 1)
        sub_layout.addWidget(self.sender_call_label, 4, 0); sub_layout.addWidget(self.sender_call, 4, 1); sub_layout.addWidget(self.sender_fax_label, 4, 2); sub_layout.addWidget(self.sender_fax, 4, 3)
        sub_layout.addWidget(self.sender_business_label, 5, 0); sub_layout.addWidget(self.sender_business, 5, 1); sub_layout.addWidget(self.sender_category_label, 5, 2); sub_layout.addWidget(self.sender_category, 5, 3)

        main_layout.addLayout(sub_layout)  # layout
        main_layout.addWidget(self.dnd_label)
        main_layout.addLayout(button_layout)    # layout
        main_layout.addWidget(self.go_button)
        main_layout.addWidget(self.files_label)
        main_layout.addWidget(self.status_bar)
        main_layout.setContentsMargins(10,10,10,10)

        self.setLayout(main_layout)
        self.setWindowIcon(QIcon("./image/web.png"))
        self.show()

    def filePushButtonClicked(self):
        self.file_name = QFileDialog.getOpenFileName(self)[0]
        self.files_label.setText(self.file_name)

    def fileDownButtonClicked(self):
        self.file_name = "Files..."
        self.files_label.setText(self.file_name)

    def fileGoButtonClicked(self):
        print("Try to create word file...")
        if self.files_label.text() == "Files...":
            reply = QMessageBox.information(self, '확인메세지', 'File을 올려주세요.', QMessageBox.Yes)
            return
        else:
            self.file_name = self.files_label.text()
        try:
            file_name = self.file_name
            number = self.sender_number.text()
            company = self.sender_company.text()
            name = self.sender_name.text()
            address = self.sender_address.text()
            call = self.sender_call.text()
            fax = self.sender_fax.text()
            business = self.sender_business.text()
            category = self.sender_category.text()
            receiver = self.receiver.text()
            register_date = self.register_date.text()
        except Exception as e:
            print(e)

        if number == "":
            reply = QMessageBox.information(self, '확인메세지', '등록번호를 입력해주세요.', QMessageBox.Yes); print("Failed creating word file: Please input data to boxes.")
            return
        if company == "":
            reply = QMessageBox.information(self, '확인메세지', '상호를 입력해주세요.', QMessageBox.Yes); print("Failed creating word file: Please input data to boxes.")
            return
        if name == "":
            reply = QMessageBox.information(self, '확인메세지', '대표자를 입력해주세요.', QMessageBox.Yes); print("Failed creating word file: Please input data to boxes.")
            return
        if address =="":
            reply = QMessageBox.information(self, '확인메세지', '주소를 입력해주세요.', QMessageBox.Yes); print("Failed creating word file: Please input data to boxes.")
            return
        if call =="":
            reply = QMessageBox.information(self, '확인메세지', '전화를 입력해주세요.', QMessageBox.Yes); print("Failed creating word file: Please input data to boxes.")
            return
        if fax == "":
            reply = QMessageBox.information(self, '확인메세지', '팩스를 입력해주세요.', QMessageBox.Yes); print("Failed creating word file: Please input data to boxes.")
            return
        if business == "":
            reply = QMessageBox.information(self, '확인메세지', '업태를 입력해주세요.', QMessageBox.Yes); print("Failed creating word file: Please input data to boxes.")
            return
        if category == "":
            reply = QMessageBox.information(self, '확인메세지', '종목을 입력해주세요.', QMessageBox.Yes); print("Failed creating word file: Please input data to boxes.")
            return
        if register_date =="":
            reply = QMessageBox.information(self, '확인메세지', '등록날짜를 입력해주세요.', QMessageBox.Yes); print("Failed creating word file: Please input data to boxes.")
            return
        if receiver == "":
            reply = QMessageBox.information(self, '확인메세지', '받는회사를 입력해주세요.', QMessageBox.Yes); print("Failed creating word file: Please input data to boxes.")
            return


        self.sender_number.setEnabled(False)
        self.sender_company.setEnabled(False)
        self.sender_name.setEnabled(False)
        self.sender_address.setEnabled(False)
        self.sender_call.setEnabled(False)
        self.sender_fax.setEnabled(False)
        self.sender_business.setEnabled(False)
        self.sender_category.setEnabled(False)
        self.receiver.setEnabled(False)
        self.register_date.setEnabled(False)

        sender = (number, company, name, address, call, fax, business, category)
        receiver = receiver
        try:
            register_date = register_date.split("-")
            r_date = register_date[0] + "년  " + register_date[1] + "월  " + register_date[2] + "일"
        except Exception as e:
            r_date = "-년  -월  -일"
        handle_word.create_sender_table(file_name = file_name, sender = sender, receiver = receiver, r_date = r_date)

        self.sender_number.setEnabled(True)
        self.sender_company.setEnabled(True)
        self.sender_name.setEnabled(True)
        self.sender_address.setEnabled(True)
        self.sender_call.setEnabled(True)
        self.sender_fax.setEnabled(True)
        self.sender_business.setEnabled(True)
        self.sender_category.setEnabled(True)
        self.receiver.setEnabled(True)
        self.register_date.setEnabled(True)

        # For progressbar
        self.doAction()

        self.step = 0
        self.status_bar.setValue(self.step)
        self.status_bar.show()


    def timerEvent(self, e):

        if self.step >= 100:
            self.timer.stop()
            return

        self.step = self.step + 1
        self.status_bar.setValue(self.step)

    def doAction(self):

        if self.timer.isActive():
            self.timer.stop()
        else:
            self.timer.start(10, self)

class FileEdit(QLabel):
    def __init__(self, *args, **kwargs):
        QLabel.__init__(self, *args, **kwargs)
        self.setAcceptDrops(True)

    def setTargetWidget(self, widget):
        self.widget = widget

    def dragEnterEvent(self, event):
        data = event.mimeData()
        urls = data.urls()
        if urls and urls[0].scheme() == 'file':
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        data = event.mimeData()
        urls = data.urls()
        if urls and urls[0].scheme() == 'file':
            event.acceptProposedAction()

    def dropEvent(self, event):
        data = event.mimeData()
        urls = data.urls()
        if urls and urls[0].scheme() == 'file':
            filepath = str(urls[0].path())[1:]
            # self.setText(filepath)
            self.widget.setText(filepath)

            # any file type here
            # if filepath[-4:].upper() == ".txt":
            #     self.setText(filepath)
            # else:
            #     dialog = QMessageBox()
            #     dialog.setWindowTitle("Error: Invalid File")
            #     dialog.setText("Only .txt files are accepted")
            #     dialog.setIcon(QMessageBox.Warning)
            #     dialog.exec_()

def on_button_clicked():
    alert = QMessageBox()
    alert.setText('You clicked the button!')
    alert.exec_()


if __name__ == '__main__':
    # Create app
    app = QtWidgets.QApplication([])

    # Create window
    window = App()
    window.setWindowTitle("v1")
    # window.setStyleSheet("background: #708090")

    # Built-in styles (Fusion, Windows, WindowsVista, GTK+)
    app.setStyle('Fusion')

    # Style
    # app.setStyleSheet("QPushButton { margin: 10ex; }")

    app.exec_()