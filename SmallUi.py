from PyQt5 import QtWidgets, QtGui
from PyQt5.QtWidgets import (QWidget, QPushButton, QApplication,
                             QHBoxLayout, QVBoxLayout, QGridLayout, QPlainTextEdit)
from PyQt5.QtCore import QDate, QTime, QDateTime, Qt
from PyQt5.QtWidgets import * 
from PyQt5.QtGui import * 
from PyQt5 import QtCore
from PyQt5 import QtGui
from PyQt5.QtWidgets import QApplication, QWidget, QLabel
from PyQt5.QtGui import QIcon, QPixmap
import sys
import GmailApi
import logger

import urllib.request
import webbrowser
req = 'https://www.youtube.com/watch?v=dQw4w9WgXcQ'

WIDTH = 630
HEIGHT = 630



api = GmailApi.GmailApi()
log = logger.sniper_logger()
 
def window():
    app = QApplication(sys.argv)
    
    win = QWidget()
    win.setWindowIcon(QtGui.QIcon('email_sniper_logo.jpg'))
    win.setGeometry(400,400,WIDTH,HEIGHT)
    win.setWindowTitle("Email Sniper Gmail Archiver")
    win.setStyleSheet('background-color: rgb(176, 196, 222)')
    #win.setStyleSheet('QMainWindow{background-color: darkgray;border: 1px solid black;}')
    #win.setStyleSheet("QLineEdit { background-color: yellow }")
    # label = QLabel(win)
    # pixmap = QPixmap('email_sniper_logo.jpg')
    # label.setPixmap(pixmap)
    # win.resize(pixmap.width(),pixmap.height())
    
    #win.setWindowFlag(Qt.FramelessWindowHint)
    email_address = '' #'lospioneerward@gmail.com'
    text_area = QPlainTextEdit(win)
    search_size = 0
    search_date = QDate(1981,12,18)
    last_msg = 'hello'
    logged_in = False


    def show_info_messagebox():
        pop_up = QMessageBox()
        pop_up.setIcon(QMessageBox.Information)

    # setting message for Message Box
        pop_up.setText("How To Information ")
        pop_up.setGeometry(150,150,WIDTH,HEIGHT)
    
    # setting Message box window title
        pop_up.setWindowTitle("How to Use Email Sniper")
    
    # declaring buttons on Message Box
        pop_up.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        pop_up.setFont(QFont('Arial', 10))

    def how_to_video(self):
        webbrowser.open('https://youtu.be/xXheqQOQzio')


    # def show():



    #     print(line.text())


    def send_user_message(msg):
        if (msg == last_msg):
            return
        if (msg == ''):
            text_area.clear()
        text_area.insertPlainText('\n'+msg)
        log.write(msg);
        text_area.repaint() #they say not to use repaint,  but in this application it seems to work well.
        #text_area.update()

    def validate_email():
        print ('email is valid')

    def email_check():
        email_address = user_email.text()
        print ('email is checked:',email_address)

    def api_connect():
        #if logged_in == False:
        #api = GmailApi.GmailApi(email_address)
        api.set_email(user_email.text())
        
        logged_in = True

    def build_search():
        #get data from combo boxes
        st = search1.currentText() + search1_item.text()
        return st

    def do_search(copy, delete):
        email_address = user_email.text()
        #search_date = email_min_date.date()
        save_location = save_path.text()
        log.start(save_location);
        if (len(email_address) < 3):
            send_user_message('you need to enter your email address first')
        else:
            api_connect()
            send_user_message('')
            send_user_message('Connecting...')

            # get emails that match the query you specify            
            
            api.set_feedback(send_user_message, log)
            api.save_path = save_location
            api.set_copy_delete(copy, delete)
            api_service = api.get_service()
            api.reset_search_count()
            search_string = build_search()
            results = api.search_messages(api_service, search_string)
            # for each email matched, read it (output plain/text to console & save HTML and attachments)
            for msg in results:
                #send_user_message("  reading message "+str(api.get_search_count()))
                api.read_message(api_service, msg)
                #send_user_message(" --- --- --- -g- -r- -e- -e- -n- -m- -a- -n- --- --- ---")

            if delete:
                results = api.delete_messages(api_service, search_string)
                for msg in results:
                    #send_user_message("  deleted message "+str(api.get_search_count()))
                    api.read_message(api_service, msg)
                    #send_user_message(" --- --- --- -g- -r- -e- -e- -n- -m- -a- -n- --- --- ---")

            if not copy and not delete:
                send_user_message('Email preview complete')
            elif copy and not delete:
                send_user_message('Email copy complete')
            elif copy and delete:
                send_user_message('Email archive complete')
            send_user_message('Emails:'+str(api.get_search_count()))
            send_user_message('google space:'+api.get_search_size())
        log.close();

    def search():
        do_search(False, False) #don't save or delete

    def copy():
        do_search(True, False) #do save but don't delete

    def delete():
        do_search(True, True) #do save and delete

    def get_save_path():
        print(save_path.text())
 
    #.....build the form
    label = QtWidgets.QLabel(win)
    label.setText("Enter full gmail address:")
    label.setFont(QFont('Arial', 10))

    user_email = QtWidgets.QLineEdit(win)
    #user_email.setFixedWidth(140)
    #user_email.text = email_address
    user_email.setMaxLength(256)
    user_email.editingFinished.connect(email_check)

    user_email_layout = QHBoxLayout()
    user_email_layout.addWidget(label)
    user_email_layout.addWidget(user_email)

    search1 = QtWidgets.QComboBox(win)
    search1.addItems(["older_than:", "newer_than:","from:","to:","cc:","bcc:","subject:","larger", "smaller", "has:","is:","category:","label:"])
    search1.setFont(QFont('Arial', 10))
    search1_item = QtWidgets.QLineEdit(win)
    #email_min_size.editingFinished.connect(get_email_size)

    email_size_layout = QHBoxLayout()
    email_size_layout.addWidget(search1)
    email_size_layout.addWidget(search1_item)

    #file
    file_label = QtWidgets.QLabel(win)
    file_label.setText("save file location: ")
    file_label.setFont(QFont('Arial', 10))

    save_path = QtWidgets.QLineEdit(win)
    save_path.setMaxLength(256)
    save_path.editingFinished.connect(get_save_path)

    save_layout = QHBoxLayout()
    save_layout.addWidget(file_label)
    save_layout.addWidget(save_path)

    #label3 = QtWidgets.QLabel(win)
    #label3.setText("return emails older than:")

    #email_min_date = QtWidgets.QDateEdit(win)
    ##email_min_date.setMinimumDate(QDate(1900,1,1))
    ##email_min_date.SetMaximumDate(QDate(2100,12,31))
    #email_min_date.setDate(search_date)
    #email_min_date.editingFinished.connect(get_search_date)

    #email_age_layout = QHBoxLayout()
    #email_age_layout.addWidget(label3)
    #email_age_layout.addWidget(email_min_date)

    preview_button = QtWidgets.QPushButton(win)
    preview_button.clicked.connect(search)
    preview_button.setText(" PREVIEW ")
    preview_button.setFont(QFont('Arial', 10))
    preview_button.setStyleSheet("QPushButton::hover"
                                "{"
                                "background-color : lightgreen;;"
                                "}")

    cancel_button = QtWidgets.QPushButton(win)
    cancel_button.clicked.connect(copy)
    cancel_button.setText(" COPY ")
    cancel_button.setFont(QFont('Arial', 10))
    cancel_button.setStyleSheet("QPushButton::hover"
                                "{"
                                "background-color : lightgreen;;"
                                "}")

    
    delete_button = QtWidgets.QPushButton(win)
    delete_button.clicked.connect(delete)
    delete_button.setText(" ARCHIVE ")
    delete_button.setFont(QFont('Arial', 10))
    delete_button.setStyleSheet("QPushButton::hover"
                                "{"
                                "background-color : lightgreen;;"
                                "}")


    info_button = QtWidgets.QPushButton(win)
    info_button.setText("HOW-TO")
    #info_button.setGeometry(5,5,WIDTH,HEIGHT)
    info_button.setFont(QFont('Arial', 10))
    info_button.setStyleSheet("QPushButton::hover"
                            "{"
                            "background-color : lightgreen;;"
                            "}")
    info_button.clicked.connect(how_to_video)

    button_layout = QHBoxLayout()
    button_layout.addWidget(preview_button)
    button_layout.addWidget(cancel_button)
    button_layout.addWidget(delete_button)
    button_layout.addWidget(info_button)

    
    text_area.insertPlainText('This is the preview and message area.')
    text_area.insertPlainText('\nEnter your email address, and click a button.')
    text_area.insertPlainText('\nI suggest that you start with PREVIEW to try things out')
    text_area.setStyleSheet('background-color: rgb(152, 251, 152)')
    text_area.setFont(QFont('Arial', 10))

    #stretch challenge pop_ups

   


    vrsl = QVBoxLayout()
    vrsl.addLayout(user_email_layout)
    vrsl.addLayout(email_size_layout)
    #vrsl.addLayout(email_age_layout)
    vrsl.addLayout(save_layout)
    vrsl.addLayout(button_layout)
    vrsl.addWidget(text_area)  

    win.setLayout(vrsl)
 
    win.show()
    sys.exit(app.exec_())
     
window() 
