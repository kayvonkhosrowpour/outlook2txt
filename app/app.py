"""
File: app.py

Author: Kayvon Khosrowpour
Description: implements the GUI and on-click methods for
the application using PyQt5.
"""

import os
import sys
import datetime
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QMainWindow, QHBoxLayout, 
    QGroupBox, QLabel, QDateEdit, QPushButton, QDialog, QTabWidget, QTextEdit)
from PyQt5.QtCore import (Qt, QEvent, QDate)
from app import process

qapp = None

def init_app(args):
    """call before creating ApplicationWindow
    args: system arguments for completeness
    Should only be called once"""
    global qapp
    if qapp is None:
        qapp = QApplication(args)
    else:
        raise ValueError('QApplication initialized more than once.')

def exec_app():
    """call after creating and displaying ApplicationWindow"""
    global qapp
    if qapp is None:
        raise ValueError('QApplication not initialized.')
    else:
        sys.exit(qapp.exec_())

class ApplicationWindow(QMainWindow):
    
    def __init__(self, config):
        QMainWindow.__init__(self)
        self.init_window(config)

    def init_window(self, config):
        """initialize window with view"""
        # window init
        self.setWindowTitle('Outlook2Txt')

        # parent layout
        layout = QVBoxLayout()
        qbox = QGroupBox(parent=self)

        # QLabel, QTextEdit for folder to save txt files to
        sf_widget = QWidget(parent=qbox)
        sf_hbox = QHBoxLayout()
        sf_label = QLabel('Save Folder:', parent=sf_widget)
        sf_tedit = QTextEdit(config.save_folder, parent=sf_widget)
        config.sf_tedit = sf_tedit
        sf_hbox.addWidget(sf_label)
        sf_hbox.addWidget(sf_tedit)
        sf_widget.setLayout(sf_hbox)

        # start date selection widget
        sd_widget, sd_qdedit = self.make_date_row('Start date:',
            config.start_date, qbox)
        config.sd_qdedit = sd_qdedit

        # end date selection widget
        ed_widget, ed_qdedit = self.make_date_row('End date:',
            config.end_date, qbox)
        config.ed_qdedit = ed_qdedit

        # edit config and run program buttons
        ec_button = QParamButton('Edit Config',
            config.path, launch_editor, parent=qbox)
        rp_button = QParamButton('Run',
            config, run_outlook2txt, parent=qbox)

        # add all controls to main layout
        layout.addWidget(sf_widget)
        layout.addWidget(sd_widget)
        layout.addWidget(ed_widget)
        layout.addWidget(ec_button)
        layout.addWidget(rp_button)
        qbox.setLayout(layout)
        self.setCentralWidget(qbox)

    def make_date_row(self, label, datetime, parent=None):
        """ makes a hbox row of date info """
        widget = QWidget(parent=parent)
        hbox = QHBoxLayout()
        label = QLabel('Start date:', parent=widget)
        date = QDate(datetime.year,
            datetime.month,
            datetime.day)
        qdate = QDateEdit(date, parent=widget)
        hbox.addWidget(label)
        hbox.addWidget(qdate)
        widget.setLayout(hbox)

        return widget, qdate

class QParamButton(QPushButton):
    def __init__(self, title, param, action, parent=None):
        QPushButton.__init__(self, title, parent=parent)
        self.param = param
        self.action = action
        self.clicked.connect(self.on_action)

    def on_action(self):
        self.action(self.param)

def launch_editor(file):
    """ launch default editor to edit config file """
    os.startfile(file)
    d = QDialog(None, Qt.WindowSystemMenuHint | Qt.WindowTitleHint)
    d.resize(250, 100)
    qwid = QWidget(parent=d)
    vbox = QVBoxLayout()
    lbl = QLabel('To see config changes in Outlook2Txt,\nreopen the program.')
    b1 = QPushButton('OK')
    b1.clicked.connect(exit)
    vbox.addWidget(lbl)
    vbox.addWidget(b1)
    qwid.setLayout(vbox)
    d.setWindowTitle('Notice')
    d.setWindowModality(Qt.ApplicationModal)
    d.exec_()
 
def qdate_to_datetime(qdate):
    return datetime.datetime(day=qdate.day(), month=qdate.month(), year=qdate.year())

def run_outlook2txt(config):
    # update in-memory config for run
    config.save_folder = config.sf_tedit.toPlainText()
    config.start_date = qdate_to_datetime(config.sd_qdedit.date())
    config.end_date = qdate_to_datetime(config.ed_qdedit.date())

    try:
        # process emails
        num = process.process(config)

        # update config start date to today (last run)
        dt = config.update_start_date()
        config.sd_qdedit.setDate(QDate(dt.year, dt.month, dt.day))

        # tell user
        d = QDialog(None, Qt.WindowSystemMenuHint | Qt.WindowTitleHint)
        d.resize(300, 100)
        qwid = QWidget(parent=d)
        vbox = QVBoxLayout()
        lbl = QLabel('Successfully processed %d emails to\n%s' % (num, config.save_folder))
        b1 = QPushButton('OK')
        b1.clicked.connect(d.close)
        vbox.addWidget(lbl)
        vbox.addWidget(b1)
        qwid.setLayout(vbox)
        d.setWindowTitle('Success')
        d.setWindowModality(Qt.ApplicationModal)
        d.exec_()
        
    except ValueError:
        d = QDialog(None, Qt.WindowSystemMenuHint | Qt.WindowTitleHint)
        d.resize(400, 100)
        qwid = QWidget(parent=d)
        vbox = QVBoxLayout()
        lbl = QLabel('Either config email %s was not found' % config.email + 
            '\nor folder %s was not found in Outlook.' % config.folder +
            ' Cannot process\nemails without valid config.')
        b1 = QPushButton('OK')
        b1.clicked.connect(exit)
        vbox.addWidget(lbl)
        vbox.addWidget(b1)
        qwid.setLayout(vbox)
        d.setWindowTitle('ERROR')
        d.setWindowModality(Qt.ApplicationModal)
        d.exec_()
     