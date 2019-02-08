"""
File: config.py

Author: Kayvon Khosrowpour
Description: contains methods for updating and recalling
the configuration file.
""" 

import os
import configparser
import datetime as dt

class Configuration:

    def __init__(self):
        cp = configparser.ConfigParser(default_section=None)
        self.path = os.path.normpath(os.path.abspath(os.path.join(os.getcwd(), 'config.ini')))
        cp.read_file(open(self.path))
        self.save_folder = os.path.abspath(os.path.normpath(cp.get('DEFAULT', 'save_folder')))
        self.start_date = dt.datetime.strptime(cp.get('DEFAULT', 'start_date'), '%m/%d/%Y')
        self.end_date = dt.datetime.today()
        self.email = cp.get('DEFAULT', 'email')
        self.folder = cp.get('DEFAULT', 'folder')
        self.subjects = list(dict(cp.items('SUBJECTS')).values())
        self.cp = cp

        # params to store on-call
        self.sf_tedit = None
        self.sd_qdedit = None
        self.ed_qdedit = None

    def update_start_date(self):
        """ updates start date in config to today """
        newdt = dt.datetime.today()
        today_str = newdt.strftime('%m/%d/%Y')
        self.cp['DEFAULT']['start_date'] =  today_str

        # Writing our configuration file to 'example.ini'
        with open('config.ini', 'w') as configfile:
            self.cp.write(configfile)

        return newdt
