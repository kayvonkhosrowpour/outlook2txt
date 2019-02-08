"""
File: launcher.py

Author: Kayvon Khosrowpour
Description: launcher of the outlook2txt program, which converts
emails from outlook within a certain range, given keywords for
the subject line, and outputs them to .txt files in a specified
directory.
"""

import sys
from PyQt5.QtWidgets import QApplication
from app import (config, app)

def main():
    # get config
    conf = config.Configuration()

    # initialize and launch app
    app.init_app(sys.argv)

    # make app
    appWindow = app.ApplicationWindow(conf)
    appWindow.show()

    # close
    app.exec_app()

if __name__ == '__main__':
    main()