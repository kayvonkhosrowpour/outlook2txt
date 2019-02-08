# Outlook2Txt
Outlook2Txt is a simple python utility to convert Outlook emails to txt files. Using a configuration file, a user can parse through Outlook emails on their Windows computer by date range, subject, account, and folder. No internet connection is required.

## Installation
Outlook2Txt is available only for Windows computers. To install:

* Install any python 3.6 version (I suggest 3.6.7 or 3.6.8)
* Using [pip](https://docs.python.org/3/installing/index.html), [anaconda](https://www.anaconda.com/), or another package manager, install:
	1. [pypiwin32](https://pypi.org/project/pypiwin32/)
	2. [PyQt5](https://pypi.org/project/PyQt5/)
* Install Outlook locally

## Configuration
The configuration file `config.ini` provided demonstrates the structure of the configuration. To use Outlook2Txt:

1. Sign into your Outlook account (Suppose the email address you're using is john@abc.com).
2. Double-click `launcher.py` and click 'Edit Configuration'. This opens the config file in your computer's default text editor.
	- Set `save_folder` to the desired default location of the text file output.
	- Set `start_date` to a date of the form DD/MM/YYYY. The end date is always today's date, but can be changed on running an instance of the progam (see Usage).
	- Set `email` to the email address signed into in Step 1 (e.g. john@abc.com).
	- Set `folder` to the Outlook folder you wish to search through (e.g. Inbox, Outbox, Deleted Items, Sent Items, Junk Email)
	- Add/remove the subjects from the [SUBJECTS] section.
	- Save the file.

## Usage
Outlook2Txt is very simple and easy-to-use. To launch the application, double-click `launcher.py`. You are presented with fields populated from the configuration file `config.ini`. You may change these fields prior to running the application.

Click 'Run' to run the program. Upon success, you will be greeted with a dialog giving information regarding the emails.
