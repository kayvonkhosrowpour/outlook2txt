"""
File: process.py

Author: Kayvon Khosrowpour
Description: processes the emails in Outlook. Much code is repurposed
from https://www.codementor.io/aliacetrefli/how-to-read-outlook-emails-by-python-jkp2ksk95
"""

import os
import win32com.client

def process(config):
	# create folder to save files
	if not os.path.exists(config.save_folder):
		os.makedirs(config.save_folder)

	outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
	accounts = win32com.client.Dispatch('Outlook.Application').Session.Accounts;
	
	# get account of interest
	str_accounts = [str(a) for a in accounts]
	try:
		account = accounts[str_accounts.index(config.email)]
	except ValueError:
		raise ValueError('No such account found.')

	# get folder of interest
	folders = outlook.Folders(account.DeliveryStore.DisplayName).Folders
	str_folders = [str(f) for f in folders]
	try:
		folder = folders[str_folders.index(config.folder)]
	except:
		raise ValueError('No such folder found.')

	# look through all email messages
	num_processed = 0
	for i, m in enumerate(folder.Items):
		# is the message in range?
		last_mod = m.LastModificationTime.replace(tzinfo=None)
		in_range = (last_mod >= config.start_date and last_mod <= config.end_date)

		# does message match at least one subject?
		match_sub = False
		for s in config.subjects:
			if s in m.Subject:
				match_sub = True
				break

		# output as text file
		if in_range and match_sub:
			# get non-conflicting filename
			filename = os.path.join(config.save_folder, 'message%d.txt' % i)
			ct = 1
			while os.path.isfile(filename):
				filename = os.path.join(config.save_folder, 'message%d-%d.txt' % (i, ct))
				ct += 1
			# output txt file
			f = open(filename, 'w')
			f.write(m.Body)
			f.close()
			num_processed += 1

	return num_processed

