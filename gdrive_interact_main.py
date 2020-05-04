import os.path
from os import path
from datetime import date
import csv
import gdrive_interact_functions as gfunc
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from email_shell import sendEmailWithAttachment
from meta import meta

# list out the permissions in scope
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
# build a mimeType dictionary
mimeType = {
	'folder': 'application/vnd.google-apps.folder',
	'spreadsheet': 'application/vnd.google-apps.spreadsheet',
	'downloadxlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
}
# Extract the defined metadata from class meta in meta.py
my_email			= meta.my_email
my_password			= meta.my_password
temp_cus_email 		= meta.temp_cus_email
filedir 			= meta.filedir
client_secret_file	= meta.client_secret_file
googlesheet_name 	= meta.googlesheet_name
col_nums 			= meta.col_nums
receipt_form_cells	= meta.receipt_form_cells
receipt_form_meta	= meta.receipt_form_meta
records_filename 	= meta.records_filename

# Setup the credentials using the client_secret_file and the permissions deifined in scope
creds = ServiceAccountCredentials.from_json_keyfile_name(client_secret_file, scope)
client = gspread.authorize(creds) # authorizes the credentials

# Open the google sheet with the name specified in googlesheet_name
sheet = client.open(googlesheet_name).sheet1
# Retrive the first column in the googlesheet
col1_val = sheet.col_values(1)
col1_len = len(col1_val) # gets the number of rows in the googlesheet

today = date.today() # retrieve today's date
date_today = today.strftime("%d/%m/%Y") # change date format
cwd_dir = os.getcwd() # get current working directory

# Open the template receipt form for editing
receipt_form = client.open(receipt_form_meta['receipts_form_name']).sheet1
receipt_form_meta['receipt_form'] = receipt_form # Add receipt form to receipt_form_meta dictionary

names_encountered = [] # initialize encountered names list
if path.exists(records_filename) == False: # if the records file does not exist in path
	with open(records_filename, 'w', newline='') as f: # Create and open the file to write into
		csv_writer = csv.writer(f) # Creates the csv writer object
		# Write the col1 values along with each row index into the csv
		for i in range(col1_len):
			csv_writer.writerow([str(i+1), col1_val[i]]) # +1 because python is 0 indexed

	# Loop through all entries, rows, in the Copy Copy Legislators spreadsheet
	print(col1_val)
	for i in range(col1_len-1): # -1 to ignore first row, column name row
		sheet_indx = i+1 # +1 because python is 0 indexed while the sheet row number is 1 indexed
		# Call the fillReceiptForm fill the Receipt form with values extracted from the spreadsheet 
		name, cus_email, file_id = gfunc.fillReceiptForm(sheet, sheet_indx, names_encountered, receipt_form_meta, date_today, [], col_nums)
		names_encountered.append(name) # adds the custermer name to names_encountered list
		filename = name + '.xlsx' # defines the name to download the customer receipt form as
		# Call the downloadFile function to download the customer receipt form
		gfunc.downloadFile(file_id, filedir+filename, mimeType['downloadxlsx'])
		# Call the sendEmailWithAttachment to send the customer an email with their receipt form attached
		sendEmailWithAttachment(my_email, my_password, name, temp_cus_email, filedir, filename)
		# Call the del_folderContent to delete the content of the download folder
		gfunc.del_folderContent(filedir)

	# Clear the receipt form fields
	gfunc.clearForm(receipt_form_meta)

else: # if the records file already exists
	with open(records_filename, 'r') as f: # opens the file to read from
		csv_reader = csv.reader(f) # creates the csv reader object
		records = [] # initializes the records list
		for line in csv_reader: # extracts content from the records csv file line by line
			records.append(line) # appends csv file lines to records list

	records_last_names = [] # initializes the records last names list
	# Loop through each line in the records list and extract the customers' last names
	[records_last_names.append(rec_line[1]) for rec_line in records] # appends last names to records_last_names list

	# Compare the length of googlesheet rows with number of entries in the csv
	if col1_len == len(records): # If number of rows in the googlesheet is identical to the number of rows in the records file
		# Call the checkContent function to compare the contents of the spreadsheet with it previously recorded contents 
		gfunc.checkContent(records, col1_val, same_len=True)

	else:
		records_len = len(records) # gets the number of rows in the records file
		new_entries_num = col1_len-records_len # calculates the number of new entries 
		# Call the checkContent function to compare the contents of the spreadsheet with it previously recorded contents 
		gfunc.checkContent(records, col1_val[:-new_entries_num], same_len=False)

		# if googlesheet content has more rows than the records content
		# and if the final last name in the records is identical to that in the googlesheet at the same index
		if col1_len > records_len and records[-1][1] == col1_val[records_len-1]:
			for i in range(new_entries_num): # loops through all the new googlesheet entries
				sheet_row=i+1+records_len # num rows in records + the current index, i, + 1 because python is 0 indexed

				# Extract the new entry's first and last names
				last_name = sheet.cell(sheet_row, 1).value
				first_name = sheet.cell(sheet_row, 2).value
				name = first_name.capitalize() + ' ' + last_name.capitalize() # constructs the full name string
				# Call the getPrevPay function to find out how much the customer as paid before
				previously_paid = gfunc.getPrevPay(sheet, records_last_names, last_name, col_nums, new=False)
				# Call the fillReceiptForm fill the Receipt form with values extracted from the spreadsheet 
				name, cus_email, file_id = gfunc.fillReceiptForm(sheet, i+records_len, names_encountered, receipt_form_meta, date_today, previously_paid, col_nums)
				names_encountered.append(name) # adds the custermer name to names_encountered list
				filename = name + '.xlsx'# defines the name to download the customer receipt form as
				# Call the downloadFile function to download the customer receipt form
				gfunc.downloadFile(file_id, filedir+filename, mimeType['downloadxlsx'])				
				# Call the sendEmailWithAttachment to send the customer an email with their receipt form attached
				sendEmailWithAttachment(my_email, my_password, name, temp_cus_email, filedir, filename)
				# Call the del_folderContent to delete the content of the download folder
				gfunc.del_folderContent(filedir)

			# Clear form fields
			gfunc.clearForm(receipt_form_meta)

			with open(records_filename, 'a+', newline='') as f: # opens the file to be updated with new entries
				csv_writer = csv.writer(f) # creates the csv writer object
				for i in range(new_entries_num): # loop through the new entries
					csv_writer.writerow([records_len+i+1, col1_val[records_len+i]]) # appends new entries to existing csv file data