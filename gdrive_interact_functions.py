from __future__ import print_function
import httplib2
import os
import io
import glob
import auth

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from apiclient import errors
from apiclient.http import MediaIoBaseDownload

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/drive-python-quickstart.json
# SCOPES = 'https://www.googleapis.com/auth/drive', 'https://spreadsheets.google.com/feeds',
SCOPES = [
	'https://www.googleapis.com/auth/drive',
	'https://www.googleapis.com/auth/drive.file',
	'https://www.googleapis.com/auth/drive.metadata',
	'https://www.googleapis.com/auth/drive.appdata',
	]
CLIENT_SECRET_FILE = 'client_secret2.json'
APPLICATION_NAME = 'Drive API Python Quickstart'

auth_Inst = auth.auth(SCOPES,CLIENT_SECRET_FILE,APPLICATION_NAME)
credentials = auth_Inst.get_credentials()

http = credentials.authorize(httplib2.Http())
drive_service = discovery.build('drive', 'v3', http=http)

mimeType = {
	'folder': 'application/vnd.google-apps.folder',
	'spreadsheet': 'application/vnd.google-apps.spreadsheet',
	'downloadxlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
}

def list_files():
    # Call the google drive api files().list() method to list all the files and folder
    # within google drive
    results = drive_service.files().list(
        pageSize=10,fields="nextPageToken, files(id, name)").execute()
    # Extract the file properties with the get() method
    items = results.get('files', [])
    # Print the results
    if not items:
        print('No files found.')
    else:
        print('Files:')
        # Print the file name and id for each file found
        for item in items:
            print('{0} ({1})'.format(item['name'], item['id']))

def searchFile(size,query):
	# Call the google drive api files().list() method with query parameter to specify
	# the exact file property to look for
    results = drive_service.files().list(
    pageSize=size,fields="nextPageToken, files(id, name, kind, mimeType)",q=query).execute()
    items = results.get('files', [])
    if not items:
        print('No files found.')
        return 0
    else:
        print('Files:')
        for item in items:
            # print(item)
            print('{0} ({1})'.format(item['name'], item['id']))
        return items


def createFolder(name):
	# Define the folder metadata: name and folder mimeType
    file_metadata = {
    'name': name,
    'mimeType': 'application/vnd.google-apps.folder'
    }
    # Call the google drive api files().create() method to create the folder
    file = drive_service.files().create(body=file_metadata,
                                        fields='id').execute()
    print ('Folder ID: %s' % file.get('id'))

def insert_file_into_folder(folder_id, file_id):
	# Retrieve the file to be moved using its unique file id
	file = drive_service.files().get(fileId=file_id, fields='parents').execute()
	# Retrieve the existing parents to remove
	previous_parents = ",".join(file.get('parents'))
	# Move the file to the new folder with the update() method
	file = drive_service.files().update(fileId=file_id,
		addParents=folder_id,
		removeParents=previous_parents,
		fields='id, parents').execute()

def createSpreadsheet(name):
	# Define the spreadsheet metadata: name and file mimeType
	metadata = {
		'name': name,
		'mimeType': 'application/vnd.google-apps.spreadsheet',
	}
	# Call the google drive api files().create() method to create the spreadsheet according
	# to the specified metadata
	sheet = drive_service.files().create(body=metadata).execute()

def copyFiletoFolder(file_id, folder_id, newFile_name):
	# Call the google drive api files().copy() method to duplicate the file specified
	# by the unique file id
	copy = (
		drive_service.files().copy(
			fileId=file_id, body={"title": "copiedFile"}).execute()
		)
	# Define the file metadata: name
	metadata = {'name': newFile_name}
	# Get the copied file's file id using the .get('id') method
	copied_file_id = copy.get("id")
	# Call the google drive api files().update() method to rename the duplicated file, 
	# new name defined in the metadata
	drive_service.files().update(
		fileId=copied_file_id,
		body=metadata).execute()
	# Call the insert_file_into_folder function to move the renamed copied file to a new folder
	# specified by its unique folder id
	insert_file_into_folder(folder_id, copied_file_id)
	# return the copied file id
	return copied_file_id

def existinList(List,target):
	# Loop through the elements in the provided list to find the specified element in target
	for elem in List:
		# if found, return True to indicate that the target exists in the list
		if elem==target:
			return True
	# else, return False to specify the target does not exist in the list
	return False

def fillReceiptForm(sheet, i, names_encountered, receipt_form_meta, date_today, previously_paid, col_nums):
	# Extract the interested column numbers
	first_name_col	= col_nums['first_name_col']
	last_name_col	= col_nums['last_name_col']
	email_col 		= col_nums['email_col']
	address_col 	= col_nums['address_col']
	paid_col 		= col_nums['paid_col']
	required_col	= col_nums['required_col']

	# Extract receipt form meta
	receipt_form 			= receipt_form_meta['receipt_form']
	receipts_folder_name	= receipt_form_meta['receipts_folder_name']
	receipts_form_name		= receipt_form_meta['receipts_form_name']

	# Extract the data from the cells in the interested columns for each row in the spreadsheet
	# ignore the first row, the column name row
	ign = 1
	first_name	= sheet.cell(i+ign, first_name_col).value
	last_name	= sheet.cell(i+ign, last_name_col).value
	name 		= first_name.capitalize() + ' ' + last_name.capitalize()

	email 		= sheet.cell(i+ign, email_col).value
	address 	= sheet.cell(i+ign, address_col).value
	paid 		= sheet.cell(i+ign, paid_col).value
	required 	= sheet.cell(i+ign, required_col).value
	payment_day = date_today

	# if previously_paid was not declared
	if not previously_paid:
		previously_paid = getPrevPay(sheet, names_encountered, name, col_nums, new=True)

	total_paid 	= float(paid) + float(previously_paid)
	outstanding = float(required) - total_paid

	# Construct the cell values dictionary
	cell_vals = {
		'name'				: name,
		'email'				: email,
		'address'			: address,
		'paid'				: paid,
		'previously_paid'	: previously_paid,
		'required'			: required,
		'total_paid'		: total_paid,
		'outstanding'		: outstanding,
		'payment_day'		: payment_day
	}

	# Call the populateForm function to fill the receipt form
	populateForm(receipt_form_meta, cell_vals)

	# Search for receipts folder
	receipts_folder = searchFile(10,"name='%s'" % receipts_folder_name)
	if receipts_folder == 0:
		createFolder(receipts_folder_name)
		receipts_folder = searchFile(10,"name='%s'" % receipts_folder_name)

	receipts_folder_id = receipts_folder[0]['id']

	# Search for the specific custumer receipts folder
	cus_receipt_folder = searchFile(10,("name='%s'" % name))
	if cus_receipt_folder == 0:
		createFolder(name)
		cus_receipt_folder = searchFile(10,("name='%s'" % name))
		cus_receipt_folder_id = cus_receipt_folder[0]['id']
		insert_file_into_folder(receipts_folder_id, cus_receipt_folder_id)

	cus_receipt_folder_id = cus_receipt_folder[0]['id']

	# Copy the Receipt form to customer folder
	receipt_form_item = searchFile(10,"name='%s'" % receipts_form_name)
	copy_id = copyFiletoFolder(receipt_form_item[0]['id'], cus_receipt_folder_id, name + ' - ' + date_today)
	return name, email, copy_id

def getPrevPay(sheet, names_encountered, name, col_nums, new):
	paid_col = col_nums['paid_col']
	"""
	The spreadsheet row number is 1 indexed
	The first row contains the column names
	Therefore, the first data row starts from index 2, i.e. shift = 2
	This will then be added to the python index
	python index 0 will then become 2 -> row 2

	If only a portion of the rows are being examined, however,
	And since the spreadsheet row number still starts from 1,
	Only 1 needs to be added to the python index to get the data within the
	spreadsheet rows section, i.e. shift = 1
	This is because this section will not have the column name row to ignore
	"""
	if new==True:
		shift = 2
	else:
		shift = 1

	# If the customer name has never been encountered
	if existinList(names_encountered, name) == False:
		previously_paid = 0
		return previously_paid
	else:
		# Create a list of with index positions where the customer name appeared in the spreadsheet
		indx = [j for j, cus_name in enumerate(names_encountered) if cus_name == name]
		# print(names_encountered)
		# print('index: ', indx)
		# If the customer name only appeard once
		if len(indx) < 2:
			# Retrieve the last paid value in the 
			previously_paid = sheet.cell(indx[0]+shift, paid_col).value
		else:
			previously_paid = 0
			# Add up the value in the paid column cells where the customer appeared
			for k in range(len(indx)):
				previously_paid += float(sheet.cell(indx[k]+shift, paid_col).value)

		return previously_paid

def populateForm(receipt_form_meta, cell_vals):
	# Populate the receipt form cell
	receipt_form = receipt_form_meta['receipt_form']
	receipt_form_cells = receipt_form_meta['receipt_form_cells']
	name_cell			= receipt_form_cells['name_cell']
	email_cell			= receipt_form_cells['email_cell']
	address_cell		= receipt_form_cells['address_cell']
	required_cell		= receipt_form_cells['required_cell']
	paid_cell			= receipt_form_cells['paid_cell']
	prev_paid_cell 		= receipt_form_cells['previously_paid_cell']
	total_paid_cell		= receipt_form_cells['total_paid_cell']
	outstanding_cell	= receipt_form_cells['outstanding_cell']
	payment_day_cell	= receipt_form_cells['payment_day_cell']

	# Replace the contents on the specified cells with empty strings to clear the form
	receipt_form.update_cell(name_cell[0], name_cell[1], cell_vals['name'])
	receipt_form.update_cell(email_cell[0], email_cell[1], cell_vals['email'])

	receipt_form.update_cell(address_cell[0], address_cell[1], cell_vals['address'])
	receipt_form.update_cell(required_cell[0], required_cell[1], cell_vals['required'])

	receipt_form.update_cell(paid_cell[0], paid_cell[1], cell_vals['paid'])
	receipt_form.update_cell(prev_paid_cell[0], prev_paid_cell[1], cell_vals['previously_paid'])	
	receipt_form.update_cell(total_paid_cell[0], total_paid_cell[1], cell_vals['total_paid'])

	receipt_form.update_cell(outstanding_cell[0], outstanding_cell[1], cell_vals['outstanding'])
	receipt_form.update_cell(payment_day_cell[0], payment_day_cell[1], cell_vals['payment_day'])

def clearForm(receipt_form_meta):
	# Populate the receipt form cell
	receipt_form = receipt_form_meta['receipt_form']
	receipt_form_cells 	= receipt_form_meta['receipt_form_cells']

	name_cell			= receipt_form_cells['name_cell']
	email_cell			= receipt_form_cells['email_cell']
	address_cell		= receipt_form_cells['address_cell']
	required_cell		= receipt_form_cells['required_cell']
	paid_cell			= receipt_form_cells['paid_cell']
	prev_paid_cell 		= receipt_form_cells['previously_paid_cell']
	total_paid_cell		= receipt_form_cells['total_paid_cell']
	outstanding_cell	= receipt_form_cells['outstanding_cell']
	payment_day_cell	= receipt_form_cells['payment_day_cell']

	# Replace the contents on the specified cells with empty strings to clear the form
	receipt_form.update_cell(name_cell[0], name_cell[1], '') # Name Cell
	receipt_form.update_cell(email_cell[0], email_cell[1], '') # Email Cell

	receipt_form.update_cell(address_cell[0], address_cell[1], '') # Address Cell
	receipt_form.update_cell(required_cell[0], required_cell[1], '') # Product cost Cell

	receipt_form.update_cell(paid_cell[0], paid_cell[1], '')# Paid Cell
	receipt_form.update_cell(prev_paid_cell[0], prev_paid_cell[1], '')# Previously paid Cell
	receipt_form.update_cell(total_paid_cell[0], total_paid_cell[1], '')# Total paid Cell

	receipt_form.update_cell(outstanding_cell[0], outstanding_cell[1], '')# Outstanding Cell
	receipt_form.update_cell(payment_day_cell[0], payment_day_cell[1], '')# Date Cell

def downloadFile(file_id,filepath,mimetype):
    # Call the google drive api files().export_media() to retrieve a file specified by its
    # file id and mimeType
    request = drive_service.files().export_media(fileId=file_id,
                                             mimeType=mimetype)
    # Allocate some ram space for the binary file to be downloaded
    fh = io.BytesIO()
    # Setup a MediaIoBaseDownload method downloader with the file to download (request) and
    # the download location (fh)
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    # Call the next_chunk() downloader method to download the binary file one chunk at a time
    # until the done == True
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))

    # save the file into the specified file path
    with io.open(filepath,'wb') as f:
    	# Call the seek(0) method to find the absolute position of the allocated ram space
    	# where the file was downloaded
        fh.seek(0)
        # Read the binary file from the ram space, then write the file into the specified
        # file path
        f.write(fh.read())

def del_folderContent(folderpath):
	# Collect all the files in the folder using the folder path + '/*'
	files = glob.glob(folderpath+'/*')
	# Loop through all the files collected
	for f in files:
		# Remove each file from the path with the remove() method
	    os.remove(f)

def checkContent(records, col1_val, same_len):
	if len(records) > len(col1_val): # if there are more data in records
		raise SystemError("There are more data in records than in the spreadsheet!")

	not_equal_indx = [] # initializes the not_equal_indx list
	# Check if the contents of the first col in the googlesheet is the same as those in the records csv file
	for i in range(len(records)): # Loops through the contents of records
		if records[i][1] != col1_val[i]: # checks if the last names at an index is different in records and the googlesheet
			not_equal_indx.append(i) # notes the index where there are differences

	if not not_equal_indx: # if not_equal_indx list is empty
		if same_len==True: # if records and the spreadsheet has the same number of rows as specified by same_length
			print("No updates were found!")
		else: # if records and the spreadsheet do not have the same number of rows as specified by same_length
			print("No changes to previously recorded data found!")
	else: # If not_equal_indx list is not empty
		raise SystemError("The contents of the googlesheet does not match what was recorded at these indices: %s " % not_equal_indx)
