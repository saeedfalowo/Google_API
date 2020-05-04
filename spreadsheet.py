import gspread
from oauth2client.service_account import ServiceAccountCredentials


# use creds to create a client to interact with the Google Drive API
# scope = ['https://spreadsheets.google.com/feeds']
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)
print(client)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
sheet = client.open("Copy of Legislators 2017").sheet1
# sheet.get_all_values()
sheet_id = sheet.id
print("sheet id: ", sheet_id)
sheet_title = sheet.title
print("sheet title: ", sheet_title)

# Extract and print all of the values
# list_of_hashes = sheet.get_all_records()
# print(list_of_hashes)
# list_of_lists = sheet.get_all_values()
# print(list_of_lists)
row_val = sheet.row_values(1)
# print(row_val)
col_val = sheet.col_values(1)
# print(col_val)
print("Number of rows in the sheet: ", len(col_val))
cell_val= sheet.cell(1, 1).value
# print(cell_val)

# Get the last row entry
last_row_entry = sheet.row_values(len(col_val))
# print(last_row_entry)

# Insert row
# row = ["I am", "inserting", "a new", "row", "into", "the", "spreadsheet", "with", "python"]
# index = len(col_val)+1
# sheet.insert_row(row,index)

# delete row
# sheet.delete_row(len(col_val))

# Find the index of the row with the value "Suozzi"
for i in range(len(col_val)):
	if col_val[i] == "Suozzi":
		indx = i+1 # because python is 0 indexed while the sheet is 1 indexed, thus the +1
		break

if 'indx' in locals():
	print(indx)
	# Get and print all the entries in that row
	indx_row_entry = sheet.row_values(indx)
	print(indx_row_entry)
else:
	print("String not found in list")

# Create a spreadsheet
# newsheet = client.create('Copy Copy Legislators')
# # Share the newly created spreadsheet with the desired email
# newsheet.share('tyslogo@googlemail.com', perm_type='user', role='writer')

sheet2 = client.open("New Spreadsheet").sheet1
# # Copy the contents of Copy of Legislators 2017 and paste in the newly created sheet Copy Copy Legislators
# row_cnt = 1
# for i in range(10):#len(col_val)):
# 	index = i+1#row_cnt
# 	sheet2.insert_row(sheet.row_values(index), index)

# Delete spreadsheet
sheet2_id = sheet2.id
print("Sheet 2 id: ", sheet2_id)
print("Sheet 2 title: ", sheet2.title)
# del_spreadsheet(sheet2_id)

# Create folder
client.create('Leg_Folder', folder_id='Leg_Folder')

# Duplicate a spreadsheet
# copy(sheet_id, title=sheet_title, copy_permissions=True)