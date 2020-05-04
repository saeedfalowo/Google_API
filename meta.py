class meta:
	my_email	= ''
	temp_cus_email = ''
	my_password	= ''
	filedir = 'forms_download/'

	# Define the spreadsheet columns we are interested in
	col_nums = {
		'last_name_col'		: 1,
		'first_name_col'	: 2,
		'email_col'			: 5,
		'paid_col'			: 7,
		'required_col'		: 8,
		'address_col'		: 10
	}

	receipt_form_cells = {
		'name_cell'				: (2,  6),
		'email_cell'			: (3,  6),
		'address_cell'			: (4,  6),
		'required_cell'			: (9,  5),
		'paid_cell'				: (10, 5),
		'previously_paid_cell'	: (11, 5),
		'total_paid_cell'		: (12, 5),
		'outstanding_cell'		: (13, 5),
		'payment_day_cell'		: (13, 2)
	}
	receipt_form_meta = {
		# 'receipt_form'			: receipt_form,
		'receipt_form_cells'	: receipt_form_cells,
		'receipts_folder_name'	: 'Receipt Bank',
		'receipts_form_name'	: 'Receipt Form'
	}

	client_secret_file = 'client_secret.json'
	googlesheet_name = "Copy Copy Legislators"
	records_filename = 'records.csv'