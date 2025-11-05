
# Combined PDF to transformed XLSX script
import pdfplumber
import pandas as pd
import os

def pdf_to_xlsx(pdf_path, temp_xlsx):
	all_tables = []
	with pdfplumber.open(pdf_path) as pdf:
		for page in pdf.pages:
			tables = page.extract_tables()
			for table in tables:
				if table and len(table) > 1:
					df = pd.DataFrame(table[1:], columns=table[0])
					all_tables.append(df)
	if all_tables:
		result = pd.concat(all_tables, ignore_index=True)
		result.to_excel(temp_xlsx, index=False)
		print(f"Intermediate XLSX written to {temp_xlsx}")
		return True
	else:
		print("No tables found in the PDF.")
		return False

def transform_xlsx_format(source_xlsx, output_xlsx):
	df = pd.read_excel(source_xlsx, header=None)
	expected_header_str = 'DESCRIPTION DURATION MEMBER PROJECT TAGS DATE TIME'
	header_indices = []
	for idx, row in df.iterrows():
		first_cell = str(row.iloc[0]).strip().upper()
		if first_cell == expected_header_str:
			header_indices.append(idx)
	header_indices.append(len(df))
	data_blocks = []
	for i in range(len(header_indices)-1):
		start = header_indices[i]
		end = header_indices[i+1]
		block = df.iloc[start+1:end].copy()
		block.columns = expected_header_str.split()
		data_blocks.append(block)
	if data_blocks:
		data = pd.concat(data_blocks, ignore_index=True)
	else:
		print("No data blocks found.")
		return
	data = data.dropna(how='all')
	data = data[data['DESCRIPTION'].str.upper() != 'DESCRIPTION']
	def split_time(row):
		time_val = str(row['TIME']).replace('\n', ' ').replace('  ', ' ').strip()
		date_val = str(row['DATE']).strip()
		if '-' in time_val:
			parts = time_val.split('-')
			if len(parts) == 2:
				start, stop = parts
				return pd.Series([date_val, start.strip(), stop.strip()])
		return pd.Series([date_val, '', ''])
	data[['Start date', 'Start time', 'Stop time']] = data.apply(split_time, axis=1)
	data = data.rename(columns={
		'DESCRIPTION': 'Description',
		'DURATION': 'Duration',
		'MEMBER': 'Member',
		'PROJECT': 'Project',
		'TAGS': 'Tags',
	})

	# Preprocessing steps
	# 1. Remove leading '•' and whitespace/newlines from 'Project'
	if 'Project' in data.columns:
		data['Project'] = data['Project'].astype(str).str.replace(r'^•', '', regex=True).str.replace(r'^\s*', '', regex=True).str.replace('\n', ' ').str.strip()

	# 2. Remove trailing ')' from 'Start date'
	if 'Start date' in data.columns:
		data['Start date'] = data['Start date'].astype(str).str.replace(r'\)', '', regex=True).str.strip()

	# 3. If 'Start date' has two dates separated by '-', keep only the first
	if 'Start date' in data.columns:
		data['Start date'] = data['Start date'].str.split('-').str[0].str.strip()

	# 4. Replace '-' in 'Tags' with 'Unknown'
	if 'Tags' in data.columns:
		data['Tags'] = data['Tags'].astype(str).replace({'-': 'Unknown'}).replace({'nan': 'Unknown'})
		data['Tags'] = data['Tags'].replace(r'^\s*-\s*$', 'Unknown', regex=True)

	final_cols = ['Description', 'Duration', 'Member', 'Project', 'Tags', 'Start date', 'Start time', 'Stop time']
	for col in final_cols:
		if col not in data.columns:
			data[col] = ''
	data = data[final_cols]
	data.to_excel(output_xlsx, index=False)
	print(f"Transformed XLSX saved as {output_xlsx}")

def pdf_to_transformed_xlsx(pdf_path, output_xlsx):
	temp_xlsx = "_temp_pdf_extract.xlsx"
	if pdf_to_xlsx(pdf_path, temp_xlsx):
		transform_xlsx_format(temp_xlsx, output_xlsx)
		os.remove(temp_xlsx)

# Example usage:
pdf_to_transformed_xlsx("files\\input\\TogglTrack_Report_Detailed_report__from_08_28_2025_to_11_04_2025_.pdf", 
						"files\\output\\final_transformed.xlsx")
