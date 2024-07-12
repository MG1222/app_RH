import json
import os
import re
import logging
import time
from datetime import datetime
import concurrent.futures
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from dateutil import parser


class ExcelOperations:
	def __init__(self):
		self.params = self.load_params()
	
	def load_params(self):
		try:
			with open("./config/setting_excel.json", 'r') as f:
				return json.load(f)
		except FileNotFoundError:
			logging.error('json file not found')
			return None
	
	def process_excel_files(self, folder_path, progress_bar=None):
		information = []
		excel_files = self.get_excel_files(folder_path)
		total_files = len(excel_files)
		if progress_bar:
			progress_bar['value'] = 5
			progress_bar.update()
		for index, (file_path, regex) in enumerate(excel_files):
			try:
				info_from_excel = self.extract_information_from_excel(file_path, regex)
				if info_from_excel is None:
					continue
				
				if any(info_from_excel['last_name'] == item['last_name'] and
				       info_from_excel['first_name'] == item['first_name'] for item in information):
					continue
				information.append(info_from_excel)
			
			except Exception as e:
				logging.error(f"Error while processing file {file_path}: {e}")
			
			if progress_bar:
				progress_value = ((index + 1) / total_files) * 100
				progress_bar['value'] = progress_value
				progress_bar.update()
		
		if progress_bar:
			progress_bar['value'] = 100
			progress_bar.update()
		
		if not information:
			return None
		
		return information
	
	def get_excel_files(self, folder_path):
		excel_files = []
		try:
			with concurrent.futures.ThreadPoolExecutor() as executor:
				futures = []
				for root, dirs, files in os.walk(folder_path):
					for file in files:
						if file.endswith(".xlsx") or file.endswith(".xls"):
							file_path = os.path.join(root, file)
							futures.append(executor.submit(self.check_file, file_path, root))
				
				for future in concurrent.futures.as_completed(futures):
					result = future.result()
					
					if result:
						excel_files.append(result)
		
		except Exception as e:
			logging.error(f"Error while fetching Excel files: {e}")
		
		return excel_files
	
	def check_file(self, file_path, root):
		if not os.path.exists(file_path):
			return None
		try:
			is_open = self.is_file_not_open(file_path)
			
			if is_open:
				check_format = self.verify_recruitment_file_format(file_path)
				
				if check_format:
					return file_path, self.extract_path_profile(root)
			return None
		except PermissionError:
			time.sleep(1)
			logging.error("Permission Error")
		except Exception as e:
			logging.error(f"Error processing file {file_path}: {e}")
		return None
	
	def extract_information_from_excel(self, file_path, regex):
		df = pd.read_excel(file_path)
		
		extracted_rows = df.iloc[1:8, :]
		status_label_row = 19
		status_value_row = 20
		status_column = "Unnamed: 6"
		
		status_value = df.at[status_value_row, status_column]
		extracted_info = {
			"last_name": "",
			"first_name": "",
			"status": status_value,
			"direction": regex[1],
			'profile': regex[0],
			"tel": "",
			"email": "",
			"interviews": []
		}
		
		phone_regex = r'(?:Tél portable: |tel : )?0?\d{1}[\s.]?(\d{2}[\s.]?){3}\d{2}|\d{10}'
		email_cleaning_regex = r'^(Mail: |Email: )?(.*)'
		
		for index, row in extracted_rows.iterrows():
			for col in df.columns:
				value = str(row[col])
				if "Nom:" in value:
					extracted_info["last_name"] = df.iloc[index, df.columns.get_loc(col) + 1]
				elif "Prénom:" in value:
					extracted_info["first_name"] = df.iloc[index, df.columns.get_loc(col) + 1]
				elif re.search(phone_regex, value):
					phone_number = re.search(phone_regex, value).group()
					phone_number = re.sub(r'\D', '', phone_number)
					formatted_phone_number = '.'.join(phone_number[i:i + 2] for i in range(0, len(phone_number), 2))
					extracted_info["tel"] = formatted_phone_number
				elif "@" in value:
					cleaned_email = re.sub(email_cleaning_regex, r'\2', value)
					extracted_info["email"] = cleaned_email
				elif "Entretien" in value:
					raw_date = df.iloc[index, df.columns.get_loc(col) + 2]
					if raw_date == "___/___/__" or pd.isna(raw_date):
						formatted_date = "Date inconnue"
					else:
						try:
							parsed_date = parser.parse(str(raw_date),
							                           dayfirst=True)
							formatted_date = parsed_date.strftime('%d/%m/%y')
						except ValueError:
							formatted_date = "Date inconnue"
					interview_info = {
						"date": formatted_date,
						"manager": df.iloc[index, df.columns.get_loc(col) + 1]
					}
					extracted_info["interviews"].append(interview_info)
		
		return extracted_info
	
	def create_new_excel_file(self, information, output_path):
		if not information:
			raise ValueError("No information to write, file would be empty except headers")
		
		headers = ['REPARTOIRE', 'PROFIL', 'NOM', 'PRENOM', 'EMAIL', 'DISPONIBILITE',
		           'DATE1', 'ENTRETIEN1', 'DATE2', 'ENTRETIEN2', 'DATE3', 'ENTRETIEN3', 'DERNIER ENTRETIEN']
		df = pd.DataFrame(columns=headers)
		
		rows_list = []
		for info in information:
			interviews = info.get('interviews', [])
			interview_info = {}
			dates = []
			
			for i, interview in enumerate(interviews, start=1):
				date_key = f'DATE{i}'
				manager_key = f'ENTRETIEN{i}'
				
				verified_date = self.verify_format_date(interview.get('date'))
				interview_info[date_key] = verified_date or ''
				interview_info[manager_key] = interview.get('manager', '')
				if verified_date:
					dates.append(verified_date)
			
			for i in range(1, 4):
				date_key = f'DATE{i}'
				manager_key = f'ENTRETIEN{i}'
				if date_key not in interview_info:
					interview_info[date_key] = ''
				if manager_key not in interview_info:
					interview_info[manager_key] = ''
			
			last_interview_date = self.find_last_interview(dates)
			
			new_row = {
				"REPARTOIRE": info['direction'],
				"PROFIL": info['profile'],
				"NOM": info['last_name'],
				"PRENOM": info['first_name'],
				"EMAIL": info['email'],
				"DISPONIBILITE": info['status'],
				"DATE1": interview_info['DATE1'],
				"ENTRETIEN1": interview_info['ENTRETIEN1'],
				"DATE2": interview_info['DATE2'],
				"ENTRETIEN2": interview_info['ENTRETIEN2'],
				"DATE3": interview_info['DATE3'],
				"ENTRETIEN3": interview_info['ENTRETIEN3'],
				"DERNIER ENTRETIEN": last_interview_date
			}
			rows_list.append(new_row)
		
		new_rows_df = pd.DataFrame(rows_list)
		df = pd.concat([df, new_rows_df], ignore_index=True)
		
		df.to_excel(output_path, index=False)
		
		# Adjust column width
		workbook = load_workbook(output_path)
		worksheet = workbook.active
		
		for col in worksheet.columns:
			max_length = 0
			column = col[0].column_letter
			for cell in col:
				try:
					if len(str(cell.value)) > max_length:
						max_length = len(cell.value)
				except:
					pass
			adjusted_width = (max_length + 2)
			worksheet.column_dimensions[column].width = adjusted_width
		
		workbook.save(output_path)
	
	def extract_path_profile(self, path):
		infos = []
		techno = re.search(r'- (.+)', path)
		path_file_of_user = re.search(r'(.+)\\', path)
		
		if techno:
			infos.append(techno.group(1))
		if path_file_of_user:
			infos.append(path_file_of_user.group(1))
		return infos
	
	@staticmethod
	def verify_phone_number(phone_number):
		return bool(re.fullmatch(r"0\d{9}", phone_number))
	
	@staticmethod
	def verify_email(email):
		return bool(re.fullmatch(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", email))
	
	def verify_recruitment_file_format(self, file_path):
		df = pd.read_excel(file_path, header=None)
		contains_text = df.iloc[:, 2].apply(lambda x: "DOSSIER RECRUTEMENT FICHE ENTRETIEN" in str(x)).any()
		
		if contains_text:
			return True
		else:
			return False
	
	def verify_format_date(self, date):
		if isinstance(date, datetime):
			return date.strftime('%d/%m/%Y')
		if isinstance(date, str) and re.match(r'\d{1,2}/\d{1,2}/\d{2,4}', date):
			return date
		return None
	
	def format_date(self, date_string):
		if not date_string:
			return None
		if isinstance(date_string, datetime):
			return date_string.strftime('%d/%m/%y')
		if isinstance(date_string, float) or not isinstance(date_string, str):
			return None
		try:
			date_object = datetime.strptime(date_string, '%d/%m/%Y')
		except ValueError:
			try:
				date_object = datetime.strptime(date_string, '%d/%m/%y')
			except ValueError:
				return None
		return date_object.strftime('%d/%m/%y')
	
	def find_last_interview(self, dates):
		valid_dates = [datetime.strptime(date, "%d/%m/%y") for date in dates if date]
		if not valid_dates:
			return None
		return max(valid_dates).strftime("%d/%m/%y")
	
	def create_headers(self):
		headers = ["repartoire", "profil", "nom", "prenom", "tel", "email", "disponibilite"]
		for i in range(3):
			headers += [f"Date{i + 1}", f"Entretien{i + 1}"]
		headers.append("dernier entretien")
		return [header.upper() for header in headers]
	
	def is_file_not_open(self, file_path):
		try:
			with open(file_path, 'r'):
				pass
			return True
		except PermissionError:
			logging.error(f"Permission error while opening file {file_path}")
			return False