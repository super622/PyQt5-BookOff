import gzip
import os
import re
import requests

import config

from pathlib import Path
from PyQt5 import QtWidgets, QtGui
from bs4 import BeautifulSoup

class ActionManagement:
	products_list = []
	file_path = ''
	
	def __init__ (self, main_window):
		self.main_window = main_window
		self.refresh_token = config.REFRESH_TOKEN
		self.client_id = config.CLIENT_ID
		self.client_secret = config.CLIENT_SECRET
		self.access_token = ''
		self.api_url = "https://sellingpartnerapi-fe.amazon.com"
	
	# drow table
	def draw_table(self, products):
		table = self.main_window.findChild(QtWidgets.QTableView, "tbl_dataview")
        
		model = QtGui.QStandardItemModel(len(products), 0)  # Adjust the number of columns accordingly
		model.setHorizontalHeaderLabels(['URL'])

		for row, product in enumerate(products):
			for col, key in enumerate({"result"}):  # This should be a list, not a set
				item = QtGui.QStandardItem(str(product))  # Convert 'product' to a string
				item.setEditable(False)
				model.setItem(row, col, item)
        
		table.setModel(model)
		header = table.horizontalHeader()
		font = QtGui.QFont()
		font.setBold(True)
		header.setFont(font)

	# get Access Token
	def get_access_token(self):
		url = "https://api.amazon.co.jp/auth/o2/token"
		payload = {
            "grant_type": "refresh_token",
            "refresh_token": self.refresh_token,
            "client_id": self.client_id,
            "client_secret": self.client_secret,
        }
		response = requests.post(url, data=payload)
		access_token = response.json().get("access_token")

		if access_token:
			return access_token
		else:
			return ''

	# get report document id
	def get_report_document_id(self, access_token):
		url = f"{self.api_url}/reports/2021-06-30/reports"
		headers = {"Accept": "application/json", "x-amz-access-token": f"{access_token}"}
		params = {"reportTypes": "GET_MERCHANT_LISTINGS_ALL_DATA"}
		response = requests.get(url, headers=headers, params=params)

		if response.status_code == 200:
			reports = response.json()
			reportDocumentId = reports["reports"][0]["reportDocumentId"]
			return reportDocumentId
		else:
			return ''
		
	# get report document url
	def get_report_gz_url(self, report_document_id, access_token):
		url = f"{self.api_url}/reports/2021-06-30/documents/{report_document_id}"
		headers = {"Accept": "application/json", "x-amz-access-token": f"{access_token}"}
		response = requests.get(url, headers=headers)

		if response.status_code == 200:
			report_doc = response.json()["url"]
			return report_doc
		else:
			return ''

	# donwload gz file
	def download_file(self, url, filepath):
		try:
			document_folder = Path.home() / "Documents"
			amazon_folder = document_folder / "Amazon"
			if not amazon_folder.exists():
				amazon_folder.mkdir(parents=True)
			
			filepath = amazon_folder / filepath

			response = requests.get(url, stream=True)
			response.raise_for_status()

			with open(filepath, "wb") as file:
				for chunk in response.iter_content(chunk_size=1024):
					if chunk:
						file.write(chunk)
			return True
		except requests.exceptions.RequestException as e:
			return False
		except IOError as e:
			return False
		except Exception as e:
			return False

	# unzip gz file
	def unzip_file(self, gz_filepath, extracted_filepath):
		try:
            # os.chmod(extracted_filepath, 0o666)
			document_folder = Path.home() / "Documents"
			amazon_folder = document_folder / "Amazon"
			filepath = amazon_folder / gz_filepath
			extracted_filepath = amazon_folder / extracted_filepath

			with gzip.open(filepath, 'rb') as gz_file:
				with open(extracted_filepath, 'wb') as output_file:
					output_file.write(gz_file.read())
			os.remove(filepath)
			self.file_path = extracted_filepath
			return ''
		except PermissionError as e:
			product_data = {
				'result': f'Permission error: {e}'
			}
			# return product_data
			return f'Permission error: {e}'
		except Exception as e:
			product_data = {
				'result': f'An error occurred: {e}'
			}
			# return product_data
			return f'An error occurred: {e}'

	# get product total count
	def get_content_from_file(self, filepath):
		try:
			i = 0
			cnt = 0
			with open(filepath, 'r', encoding='utf-8') as file:
				for line in file.readlines():
					line = line.strip().split(',')
					fields = line[0].split('\t')
					
					if i >= 1 and len(fields) >= 2 and fields[-1] == 'Active' and fields[-2] == '送料無料(お急ぎ便無し)':
						cnt += 1
					i += 1
					
			result = {
                'filepath': filepath,
                'total': cnt,
            }
			return result
		except FileNotFoundError:
			return ''
		except Exception as e:
			return ''

	# get product list from amazon
	def product_list_download_from_amazon(self):
		self.access_token = self.get_access_token()
		if(self.access_token == ''):
			product_data = {
				"result": 'アクセストークンを取得できませんでした。'
			}
			# self.products_list.append(product_data)
			return 'アクセストークンを取得できませんでした。'
		
		report_document_id = self.get_report_document_id(self.access_token)
		if(report_document_id == ''):
			product_data = {
				'result': 'report document idを取得できません。'
			}
			# self.products_list.append(product_data)
			return 'report document idを取得できません。'
		
		report_document_url = self.get_report_gz_url(report_document_id, self.access_token)
		if(report_document_url == ''):
			product_data = {
				'result': 'リストファイルのパスを取得できません。'
			}
			# self.products_list.append(product_data)
			return 'リストファイルのパスを取得できません。'
		
		download_flag = self.download_file(report_document_url, f"{report_document_id}.gz")
		if(download_flag == False):
			product_data = {
				'result': 'ファイルをダウロドしていた途中にエラーが発生しました。'
			}
			# self.products_list.append(product_data)
			return 'ファイルをダウロドしていた途中にエラーが発生しました。'
		
		unzip_flag = self.unzip_file(f"{report_document_id}.gz", f"{report_document_id}")
		if(unzip_flag != ''):
			# self.products_list.append(unzip_flag)
			return unzip_flag
		
		result = self.get_content_from_file(f"{report_document_id}")
		if(result == ''):
			return '無効なファイルです'
		
		return result

	# get product list from file
	def read_product_list_from_file(self):
		try:
			i = 0
			cnt = 0
			with open(self.file_path, 'r', encoding='utf-8') as file:
				for line in file.readlines():
					line = line.strip().split(',')
					fields = line[0].split('\t')
					
					if i >= 1 and len(fields) >= 2 and fields[-1] == 'Active' and fields[-2] == '送料無料(お急ぎ便無し)':
						print('')
					i += 1
			return 'success'
		except FileNotFoundError:
			return ''
		except Exception as e:
			return ''

	# get product url
	def get_product_url(self, key_code, other_price):
		res = requests.get('https://shopping.bookoff.co.jp/search/keyword/' + key_code)

		if res.status_code == 200:
			page = BeautifulSoup(res.content, "html.parser")

			product_url = page.find(class_='productItem__link').get('href')
			price_element = page.find(class_='productItem__price').text
			stock_element = page.find_all(class_="productInformation__stock")

			price_element = price_element.replace(',', '')
			price = re.findall(r'\d+', price_element)

			product_url = "https://shopping.bookoff.co.jp" + product_url
			price = price[0]

			print(len(stock_element))

			# if other_price > price:
			# 	percent = price / (other_price / 100)
			# 	flag = False

			# 	if((100 - percent) >= 35):
			# 		flag = True
				
			# 	self.products_list.append(product_url)
			# 	self.draw_table(self.products_list)

			self.products_list.append(product_url)
			self.draw_table(self.products_list)
		else:
			self.products_list.append("Not Scraped !")
			self.draw_table(self.products_list)
