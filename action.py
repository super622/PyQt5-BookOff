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
	document_folder = Path.home() / "Documents"
	amazon_folder = document_folder / "Amazon"
	
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
        
		model = QtGui.QStandardItemModel(len(products), 5)  # Adjust the number of columns accordingly
		model.setHorizontalHeaderLabels(["JAN", "URL", "在庫", "サイト価格", "Amazonの価格", "価格差"])

		for row, product in enumerate(products):
			for col, key in enumerate(['jan', 'url', 'stock', 'site_price', 'amazon_price', 'price_status']):  # This should be a list, not a set
				item = QtGui.QStandardItem(product.get(key, ""))  # Convert 'product' to a string
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
	def download_report_document_file(self, url, filepath):
		try:
			if not self.amazon_folder.exists():
				self.amazon_folder.mkdir(parents=True)
			
			filepath = self.amazon_folder / filepath

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
	def unzip_report_document_file(self, gz_filepath, extracted_filepath):
		try:
            # os.chmod(extracted_filepath, 0o666)
			filepath = self.amazon_folder / gz_filepath
			extracted_filepath = self.amazon_folder / extracted_filepath

			with gzip.open(filepath, 'rb') as gz_file:
				with open(extracted_filepath, 'wb') as output_file:
					output_file.write(gz_file.read())
			os.remove(filepath)
			return ''
		except PermissionError as e:
			return f'Permission error: {e}'
		except Exception as e:
			return f'An error occurred: {e}'

	# get product total count
	def get_content_from_file(self, origin_filepath):
		try:
			i = 0
			cnt = 0
			filepath = self.amazon_folder / origin_filepath
			with open(filepath, 'r', encoding='utf-8') as file:
				for line in file.readlines():
					line = line.strip().split(',')
					fields = line[0].split('\t')
					
					if i >= 1 and len(fields) >= 2 and fields[-1] == 'Active' and fields[-2] == '送料無料(お急ぎ便無し)':
						cnt += 1
					i += 1
					
			result = {
                'filepath': origin_filepath,
                'total': cnt,
            }
			return result
		except FileNotFoundError:
			return ''
		except Exception as e:
			return ''

	# get Jan code by asin code
	def get_jan_code_by_asin(self, temp_asin_arr, asins):
		url = "https://sellingpartnerapi-fe.amazon.com/catalog/2022-04-01/items"
		headers = {
            "x-amz-access-token": self.access_token,
            "Accept": "application/json"
        }
		params = {
            "marketplaceIds": config.MAKETPLACEID,
            "sellerId": config.SELLERID,
            "includedData": "identifiers,attributes,salesRanks",
            "identifiersType": "ASIN",
            "identifiers": asins
        }
		response = requests.get(url, headers=headers, params=params)
		result_arr = [['', '', '', '']] * len(temp_asin_arr) # 1. jan code, 2. category, 3. ranking, 4. price
		if response.status_code == 200:
			json_response = response.json()
			if (json_response['items']):
				for product in json_response['items']:
					for i in range(len(temp_asin_arr)):
						if(temp_asin_arr[i] == product['asin']):
							result_arr[i][0] = product['identifiers'][0]['identifiers'][0]['identifier']
							result_arr[i][1] = product['salesRanks'][0]['displayGroupRanks'][0]['title']
							result_arr[i][2] = product['salesRanks'][0]['displayGroupRanks'][0]['rank']
							result_arr[i][3] = product['attributes']['list_price'][0]['value']
				return result_arr
			else:
				return ''
		else:
			return ''

	# convert array to str
	def convert_array_to_string(self, arr):
		result_str = ''
		for i in range(len(arr)):
			if(i == 0):
				result_str += arr[i]
			else:
				result_str += f",{arr[i]}"
				
		return result_str

	# get product list from amazon
	def product_list_download_from_amazon(self):
		self.access_token = self.get_access_token()
		if(self.access_token == ''):
			return 'アクセストークンを取得できませんでした。'
		
		report_document_id = self.get_report_document_id(self.access_token)
		if(report_document_id == ''):
			return 'report document idを取得できません。'
		
		report_document_url = self.get_report_gz_url(report_document_id, self.access_token)
		if(report_document_url == ''):
			return 'リストファイルのパスを取得できません。'
		
		download_flag = self.download_report_document_file(report_document_url, f"{report_document_id}.gz")
		if(download_flag == False):
			return 'ファイルをダウロドしていた途中にエラーが発生しました。'
		
		unzip_flag = self.unzip_report_document_file(f"{report_document_id}.gz", f"{report_document_id}")
		if(unzip_flag != ''):
			return unzip_flag
		
		result = self.get_content_from_file(f"{report_document_id}")
		if(result == ''):
			return '無効なファイルです'
		
		return result

	# get product list from file
	def read_product_list_from_file(self, filepath):
		try:
			i = 0
			filepath = self.amazon_folder / filepath
			with open(filepath, 'r', encoding='utf-8') as file:
				for line in file.readlines():
					line = line.strip().split(',')
					fields = line[0].split('\t')
					
					if i >= 1 and len(fields) >= 2 and fields[-1] == 'Active' and fields[-2] == '送料無料(お急ぎ便無し)':
						self.products_list.append(fields[1])
					i += 1
			return 'success'
		except FileNotFoundError as e:
			return e
		except Exception as e:
			return e
	
	# get product info
	def get_product_info_by_product_list(self, position):
		cnt = 0
		asin_arr = []
		asins = ''
		for asin in self.products_list:
			if(position >= cnt):
				if(position == cnt):
					self.access_token = self.get_access_token()
				
				if(cnt == (position + 20)):
					break
				
				asin_arr.append(asin)
			cnt += 1

		asins = self.convert_array_to_string(asin_arr)
		result = self.get_jan_code_by_asin(asin_arr, asins)
		return result

	# get product url
	def get_product_url(self, product):
		key_code = product[0]
		category = product[1]
		ranking = product[2]
		other_price = product[3]
		ranking_flag = 150000
		
		if(category == 'game'):
			ranking_flag = 150000
		if(category == 'cd'):
			ranking_flag = 80000

		if ranking <= ranking_flag:
			res = requests.get('https://shopping.bookoff.co.jp/search/keyword/' + key_code)

			if res.status_code == 200:
				page = BeautifulSoup(res.content, "html.parser")

				product_url = page.find(class_='productItem__link').get('href')
				price_element = page.find(class_='productItem__price').text
				stock_element = page.find_all(class_="productItem__stock--alert")

				price_element = price_element.replace(',', '')
				price = re.findall(r'\d+', price_element)
				stock = ''

				product_url = "https://shopping.bookoff.co.jp" + product_url
				price = int(price[0])

				if len(stock_element) > 0:
					stock = '在庫なし'
				else:
					stock = ''
				
				if other_price > price:
					percent = price / (other_price / 100)
					price_status = ''

					if((100 - percent) >= 35):
						price_status = 'T'

				product_data = {
					'jan': key_code,
					'url': product_url,
					'stock': stock,
					'site_price': str(price),
					'amazon_price': str(other_price),
					'price_status': price_status
				}
				self.products_list.append(product_data)
				self.draw_table(self.products_list)
			# else:
				# self.products_list.append("Not Scraped !")
				# self.draw_table(self.products_list)
