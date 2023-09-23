import gzip
import os
import re
import requests
import time
from pathlib import Path

import config

from PyQt5 import QtWidgets, QtGui
from bs4 import BeautifulSoup
from util import requestsSite
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

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
        
		model = QtGui.QStandardItemModel(len(products), 2)  # Adjust the number of columns accordingly
        
		for row, product in enumerate(products):
			for col, key in enumerate({"Result"}):  # This should be a list, not a set
				item = QtGui.QStandardItem(str(product))  # Convert 'product' to a string
				item.setEditable(True)
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
	def get_product_url(self, key_code):
		res = requests.get('https://shopping.bookoff.co.jp/search/keyword/' + key_code)

		if res.status_code == 200:
			page = BeautifulSoup(res.content, "html.parser")

			product_url = page.find(class_='productItem__link').get('href')
			price_element = page.find(class_='productItem__price').text

			price_element = price_element.replace(',', '')
			price = re.findall(r'\d+', price_element)

			product_url = "https://shopping.bookoff.co.jp/" + product_url
			price = price[0]

			self.products_list.append(product_url)
			self.draw_table(self.products_list)
		else:
			self.products_list.append("Not Scraped !")
			self.draw_table(self.products_list)


	def get_product_givensandcompany_data (page):
		if page != -1:
			img_url = page.find(id = "get-image-item-id").get('href')
			title = page.find("h1", {"data-hook": "product-title"}).text.strip()
			sku = page.find("div", {"data-hook": "sku"}).text.strip()
			price = page.find("span", {"data-hook": "formatted-primary-price"}).text.strip()
			description_element = page.find("div", {"data-hook": "info-section-description"})
			description_paragraph = description_element.find("p").text.strip() if description_element else ""
			numeric_value = float(price.replace("$", ""))
			discount_amount = numeric_value * 0.75
			product_data = {
				"title": title.replace("\n", ""),
				"sku": sku.replace("SKU: ", ""),
				"agb": "GC-" + sku.replace("SKU: ", ""),
				"vendor": sku.replace("SKU: ", ""),
				"cost": price,
				"price": "$" + str(round(discount_amount, 2)),
				"description": description_paragraph.replace("\n", ""),
				"status": "Active",
				"img_url": img_url,
				"image_command": "REPLACE",
				"image_position": "1",
				"tags": "GC",
				"tags_command": "REPLACE",
				"command": "MERGE",
				"variant_taxable": "TRUE",
				"vendor_status": "YES",
				"published": "TRUE",
				"published_scope": "GLOBAL",
				"variant_inventory_tracker": "",
				"variant_inventory_policy": "continue",
				"variant_fulfillment_service": "manual",
				"variant_inventory_qty": ""
			}
			return product_data
		return -1
	
	# gourmetgiftbaskets site scraping
	def get_product_gourmetgiftbaskets_data (page):
		if page != -1:
			img_url = ""
			easyzoom_div = page.find('div', class_ = 'easyzoom')
			if easyzoom_div:
				img_tag = easyzoom_div.find('img')
				if img_tag and 'src' in img_tag.attrs:
					img_url = img_tag['src']
				else:
					print("No img tag found within the easyzoom div.")
			else:
				print("No div with class 'easyzoom' found.")
			title = page.find(id = "ctl00_MainContentHolder_lblName").text.strip()
			price = page.find(id = "ctl00_MainContentHolder_lblSitePrice").text.strip()
			sku = page.find(id = "ctl00_MainContentHolder_lblSku").text.strip()
			description_paragraph = page.find(id = "ctl00_MainContentHolder_lblDescription").text.strip()
			numeric_value = float(price.replace("$", ""))
			discount_amount = numeric_value * 0.65
			product_data = {
				"title": title.replace("\n", ""),
				"sku": sku.replace("\n", ""),
				"agb": "GGB-" + sku.replace("\n", ""),
				"vendor": sku.replace("\n", ""),
				"cost": price,
				"price": "$" + str(round(discount_amount, 2)),
				"description": description_paragraph.replace("\n", ""),
				"status": "Active",
				"img_url": "https://www.gourmetgiftbaskets.com/" + img_url,
				"image_command": "REPLACE",
				"image_position": "1",
				"tags": "GGB",
				"tags_command": "REPLACE",
				"command": "MERGE",
				"variant_taxable": "TRUE",
				"vendor_status": "YES",
				"published": "TRUE",
				"published_scope": "GLOBAL",
				"variant_inventory_tracker": "",
				"variant_inventory_policy": "continue",
				"variant_fulfillment_service": "manual",
				"variant_inventory_qty": ""
			}
			return product_data
		return -1
	
	def fetch_gourmetgiftbaskets (self, txt_keyword):
		res = requests.get(
				"https://tdl46g.a.searchspring.io/api/search/search.json?ajaxCatalog=v3&resultsFormat=native&siteId=tdl46g&domain=https%3A%2F%2Fwww.gourmetgiftbaskets.com%2Fsearch.aspx%3Fkeyword%3D" + txt_keyword + "%23%2Fperpage%3A500&bgfilter.ss_hide=0&resultsPerPage=500&q=" + txt_keyword)
		print(
			"https://tdl46g.a.searchspring.io/api/search/search.json?ajaxCatalog=v3&resultsFormat=native&siteId=tdl46g&domain=https%3A%2F%2Fwww.gourmetgiftbaskets.com%2Fsearch.aspx%3Fkeyword%3D" + txt_keyword + "%23%2Fperpage%3A500&bgfilter.ss_hide=0&resultsPerPage=500&q=" + txt_keyword)
		if res.status_code == 200:
			json_data = res.json()
			if len(json_data["results"]) > 0:
				result_url = ""
				if len(json_data["results"]) == 1:
					result_url = json_data["singleResult"]
				else:
					result_url = json_data["results"][0]["url"]
				try:
					response = requestsSite(result_url, "", "get")
					result = ActionManagement.get_product_gourmetgiftbaskets_data(response)
					if result != -1:
						result["url"] = result_url
						if json_data["results"][0]['instock'] == "0":
							result["status"] = "Out of Stock"
						self.products.append(result)
						self.draw_table(self.products)
						return
				except Exception as e:
					print(e)
		result = {
			"title": "",
			"sku": txt_keyword,
			"agb": "GGB-" + txt_keyword,
			"vendor": txt_keyword,
			"cost": "",
			"price": "",
			"description": "",
			"status": "Disco'd",
			"url": "",
			"img_url": "",
			"command": "MERGE",
			"published": "FALSE",
			"published_scope": "GLOBAL",
			"variant_inventory_tracker": "shopify",
			"variant_inventory_policy": "deny",
			"variant_fulfillment_service": "manual",
			"variant_inventory_qty": "0"
		}
		self.products.append(result)
		self.draw_table(self.products)
	
	# winecountrygiftbaskets site scraping
	def fetch_winecountrygiftbaskets (self, txt_keyword):
		try:
			req_url = "https://www.winecountrygiftbaskets.com/product/giftbasketsearch?SKW=" + txt_keyword + "&searchbox=searchbox"
			res = requests.get(req_url, verify = False)
			html_content = res.content.decode('utf-8')
			stock = "Active"
			if "OUT OF STOCK" in html_content:
				stock = "Out of Stock"
			content = BeautifulSoup(res.content, "html.parser")
			result = ActionManagement.get_product_winecountrygiftbaskets_data(content)
			if result == -1:
				result = {
					"title": "",
					"sku": txt_keyword,
					"agb": "WC-" + txt_keyword,
					"vendor": txt_keyword,
					"cost": "",
					"price": "",
					"description": "",
					"status": "Disco'd",
					"url": "",
					"img_url": "",
					"command": "MERGE",
					"published": "FALSE",
					"published_scope": "GLOBAL",
					"variant_inventory_tracker": "shopify",
					"variant_inventory_policy": "deny",
					"variant_fulfillment_service": "manual",
					"variant_inventory_qty": "0"
				}
			else:
				result["url"] = req_url
				result["status"] = stock
			self.products.append(result)
			self.draw_table(self.products)
		except Exception as e:
			result = {
				"title": "",
				"sku": txt_keyword,
				"agb": "WC-" + txt_keyword,
				"vendor": txt_keyword,
				"cost": "",
				"price": "",
				"description": "",
				"status": "Disco'd",
				"url": "",
				"img_url": "",
				"command": "MERGE",
				"published": "FALSE",
				"published_scope": "GLOBAL",
				"variant_inventory_tracker": "shopify",
				"variant_inventory_policy": "deny",
				"variant_fulfillment_service": "manual",
				"variant_inventory_qty": "0"
			}
			self.products.append(result)
			self.draw_table(self.products)
			print("An error occurred:", e)
	
	def get_product_winecountrygiftbaskets_data (page):
		if page != -1:
			img_url = page.find('meta', {'property': 'og:image'}).get("content")
			url_parts = img_url.split("?")
			img_url = url_parts[0]
			title = page.find(id = "desc").text.strip()
			price_el = page.find(class_ = "p_price")
			price = price_el.findChild().text.strip()
			sku = page.find(class_ = "fcwcgrey fssmall align_right").text.strip()
			description_paragraph = page.find(class_ = "prt_detail_info_upsell").text.strip()
			numeric_value = float(price.replace("$", ""))
			discount_amount = numeric_value * 0.80
			description = description_paragraph.replace("\n", "")
			description = description.replace("DESCRIPTIONCONTENTSSHIPPING", "")
			product_data = {
				"title": title.replace("\n", ""),
				"sku": sku.replace("Item No: ", ""),
				"agb": "WC-" + sku.replace("Item No: ", ""),
				"vendor": sku.replace("Item No: ", ""),
				"cost": price,
				"price": "$" + str(round(discount_amount, 2)),
				"description": description,
				"status": "Active",
				"img_url": img_url,
				"image_command": "REPLACE",
				"image_position": "1",
				"tags": "WC",
				"tags_command": "REPLACE",
				"command": "MERGE",
				"variant_taxable": "TRUE",
				"vendor_status": "YES",
				"published": "TRUE",
				"published_scope": "GLOBAL",
				"variant_inventory_tracker": "",
				"variant_inventory_policy": "continue",
				"variant_fulfillment_service": "manual",
				"variant_inventory_qty": ""
			}
			return product_data
		return -1
	
	def extract_all_image_urls(self, data):
		urls = []
		for key, value in data.items():
			if isinstance(value, dict):
				urls.extend(self.extract_all_image_urls(value))  # Recursively search for URLs in nested dictionaries
			elif isinstance(value, list):
				for item in value:
					urls.extend(self.extract_all_image_urls(item))  # Recursively search for URLs in nested lists
			elif key == "url" and value:
				urls.append(value)  # Add the URL to the list
		return urls
	
	# capalbosonline site scraping
	def fetch_capalbosonline (self, txt_keyword):
		try:
			req_url = "https://www.capalbosonline.com/api/items?c=4666042&country=US&currency=USD&fieldset=search&language=en&q=" + txt_keyword
			print(req_url)
			res = requests.get(req_url, verify = False)
			if res.status_code == 200:
				json_data = res.json()
				if len(json_data["items"]) > 0:
					result = json_data["items"][0]
					result["url"] = req_url
					isinstock = result["isinstock"]
					
					stock = ""
					if isinstock == True:
						stock = "Active"
					else:
						stock = "Out of Sock"
					product_data = {
						"title": result['storedisplayname2'],
						"sku": result['itemid'],
						"agb": "CPB-" + result['itemid'],
						"vendor": result['itemid'],
						"cost": result['onlinecustomerprice_formatted'],
						"price": result['onlinecustomerprice_formatted'],
						"description": result['storedescription'],
						"status": stock,
						"url": "https://www.capalbosonline.com/" + result["urlcomponent"],
						"img_url": self.extract_all_image_urls(result['itemimages_detail'])[0]
					}
					self.products.append(product_data)
					self.draw_table(self.products)
					return
		except Exception as e:
			print("An error occurred:", e)
		product_data = {
			"title": "",
			"sku": txt_keyword,
			"agb": "CPB-" + txt_keyword,
			"vendor": txt_keyword,
			"cost": "",
			"price": "",
			"description": "",
			"status": "Disco'd",
			"url": "",
			"img_url": ""
		}
		self.products.append(product_data)
		self.draw_table(self.products)
	
	def fetch_giftbasketdropshipping (self, txt_keyword):
		chrome_options = Options()
		chrome_options.add_argument("--headless")
		chrome_options.add_argument("--disable-gpu")
		chrome_options.add_argument('--log-level=3')
		driver = webdriver.Chrome(options = chrome_options)
		try:
			driver.get("https://www.giftbasketdropshipping.com/cgi-bin/accountedit2.pl")
			username = driver.find_element(By.NAME, "username1")
			username.send_keys("sales@amgiftbaskets.com")
			password = driver.find_element(By.NAME, "authen1")
			password.send_keys("ezpass")
			submit_button = driver.find_element(By.CSS_SELECTOR, "input[type='submit']")
			submit_button.click()
			time.sleep(0.5)
			input_element = driver.find_element(By.CSS_SELECTOR, "input[value='Enter Order']")
			input_element.click()
			time.sleep(0.5)
			javascript_code = "return customarray.filter(e => e.indexOf('" + txt_keyword + "') != -1);"
			value = driver.execute_script(javascript_code)
			arr = value[0].split("-")
			product_data = {
				"title": arr[1],
				"sku": arr[0],
				"agb": "GBD-" + arr[0],
				"vendor": arr[0],
				"cost": arr[3],
				"price": arr[3],
				"description": value[0],
				"status": "Active",
				"url": "",
				"img_url": "",
				"image_command": "REPLACE",
				"image_position": "1",
				"tags": "GBD",
				"tags_command": "REPLACE",
				"command": "MERGE",
				"variant_taxable": "TRUE",
				"vendor_status": "YES",
				"published": "TRUE",
				"published_scope": "GLOBAL",
				"variant_inventory_tracker": "",
				"variant_inventory_policy": "continue",
				"variant_fulfillment_service": "manual",
				"variant_inventory_qty": ""
			}
			self.products.append(product_data)
			self.draw_table(self.products)
			return
		except Exception as e:
			print("An error occurred:", str(e))
		finally:
			# Close the WebDriver
			driver.quit()
		product_data = {
			"title": "",
			"sku": txt_keyword,
			"agb": "GBD-" + txt_keyword,
			"vendor": txt_keyword,
			"cost": "",
			"price": "",
			"description": "",
			"status": "Disco'd",
			"url": "",
			"img_url": "",
			"command": "MERGE",
			"published": "FALSE",
			"published_scope": "GLOBAL",
			"variant_inventory_tracker": "shopify",
			"variant_inventory_policy": "deny",
			"variant_fulfillment_service": "manual",
			"variant_inventory_qty": "0"
		}
		self.products.append(product_data)
		self.draw_table(self.products)
