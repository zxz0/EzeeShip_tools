"""
	Todo:
		save temp file
	Date:
		02/16/2019 (parse the excel file with correct format)
		02/11/2019 (jsonify problem. keep it simple)
		02/09/2019 (design class and relations: InfoGetter only get info from xls, EstimateRates get rate info, with other "POJO")
		02/01/2019 (created, 
			test example order rate estimate,
			test real order rate estimate to make sure rate is right,
			structure: main, handle .ini file, log system)
	Author: Zixuan Zhang
	Function: get best price among several shipping method for different orders according to several rules, output to csv file as the original order
	Usage: python compare.py products.csv
"""

import requests
import json
import os
import configparser
import logging
import sys
import xlrd
import datetime

'''class Address:
	def __init__(self, country_code = 'US', state_code = 'CA', city = 'Rancho Cucamonga', address_line1 = '9370 7Th St, Ste G.', zip_code = '91730'):
		self.info = {}
		self.info['countryCode'] = country_code
		self.info['stateCode'] = state_code
		self.info['city'] = city
		self.info['addressLine1'] = address_line1
		self.info['zipCode'] = zip_code

class Parcel:
	def __init__(self):
		self.info = {}

	def __init__(self, length, width, height, weight, distance_unit = 'in', package_code = 'your_package', package_num = 1, mass_unit = 'lb'):
		self.info = {}
		self.info['lenth'] = length
		self.info['width'] = width
		self.info['height'] = height
		self.info['weight'] = weight
		self.info['packageCode'] = packageCode
		self.info['packageNum'] = package_num
		self.info['massUnit'] = mass_unit'''

class Validator:
	def __init__(self):
		# address can be validate later using API
		self.country_code = ('BD', 'BE', 'BF', 'BG', 'BA', 'BB', 'WF', 'BL', 'BM', 'BN', 'BO', 'BH', 'BI', 'BJ', 'BT', 'JM', 'BV', 'BW', 'WS', 'BQ', 'BR', 'BS', 'JE', 'BY', 'BZ', 'RU', 'RW', 'RS', 'TL', 'RE', 'TM', 'TJ', 'RO', 'TK', 'GW', 'GU', 'GT', 'GS', 'GR', 'GQ', 'GP', 'JP', 'GY', 'GG', 'GF', 'GE', 'GD', 'GB', 'GA', 'SV', 'GN', 'GM', 'GL', 'GI', 'GH', 'OM', 'TN', 'JO', 'HR', 'HT', 'HU', 'HK', 'HN', 'HM', 'VE', 'PR', 'PS', 'PW', 'PT', 'SJ', 'PY', 'IQ', 'PA', 'PF', 'PG', 'PE', 'PK', 'PH', 'PN', 'PL', 'PM', 'ZM', 'EH', 'EE', 'EG', 'ZA', 'EC', 'IT', 'VN', 'SB', 'ET', 'SO', 'ZW', 'SA', 'ES', 'ER', 'ME', 'MD', 'MG', 'MF', 'MA', 'MC', 'UZ', 'MM', 'ML', 'MO', 'MN', 'MH', 'MK', 'MU', 'MT', 'MW', 'MV', 'MQ', 'MP', 'MS', 'MR', 'IM', 'UG', 'TZ', 'MY', 'MX', 'IL', 'FR', 'IO', 'SH', 'FI', 'FJ', 'FK', 'FM', 'FO', 'NI', 'NL', 'NO', 'NA', 'VU', 'NC', 'NE', 'NF', 'NG', 'NZ', 'NP', 'NR', 'NU', 'CK', 'XK', 'CI', 'CH', 'CO', 'CN', 'CM', 'CL', 'CC', 'CA', 'CG', 'CF', 'CD', 'CZ', 'CY', 'CX', 'CR', 'CW', 'CV', 'CU', 'SZ', 'SY', 'SX', 'KG', 'KE', 'SS', 'SR', 'KI', 'KH', 'KN', 'KM', 'ST', 'SK', 'KR', 'SI', 'KP', 'KW', 'SN', 'SM', 'SL', 'SC', 'KZ', 'KY', 'SG', 'SE', 'SD', 'DO', 'DM', 'DJ', 'DK', 'VG', 'DE', 'YE', 'DZ', 'US', 'UY', 'YT', 'UM', 'LB', 'LC', 'LA', 'TV', 'TW', 'TT', 'TR', 'LK', 'LI', 'LV', 'TO', 'LT', 'LU', 'LR', 'LS', 'TH', 'TF', 'TG', 'TD', 'TC', 'LY', 'VA', 'VC', 'AE', 'AD', 'AG', 'AF', 'AI', 'VI', 'IS', 'IR', 'AM', 'AL', 'AO', 'AQ', 'AS', 'AR', 'AU', 'AT', 'AW', 'IN', 'AX', 'AZ', 'IE', 'ID', 'UA', 'QA', 'MZ')
		self.state_code = ('AL', 'AK', 'AS', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DE', 'DC', 'FM', 'FL', 'GA', 'GU', 'HI', 'ID', 'IL', 'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 'MH', 'MD', 'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ', 'NM', 'NY', 'NC', 'ND', 'MP', 'OH', 'OK', 'OR', 'PW', 'PA', 'PR', 'RI', 'SC', 'SD', 'TN', 'TX', 'UT', 'VT', 'VI', 'VA', 'WA', 'WV', 'WI', 'WY')


class Order:
	def __init__(self):
		self.request_dict = {}
		self.request_dict['isTest'] = False
		self.request_dict['carrierCode'] = ''
		self.request_dict['serviceCode'] = ''
		self.request_dict['from'] = {} 
		self.request_dict['to'] = {}
		self.request_dict['parcels'] = [{}]

class PriceEstimater():
	def get_estimated_rate():
		url = 'https://ezeeship.com/api/ezeeship-openapi/shipment/estimateRate'

		headers = {'Authorization': api_key, 'Content-Type': 'application/json'}

		payload = ''''''

		self.small = ['fedex_smartpost', 'usps_priority', 'ups_ground']   # 14" - 20"
		self.big = ['usps_priority', 'ups_ground']   # 24"~
		self.transformer = ['usps_priority']
		self.trans_3kv = ['usps_first_class']
		#json.dumps(test)
		response = requests.post(url, headers = headers, data = payload)
		response.json()

class XlsReader():
	def __init__(self, input_file, head_to_ignore = 1, sheet_number = 0):
		self.input_file = input_file
		self.orders = []
		self.head_to_ignore = head_to_ignore
		self.sheet_number = sheet_number
 
	"""
	Parse input file, get orders with info

	Return:
	- orders: the orders with information extracted from the input file
	"""
	def parse(self):
		orders = []
		workbook = xlrd.open_workbook(self.input_file)
		sheetbooksheet = workbook.sheet_by_index(self.sheet_number)

		for row_idx in range(self.head_to_ignore, sheetbooksheet.nrows):	# Ignore the head
			current_row = sheetbooksheet.row(row_idx)
			current_order = Order()

			# get info from the columns with correct format
			reference = current_row[2].value.strip()

			sender_country = current_row[5].value.strip()
			sender_address = current_row[7].value.strip()
			sender_city = current_row[9].value.strip()
			sender_state = current_row[10].value.strip()
			sender_zipcode = current_row[11].value.strip() if current_row[11].ctype == xlrd.book.XL_CELL_TEXT else str(int(current_row[11].value))

			recipient_country = current_row[16].value.strip()
			recipient_address = current_row[18].value.strip()
			recipient_city = current_row[20].value.strip()
			recipient_state = current_row[21].value.strip()
			recipient_zipcode = current_row[22].value.strip() if current_row[22].ctype == xlrd.book.XL_CELL_TEXT else str(int(current_row[22].value))

			is_cod = current_row[26].value
			cod_amount = current_row[27].value if current_row[27].ctype == xlrd.book.XL_CELL_NUMBER else 0

			length = str(current_row[29].value)
			width = str(current_row[30].value)
			height = str(current_row[31].value)
			weight = str(current_row[32].value)

			insurance_amount = current_row[34].value if current_row[34].ctype == xlrd.book.XL_CELL_NUMBER else 0

			# structurize
			sender_info = {}
			sender_info['countryCode'] = sender_country
			sender_info['stateCode'] = sender_state
			sender_info['city'] = sender_city
			sender_info['addressLine1'] = sender_address
			sender_info['zipCode'] = sender_zipcode

			recipient_info = {}
			recipient_info['countryCode'] = recipient_country
			recipient_info['stateCode'] = recipient_state
			recipient_info['city'] = recipient_city
			recipient_info['addressLine1'] = recipient_address
			recipient_info['zipCode'] = recipient_zipcode

			print(sender_info)
			print(recipient_info)

			parcel_info = {}
			parcel_info['packageNum'] = 1
			parcel_info['length'] = length
			parcel_info['width'] = width
			parcel_info['height'] = height
			parcel_info['distanceUnit'] = 'in'
			parcel_info['weight'] = weight
			parcel_info['massUnit'] = 'lb'
			parcel_info['packageCode'] = ''
			extra_info = {}
			extra_info['insurance'] = insurance_amount if 
			extra_info['isCod'] = True if is_cod else False
			extra_info['codAmount'] = 0
			extra_info['paymentMethod'] = 'any'
			extra_info['dryIceWeight'] = 0
			parcel_info['extra'] = extra_info

			# populate the order instance
			current_order.request_dict['from'] = sender_info
			current_order.request_dict['to'] = recipient_info
			current_order.request_dict['parcels'][0] = parcel_info

			# validate datatype, make sure it meets requested format
			# current_order.validate()
			# validate_address(current_order)
			orders.append(current_order)
			logging.info('get request dictionary: {request_dict}'.format(request_dict = current_order.request_dict))

		return orders

def set_logger():
	folder = './logs'
	if not os.path.isdir(folder):
		os.makedirs(folder)
	module_name = os.path.splitext(os.path.basename(__file__))[0]
	date = datetime.datetime.now().strftime("%Y-%m-%d")
	log_file = '{folder}/[{module_name}]{date}.log'.format(folder = folder, module_name = module_name, date = date)

	logging.basicConfig(level=logging.DEBUG, format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s', datefmt='%a, %d %b %Y %H:%M:%S', filename=log_file)	# default file mode: a
	console=logging.StreamHandler()
	console.setLevel(logging.INFO)
	formatter=logging.Formatter('%(name)-12s: %(levelname)-8s %(message)s')
	console.setFormatter(formatter)
	logging.getLogger('').addHandler(console)

def main():
	# Initinalization
	set_logger()
	
	# Get API key from config file
	api_key = ''
	config_file = 'config.ini'

	if os.path.isfile(config_file):
		config = configparser.ConfigParser()
		config.read(config_file)
		api_key = config.get('Keys', 'api_key')
		logging.debug('Get API key: {api_key}'.format(api_key = api_key))
	else:
		logging.error('config file: {config_file} not exist!'.format(config_file = config_file))
		exit(2)

	# Get order info from xls file
	xls_file = 'input.xls'
	orders = []
	if os.path.isfile(xls_file):
		xls_reader = XlsReader(xls_file)
		orders = xls_reader.parse()
	else:
		logging.error('order file: {xls_file} not exist!'.format(xls_file = xls_file))
		exit(2)

	# Get price for orders

	# Write results to xls file


if __name__ == "__main__":
	main()
