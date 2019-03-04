"""
	Todo:
		save temp file, config file adoption, (test) structurize, modulization, packing (to exe), validate address before, automatically revise, sort by key
	Author: Zixuan Zhang
	Function: 
		parse xls to get order info,
		get best price among several shipping method according to several rules, 
		output to xls file following the original order
	Usage: python compare.py products.csv
"""

import requests
import json
import os
import configparser
import logging
import sys
import xlrd
import xlwt
import datetime

class RequestError(Exception):
    """Exception raised for errors during the request.

    Attributes:
        request:
        	request producing the error
        response:
        	response given
        message:
        	error message, explanation of the error
    """

    def __init__(self, request, response, message):
        self.request = request
        self.response = response
        self.message = message

class Order:
	def __init__(self):
		self.request_dict = {}
		self.request_dict['isTest'] = False
		self.request_dict['carrierCode'] = ''
		self.request_dict['serviceCode'] = ''
		self.request_dict['from'] = {}
		self.request_dict['to'] = {}
		self.request_dict['parcels'] = [{}]
		self.shipping_rates = {}	# {'serviceCode': rate}
		self.reference = ''

	def populate_other_properties(self):
		# Handle special address
		to_address = self.request_dict['to']
		if to_address['zipCode'] == '41025' and to_address['addressLine1'] == '1850 Airport Exchange Blvd #200' and to_address['city'] == 'Erlanger' and to_address['stateCode'] == 'KY' and to_address['countryCode'] == 'US':
			self.shipping_rates['fedex_ground'] = None
			self.shipping_rates['usps_priority'] = None
			return

		# Handle transformers
		if 'kv' in self.reference.lower():
			if '3kv' in self.reference.lower(): 	# 3kv
				self.shipping_rates['usps_first'] = None
				self.request_dict['parcels'][0]['packageCode'] = 'thick_envelope'
			else: 	# not 3kv
				self.shipping_rates['usps_priority'] = None
				if 'x' in self.reference.lower(): 	# 2+ transformers, not 3kv
					self.request_dict['parcels'][0]['packageCode'] = 'medium_flat_rate_box'
				else: 	# just 1 transformer, not 3kv
					self.request_dict['parcels'][0]['packageCode'] = 'flat_rate_envelope'
			return

		# Handle LED
		if 'led' in self.reference.lower():
			self.shipping_rates['fedex_smart_post'] = None
			self.shipping_rates['usps_priority'] = None
			# self.shipping_rates['ups_ground'] = None
			return

		# Handle neon sign
		size = int(self.reference[1:3])
		if size == 17 and self.request_dict['parcels'][0]['length'] == 18:	# 17", paper box
			self.shipping_rates['fedex_smart_post'] = None
		else:
			# if not 'po box' in to_address['addressLine1'].lower():
			# 	self.shipping_rates['ups_ground'] = None
			if size <= 17:
				self.shipping_rates['fedex_smart_post'] = None
			elif size <= 20:
				self.shipping_rates['fedex_smart_post'] = None
				self.shipping_rates['usps_priority'] = None
				# self.shipping_rates['fedex_home_delivery'] = None	# residential
				# self.shipping_rates['fedex_ground'] = None			# commercial
			elif size >= 24 and size < 32: 	# nothing between 20 and 24?
				self.shipping_rates['fedex_home_delivery'] = None
				self.shipping_rates['fedex_ground'] = None
				self.shipping_rates['usps_priority'] = None
			else: 	# size >= 32
				self.shipping_rates['usps_parcel_select'] = None

	def set_residential_commercial_method(self, is_residential_address):
		if is_residential_address:
			self.shipping_rates['fedex_home_delivery'] = None 	# residential
			self.request_dict['to']['isResidential'] = True
			self.request_dict['to']['isValie'] = True
		else:
			self.shipping_rates['fedex_ground'] = None			# commercial

def is_residential(api_key, payload):
	url = 'https://ezeeship.com/api/ezeeship-openapi/address/validate'

	headers = {'Authorization': api_key, 'Content-Type': 'application/json'}

	logging.debug("Send request: {request}".format(request = payload))
	response = requests.post(url, headers = headers, data = payload)
	json_response = response.json()
	logging.debug("Received response: {response}".format(response = json_response))
	if json_response['result'] == 'OK':
		return json_response['data']['isResidential']
	elif json_response['result'] == 'ERR':
		raise RequestError(payload, json_response, json_response['message'])
	# other errors should have been handled by requests

def get_estimated_rate(api_key, payload):
	url = 'https://ezeeship.com/api/ezeeship-openapi/shipment/estimateRate'

	headers = {'Authorization': api_key, 'Content-Type': 'application/json'}

	logging.debug("Send request: {request}".format(request = payload))
	response = requests.post(url, headers = headers, data = payload)
	json_response = response.json()
	logging.debug("Received response: {response}".format(response = json_response))
	if json_response['result'] == 'OK':
		return json_response['data']['rate']
	elif json_response['result'] == 'ERR':
		raise RequestError(payload, json_response, json_response['message'])
	# other errors should have been handled by requests

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

			parcel_info = {}
			parcel_info['packageNum'] = 1
			parcel_info['length'] = length
			parcel_info['width'] = width
			parcel_info['height'] = height
			parcel_info['distanceUnit'] = 'in'
			parcel_info['weight'] = weight
			parcel_info['massUnit'] = 'lb'
			parcel_info['packageCode'] = 'your_package'
			extra_info = {}
			extra_info['insurance'] = insurance_amount
			extra_info['isCod'] = True if is_cod else False
			extra_info['codAmount'] = 0
			extra_info['paymentMethod'] = 'any'
			extra_info['dryIceWeight'] = 0
			parcel_info['extra'] = extra_info

			# populate the order instance
			current_order.request_dict['from'] = sender_info
			current_order.request_dict['to'] = recipient_info
			current_order.request_dict['parcels'][0] = parcel_info
			current_order.reference = reference

			# validate datatype, make sure it meets requested format
			# current_order.validate()
			# validate_address(current_order)
			orders.append(current_order)
			logging.info('finished parsing line {line_num}'.format(line_num = row_idx + 1))
			logging.debug('get request dictionary: {request_dict}'.format(request_dict = current_order.request_dict))

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
		logging.info('Geting API key...')
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
		logging.info('Parsing xls file...')
		xls_reader = XlsReader(xls_file)
		orders = xls_reader.parse()
	else:
		logging.error('order file: {xls_file} not exist!'.format(xls_file = xls_file))
		exit(2)

	# Get rates for orders
	logging.info('Estimating rates...')
	for index, order in enumerate(orders):
		logging.info('Estimating rates for row: {row_index}...'.format(row_index = index + 1))
		min_price = 500
		min_shipping = ''
		order.populate_other_properties()
		try:
			order.set_residential_commercial_method(is_residential(api_key, json.dumps(order.request_dict['to'])))
		except RequestError as re:
			logging.error('cannot get rate for row {row_num}: {reason}'.format(row_num = index + 1, reason = re.message))

		for shipping_method in order.shipping_rates.keys():
			order.request_dict['serviceCode'] = shipping_method
			order.request_dict['carrierCode'] = shipping_method.split('_')[0]
			try:
				price = get_estimated_rate(api_key, json.dumps(order.request_dict))
				order.shipping_rates[shipping_method] = price
				if price < min_price:
					min_price = price
					min_shipping = shipping_method
			except RequestError as re:
				logging.error('cannot get rate for row {row_num}: {reason}'.format(row_num = index + 1, reason = re.message))
				order.shipping_rates[shipping_method] = re.message
		order.request_dict['serviceCode'] = min_shipping
		order.request_dict['carrierCode'] = min_shipping

	# Write results to xls file
	logging.info('Writing results...')
	workbook = xlwt.Workbook()
	sheet = workbook.add_sheet('Sheet 1')
	for row_index, order in enumerate(orders):
		row = [order.request_dict['serviceCode']]
		for shipping_method, price in order.shipping_rates.items():
			row.extend([shipping_method, price])
		for i in range(len(row)):
			sheet.write(row_index, i, row[i])
	workbook.save('shipping_adviser.xls')

if __name__ == "__main__":
	main()
