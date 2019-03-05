"""
	Todo:
		save temp file
		(test) structurize
		modulization
		packing (to exe)
		log info and debug level files
	Author: Zixuan Zhang
	Function: 
		parse xls to get order info,
		get best price among several shipping methods/services, 
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
from optparse import OptionParser

CURRENT_VERSION = '0.8.0'

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
		self.best_shipping_service = ''

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

		# Handle signs
		self.shipping_rates['fedex_smart_post'] = None
		self.shipping_rates['usps_priority'] = None
		self.shipping_rates['ups_ground'] = None

	def set_residential_commercial_method(self, is_residential_address):
		if is_residential_address:
			self.shipping_rates['fedex_home_delivery'] = None 	# residential
			self.request_dict['to']['isResidential'] = True
			self.request_dict['to']['isValid'] = True
		else: 	# commercial
			self.shipping_rates['fedex_ground'] = None

	def set_best_rate(self):
		# just the lowest price
		best_rate_tuple = min(self.shipping_rates.items(), key=lambda kv: kv[1] if isinstance(kv[1], float) else sys.float_info.max) 
		self.best_shipping_service = best_rate_tuple[0]
		# self.request_dict['carrierCode'] = get_carrier_code_from_service_code(self.request_dict['serviceCode'])

def get_carrier_code_from_service_code(service_code):
	return service_code.split('_')[0]

def request_data(url, api_key, payload, result_key):
	headers = {'Authorization': api_key, 'Content-Type': 'application/json'}
	logging.debug("Send request: {request}".format(request = payload))
	response = requests.post(url, headers = headers, data = payload)
	json_response = response.json()
	logging.debug("Received response: {response}".format(response = json_response))
	if json_response['result'] == 'OK':
		return json_response['data'][result_key]
	elif json_response['result'] == 'ERR':
		raise RequestError(payload, json_response, json_response['message'])
	# other errors should have been handled by requests

def is_residential(api_key, payload):
	url = 'https://ezeeship.com/api/ezeeship-openapi/address/validate'
	result_key = 'isResidential'
	try:
		residential_flag = request_data(url, api_key, payload, result_key)
	except RequestError:
		raise

	return residential_flag
	

def get_estimated_rate(api_key, payload):
	url = 'https://ezeeship.com/api/ezeeship-openapi/shipment/estimateRate'
	result_key = 'rate'
	try:
		rate = request_data(url, api_key, payload, result_key)
	except RequestError:
		raise

	return rate

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

	# Get (possible) arguments
	usage = 'usage: %prog [-options <argument>]'
	parser = OptionParser(usage = usage, version = '%prog {}'.format(CURRENT_VERSION))
	parser.set_defaults(config_file = 'config.ini', xls_file = 'input.xls')
	parser.add_option('-c', '--config', type = 'string', dest = 'config_file', help = 'use CONFIG_FILE to specify the API key merge', metavar = 'CONFIG_FILE')
	parser.add_option('-f', '--xls-file', type = 'string', dest = 'xls_file', help = 'use XLS_FILE as the uploaded xls file', metavar = 'XLS_FILE')

	(options, args) = parser.parse_args()
	logging.debug('using config file: {config_file} to processe {data_file}'.format(config_file = options.config_file, data_file = options.xls_file))

	# Get API key from config file
	api_key = ''
	if os.path.isfile(options.config_file):
		logging.info('Geting API key...')
		config = configparser.ConfigParser()
		config.read(options.config_file)
		api_key = config.get('Keys', 'api_key')
		logging.debug('Get API key: {api_key}'.format(api_key = api_key))
	else:
		logging.error('config file: {config_file} not exist!'.format(config_file = config_file))
		exit(2)

	# Get order info from xls file
	orders = []
	if os.path.isfile(options.xls_file):
		logging.info('Parsing xls file...')
		xls_reader = XlsReader(options.xls_file)
		orders = xls_reader.parse()
	else:
		logging.error('order file: {xls_file} not exist!'.format(xls_file = xls_file))
		exit(2)

	# Validate and get rates for orders
	logging.info('Estimating rates...')
	for index, order in enumerate(orders):
		logging.info('Estimating rates for row: {row_index}...'.format(row_index = index + 1))
		order.populate_other_properties()

		# Validate address
		try:
			residential_flag = is_residential(api_key, json.dumps(order.request_dict['to']))
			order.set_residential_commercial_method(residential_flag)
		except RequestError as re:
			logging.error('failed to pass validation in row {row_num}: {reason}'.format(row_num = index + 1, reason = re.message))

		# Get rates
		for shipping_method in order.shipping_rates.keys():
			order.request_dict['serviceCode'] = shipping_method
			order.request_dict['carrierCode'] = get_carrier_code_from_service_code(shipping_method)
			try:
				price = get_estimated_rate(api_key, json.dumps(order.request_dict))
				order.shipping_rates[shipping_method] = price
			except RequestError as re:
				logging.error('cannot get rate for row {row_num}: {reason}'.format(row_num = index + 1, reason = re.message))
				order.shipping_rates[shipping_method] = re.message
		# order.set_best_rate()

	# Write results to xls file
	logging.info('Writing results...')
	workbook = xlwt.Workbook()
	sheet = workbook.add_sheet('Sheet 1')
	for row_index, order in enumerate(orders):
		# Sort the shipping price dictionary, put best in the front, put error message at the end
		sorted_rates = sorted(order.shipping_rates.items(), key=lambda kv: kv[1] if isinstance(kv[1], float) else sys.float_info.max)
		row = [sorted_rates[0][0]]	# best shipping service (with lowest price)
		for shipping_method, price in sorted_rates:
			row.extend([shipping_method, price])
		for i in range(len(row)):
			sheet.write(row_index, i, row[i])
	workbook.save('shipping_adviser.xls')

if __name__ == "__main__":
	main()
