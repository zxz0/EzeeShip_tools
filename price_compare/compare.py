"""
	Todo:
		save temp file (per block or per record: need unique identifier. considering row change... instead of start to request from beginning)
		(test) structurize
		modulization
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
import traceback
import re

CURRENT_VERSION = '2.0.0'

rules = {}
rules['Shipping'] = {}
rules['Packing'] = {}

rules['Shipping']["forward_address"] = []
rules['Shipping']["3kv_1_or_2"] = []
rules['Shipping']["3kv_3_or_more"] = []
rules['Shipping']["other_transformer"] = []
rules['Shipping']["parts"] = []
rules['Shipping']["normal"] = []
rules['Shipping']["residential_additional"] = []
rules['Shipping']["commercial_additional"] = []

rules['Packing']["3kv_1_or_2"] = ''
rules['Packing']["3kv_3_or_more"] = ''
rules['Packing']["other_trans_1"] = ''
rules['Packing']["other_trans_2"] = ''
rules['Packing']["other_trans_3_or_more"] = ''
rules['Packing']["others"] = ''

apply_desung_rules = True

positions = {}
positions['reference'] = 0
positions['sender_country'] = 0
positions['sender_address'] = 0
positions['sender_city'] = 0
positions['sender_state'] = 0
positions['sender_zipcode'] = 0
positions['recipient_country'] = 0
positions['recipient_address'] = 0
positions['recipient_city'] = 0
positions['recipient_state'] = 0
positions['recipient_zipcode'] = 0
positions['is_cod'] = 0
positions['cod_amount'] = 0
positions['length'] = 0
positions['width'] = 0
positions['height'] = 0
positions['weight'] = 0
positions['insurance_amount'] = 0

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
		self.is_residential_address = False

	# Mainly the shipping method and package
	def populate_other_properties(self):
		# Handle special address
		to_address = self.request_dict['to']
		if to_address['zipCode'] == '41025' and to_address['addressLine1'] == '1850 Airport Exchange Blvd #200' and to_address['city'] == 'Erlanger' and to_address['stateCode'] == 'KY' and to_address['countryCode'] == 'US':
			self.shipping_rates['fedex_ground'] = None
			self.shipping_rates['usps_priority'] = None
			return

		# Handle transformers
		lower_reference = self.reference.lower()
		if 'kv' in lower_reference:
			if '3kv' in lower_reference: 	# 3kv
				if 'x' not in lower_reference or 'x 2' in lower_reference or 'x2' in lower_reference:	# <= 2 (num), < 1lb
					for shipping_method in rules['Shipping']['3kv_1_or_2']:
						self.shipping_rates[shipping_method] = None 	# usps_first
					self.request_dict['parcels'][0]['packageCode'] = rules['Packing']['3kv_1_or_2'] 	# thick_envelope
				else:
					for shipping_method in rules['Shipping']['3kv_3_or_more']:
						self.shipping_rates[shipping_method] = None 	# usps_priority
					self.request_dict['parcels'][0]['packageCode'] = rules['Packing']['3kv_3_or_more'] 	# flat_rate_envelope
			else: 	# not 3kv
				for shipping_method in rules['Shipping']['other_transformer']:
						self.shipping_rates[shipping_method] = None 	# usps_priority
				if 'x' in lower_reference or '+' in lower_reference: 	# 2+ transformers, not 3kv
					if 'x 2' in lower_reference or 'x2' in lower_reference or lower_reference.count('+') == 1:	# 2 (num)
						self.request_dict['parcels'][0]['packageCode'] = rules['Packing']['other_trans_2'] 	# medium_flat_rate_box
					else:
						self.request_dict['parcels'][0]['packageCode'] = rules['Packing']['other_trans_3_or_more'] 	# large_flat_rate_box
				else: 	# just 1 transformer, not 3kv
					self.request_dict['parcels'][0]['packageCode'] = rules['Packing']['other_trans_1'] 	# flat_rate_envelope
			return

		# Handle signs
		if float(self.request_dict['parcels'][0]['weight'].strip()) < 1:	# parts
			for shipping_method in rules['Shipping']['parts']:
				self.shipping_rates[shipping_method] = None 	# usps_first
			return

		for shipping_method in rules['Shipping']['normal']:
			self.shipping_rates[shipping_method] = None 	# fedex_smart_post,usps_priority,ups_ground

		if self.is_residential_address: # residential
			for shipping_method in rules['Shipping']['residential_additional']:
				self.shipping_rates[shipping_method] = None 	# fedex_home_delivery
			self.request_dict['to']['isResidential'] = True
			self.request_dict['to']['isValid'] = True
		else: 	# commercial
			for shipping_method in rules['Shipping']['commercial_additional']:
				self.shipping_rates[shipping_method] = None 	# fedex_ground

	def set_best_rate(self):
		# just the lowest price
		best_rate_tuple = min(self.shipping_rates.items(), key=lambda kv: kv[1] if isinstance(kv[1], float) else sys.float_info.max) 
		self.best_shipping_service = best_rate_tuple[0]
		# self.request_dict['carrierCode'] = get_carrier_code_from_service_code(self.request_dict['serviceCode'])

def get_carrier_code_from_service_code(service_code):
	return service_code.split('_')[0]

def get_clear_cell_number_str(cell):
	return cell.value.strip() if cell.ctype == xlrd.book.XL_CELL_TEXT else str(int(cell.value))

def request_data(url, api_key, payload, result_key):
	headers = {'Authorization': api_key, 'Content-Type': 'application/json'}
	logging.debug("Send request: {request}".format(request = payload))
	response = requests.post(url, headers = headers, data = payload)
	json_response = response.json()
	logging.debug("Received response: {response}".format(response = json_response))

	if 'result' not in json_response:
		raise RequestError(payload, json_response, 'cannot get proper response! seems like internet problem')
	elif json_response['result'] == 'OK':
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

		logging.info('ignore head until row {row_num}'.format(row_num = self.head_to_ignore))
		for row_idx in range(self.head_to_ignore, sheetbooksheet.nrows):	# Ignore the head
			current_row = sheetbooksheet.row(row_idx)
			current_order = Order()

			# get info from the columns with correct format
			try:
				logging.info('parsing row {row_num}'.format(row_num = row_idx + 1))
				reference = current_row[positions['reference']].value.strip()

				sender_country = current_row[positions['sender_country']].value.strip()
				sender_address = current_row[positions['sender_address']].value.strip()
				sender_city = current_row[positions['sender_city']].value.strip()
				sender_state = current_row[positions['sender_state']].value.strip()
				sender_zipcode = get_clear_cell_number_str(current_row[positions['sender_zipcode']])

				recipient_country = current_row[positions['recipient_country']].value.strip()
				recipient_address = current_row[positions['recipient_address']].value.strip()
				recipient_city = current_row[positions['recipient_city']].value.strip()
				recipient_state = current_row[positions['recipient_state']].value.strip()
				recipient_zipcode = get_clear_cell_number_str(current_row[positions['recipient_zipcode']])

				is_cod = current_row[positions['is_cod']].value
				cod_amount = current_row[positions['cod_amount']].value if current_row[positions['cod_amount']].ctype == xlrd.book.XL_CELL_NUMBER else 0

				length = str(current_row[positions['length']].value)
				width = str(current_row[positions['width']].value)
				height = str(current_row[positions['height']].value)
				weight = str(current_row[positions['weight']].value)

				insurance_amount = current_row[positions['insurance_amount']].value if current_row[positions['insurance_amount']].ctype == xlrd.book.XL_CELL_NUMBER else 0
                
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
				parcel_info['packageCode'] = rules['Packing']["others"] # 'your_package'
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
				orders.append(current_order)
				logging.debug('get request dictionary: {request_dict}'.format(request_dict = current_order.request_dict))
			except ValueError as ve:
				traceback_info = traceback.format_exc()
				variable_name = re.search('(\w+) =', traceback_info)[1]
				column_index = re.search('\[(\d{1,2})]', traceback_info)[1]
				logging.error('wrong value for row {row_num}({row_info}), column {column_num}({column_meaning}): {cell_value}'.format(column_num = int(column_index) + 1, column_meaning = variable_name.replace('_', ' '), row_num = row_idx + 1, row_info = reference, cell_value = current_row[int(column_index)].value))
				exit(2)

		return orders

# sorted_rates, rule[0], rule[1], rule[2]);
def apply_rule(rates, delivery_method, compared_with, price_diff):
	compared_with_all_flag = (compared_with.lower() == 'all')
	for i in range(len(rates)):
		current_method = rates[i][0]
		current_rate = rates[i][1]
		current_pos = i

		if delivery_method == current_method:
			# compare all tags later (may have leap diff, so cannot only compare the next)
			for j in range(i + 1, lent(rates)):
				if compared_with_all_flag or compared_with in rates[j][0]: # if this is the one we want to compare
					if (rates[j][1] - current_rate < price_diff): 	# if price meet criteria
						rates[i], rates[j] = rates[j], rates[i] 	# switch
					break # if switched, end (only switch the first one, or problem happends for other equality?); if price not compatible: end

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
	parser.set_defaults(config_file = 'config.ini', input_file = 'input.xls', dest_file = 'shipping_adviser.xls')
	parser.add_option('-c', '--config', type = 'string', dest = 'config_file', help = 'use CONFIG_FILE to specify the API key', metavar = 'CONFIG_FILE')
	parser.add_option('-s', '--source', type = 'string', dest = 'input_file', help = 'read INPUT_XLS_FILE to load data', metavar = 'INPUT_XLS_FILE')
	parser.add_option('-d', '--dest', type = 'string', dest = 'dest_file', help = 'save rates information to OUTPUT_XLS_FILE', metavar = 'OUTPUT_XLS_FILE')

	(options, args) = parser.parse_args()
	logging.debug('using config file: {config_file} to processe {data_file}'.format(config_file = options.config_file, data_file = options.input_file, dest_file = options.dest_file))

	# Get API key, shipping rules, packing rules, sorting method, spreadsheet column position from config file
	api_key = ''
	if os.path.isfile(options.config_file):
		try:
			logging.info('Reading config file...')
			config = configparser.ConfigParser()
			config.read(options.config_file)

			logging.info('Geting API key...')
			api_key = config.get('Keys', 'api_key')
			logging.debug('Get API key: {api_key}'.format(api_key = api_key))

			logging.info('Geting rules...')
			global rules
			for key, value in rules['Shipping'].items():
				rules['Shipping'][key] = [delivery_method.strip() for delivery_method in config.get('Shipping', key).split(',')]
			for key, value in rules['Packing'].items():
				rules['Packing'][key] = config.get('Packing', key)
				if ',' in rules['Packing'][key]:
					logging.error('multiple packing options not allowed')
					exit(2)
			logging.debug(rules)

			logging.info('Geting sorting method...')
			global apply_desung_rules
			apply_desung_rules = config.getboolean('Sorting', 'apply_desung_rules')
			logging.debug('Get sorting method: apply desung rules? {apply_desung_rules}'.format(apply_desung_rules = apply_desung_rules))

			logging.info('Geting positions...')
			global positions
			for key, value in positions.items():
				positions[key] = config.getint('Position', key) - 1
			logging.debug(positions)
		except configparser.NoOptionError as noe:
			logging.error(noe)
			exit(2)
		
	else:
		logging.error('config file: {config_file} not exist!'.format(config_file = options.config_file))
		exit(2)

	# Get order info from xls file
	orders = []
	if os.path.isfile(options.input_file):
		logging.info('Parsing xls file...')
		xls_reader = XlsReader(options.input_file)
		orders = xls_reader.parse()
	else:
		logging.error('order file: {xls_file} not exist!'.format(xls_file = options.input_file))
		exit(2)

	# Validate and get rates for orders
	logging.info('Estimating rates...')
	for index, order in enumerate(orders):
		logging.info('Estimating rates for row: {row_index}...'.format(row_index = index + 1))

		# Validate address
		try:
			residential_flag = is_residential(api_key, json.dumps(order.request_dict['to']))
			order.is_residential_address = residential_flag
		except RequestError as re:
			logging.error('failed to pass validation in row {row_num}: {reason}'.format(row_num = index + 1, reason = re.message))

		order.populate_other_properties()
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
		rules = [['fedex_smart_post', 'all', 0.5], ['usps', 'fedex', 0.5]]
		# others - fedex_smart_post < 0.5: smart post
		# USPS - fedex_smart_post < 0.5: USPS
		for rule in rules:
			apply_rule(sorted_rates, rule[0], rule[1], rule[2]);
		if sorted_rates[0][0] == 'fedex_smart_post':
			if sorted_rates[0][1] and (sorted_rates[1][1] - sorted_rates[0][1] < 0.5):
				sorted_rates[0], sorted_rates[1] = sorted_rates[1], sorted_rates[0]
		row = [sorted_rates[0][0]]	# best shipping service (with lowest price)
		for shipping_method, price in sorted_rates:
			row.extend([shipping_method, price])
		for i in range(len(row)):
			sheet.write(row_index, i, row[i])
	workbook.save(options.dest_file)

if __name__ == "__main__":
	main()
