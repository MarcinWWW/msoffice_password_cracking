#!/usr/bin/env python3

import pandas as pd
import io
import msoffcrypto
import openpyxl
import aux_dict as ad
from itertools import product

def getListString(charset_start, charset_stop):
	ls = ""
	for item in range(charset_start, int(charset_stop)+1, 1):
		ls = ls + str(ad.getChar(item))
	return ls

def getListOfVariations(ls, c_num_start, c_num_stop):
	l = []
	for i in range(int(c_num_start), int(c_num_stop) + 1, 1):
		l.extend(["".join(item) for item in product(ls, repeat=i)])
	
	return l
	
#####################
# setup configuration
charset_start = 0
charset_stop = 35

pass_char_number_start = 1
pass_char_number_stop = 5

#FILE
location = '\\input\\abc.xlsx'

# optional, manual string
#list_string = "930"
######################

list_string = getListString(charset_start, charset_stop)
print(list_string)

list_of_variations = getListOfVariations(list_string, pass_char_number_start, pass_char_number_stop)
#print(list_of_variations)

ls_len = len(list_of_variations)
print("size = " + str(ls_len))

file_name = location[location.rfind("\\")+1:location.rfind(".")]
file_location = location[:location.rfind("input")]
file_out = open(file_location + "output\\" + file_name + ".txt", "w")

decrypted_workbook = io.BytesIO()

with open(location, 'rb') as file:
	office_file = msoffcrypto.OfficeFile(file)

	for pswd in list_of_variations:
	
		office_file.load_key(password=pswd)
		
		try:
			office_file.decrypt(decrypted_workbook)
			file_out.write("\nok_password = " + pswd)
			
			workbook = openpyxl.load_workbook(filename=decrypted_workbook)
			df = pd.read_excel(decrypted_workbook, engine="openpyxl")
			print("bingo = " + pswd)
			print(df)
			break
			
		except msoffcrypto.exceptions.InvalidKeyError:
			print(str(ls_len) + ' ' + pswd)
			file_out.write(pswd + "\n")
			ls_len -= 1
		
file.close()
file_out.close()
