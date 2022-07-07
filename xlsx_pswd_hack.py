#!/usr/bin/env python3
#if a job was stopped ctrl + c you can continue with a set of checked passwords
#by running the program with the file of set of passwords as parameter
#e.g. python xlsx_pswd_hack.py parameter.txt

import sys
#import xlrd
#import win32com.client

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
	
def readInProgress(file):
	ls_ex = []
	with open(file, 'r') as my_file:
		for line in my_file.readlines():
			ls_ex.append(line.rstrip())
	return ls_ex
	
###################	
#SETUP
charset_start = 26
charset_stop = 35

#custom list_string
#list_string = "930"
list_string = getListString(charset_start, charset_stop)
print(list_string)

pass_char_number_start = 1
pass_char_number_stop = 4

#FILE
location = 'C:\\Mar\\dokumenty\\Python\\xlsx_hack\\input\\pin_9930.xlsx'

####################


list_of_variations = getListOfVariations(list_string, pass_char_number_start, pass_char_number_stop)
#print(list_of_variations)

list_of_exclusions = []
list_excluded = []

if len(sys.argv) > 1:
	list_of_exclusions = readInProgress(sys.argv[1])
	list_excluded = [x for x in list_of_variations if x not in list_of_exclusions]
else:
	list_excluded = list_of_variations

ls_excluded_len = len(list_excluded)
print("TODO variations = " + str(ls_excluded_len))
#print(list_of_exclusions)


file_name = location[location.rfind("\\")+1:location.rfind(".")]
file_location = location[:location.rfind("input")]
file_out = open(file_location + "output\\" + file_name + ".txt", "w")
file_out.writelines([str(l) + '\n' for l in list_of_exclusions])

decrypted_workbook = io.BytesIO()

cstart = charset_start
cdiff = charset_stop - charset_start + 1

with open(location, 'rb') as file:
	office_file = msoffcrypto.OfficeFile(file)

	for pswd in list_excluded:
	
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
			print("pswd = " + pswd)
			file_out.write(pswd + "\n")
		
file.close()
file_out.close()