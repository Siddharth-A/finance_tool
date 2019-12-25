#!/usr/bin/env python3

import sys
import os
import csv
import json
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

import categories

csv_file = 'cibc.csv'
orange = PatternFill(start_color='fc6e21',end_color='fc6e21',fill_type='solid')
purple = PatternFill(start_color='e278ff',end_color='e278ff',fill_type='solid')
red    = PatternFill(start_color='de6464',end_color='de6464',fill_type='solid')
yellow = PatternFill(start_color='fffa99',end_color='fffa99',fill_type='solid')
green  = PatternFill(start_color='a3f799',end_color='a3f799',fill_type='solid')
blue   = PatternFill(start_color='00d8f5',end_color='00d8f5',fill_type='solid')
white  = PatternFill(start_color='ffffff',end_color='ffffff',fill_type='solid')


# converts the downloaded csv file to xlsx format
def csv_to_xlsx(input_file):
	print ("\n1) Converting dowloaded CSV file to XLSX\n")
	output_file = input_file.replace('csv','xlsx')			#create output name string
	wb = Workbook()
	ws = wb.active
	with open(input_file, 'r') as f:						#copy csv contents to xlsx 
	    for row in csv.reader(f):
	        ws.append(row)
	wb.save(output_file)
	return output_file


def cell_color(input_file):
	wb1 = load_workbook(input_file)
	ws_of_wb1 = wb1["cibc"]
	colB = ws_of_wb1['B']
	for cell_des in colB:
		cell = cell_des.value
		if any(word in cell for word in categories.resteraunts):
			cell_des.fill = orange
			print ("resteraunt: ", cell)
			continue 

		elif any(word in cell for word in categories.transport):
			cell_des.fill = purple
			print ("transport: ", cell)
			continue

		else:
			continue

	wb1.save(input_file)


#xlsx_file = csv_to_xlsx(csv_file)
xlsx_file = 'cibc.xlsx'
cell_color(xlsx_file)
