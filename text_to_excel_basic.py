import os
import glob
import shutil
from pathlib import Path
import time
import xlsxwriter
from string import ascii_uppercase as ALPHABETS
import zipfile as zp
import logging
import re
from decimal import Decimal
'''
Note : Latest python has best practice syntax for defining functions as:
		def func(arg:type)->return_type:
			---
		Although it can be also written as :
		def func(arg):
			---


====
MARGIN : row number from which actual data with column headers will be written
====
'''
MARGIN = 3

# basic worksheet formatting style 
BASIC = {'bold': 1,
	'border': 1,
	'align': 'center',
	'valign': 'vcenter'
	}



def rem_empty_str(li: list) -> filter:
	'''
		removes empty strings from list obtained from file object.
		Return: A filter object
	'''
	return filter(None, li.split('  '))

# utility function for dividing numbers with 100
def div_ten_float(arg):
	if arg!='':
		return float('%.2f'%float(arg))/100
	else:
		return arg	


def clean_line(st: str) -> list:
	return [i.strip() for i in rem_empty_str(st)]


def get_dic(st:str)->dict:

	'''
		Note: These slices are taken according to provided raw txt file.
	'''

	# slicing strings according to the headings.

	dic = {
		'employee_name':st[:32],
		'earn_desc':st[32:45],
		'pay_rate':st[45:53],
		'cur_hrs':st[54:63],
		'cur_amt':st[64:73],
		'yt_hrs':st[74:83],
		'yt_amt':st[84:93],
		'deduc_desc':st[93:109],
		'curnt_amt1':st[109:120],
		'ytd_amt1':st[121:131],
		'taxes_desc':st[131:143],
		'curnt_amt2':st[145:154],
		'ytd_amt2':st[156:165],
		'net_pay':st[167:177],
	}

	# cleaning the values
	for key,val in dic.items():
		v = val.strip()
		
		dic[key]=v
		# making ints if possible
		try:
			if key in ['cur_hrs','cur_amt','yt_hrs','yt_amt']:

				# import pdb
				# pdb.set_trace()
				# print(v)
				dic[key]=float('%.2f'%float(v))/100
			else:
				dic[key]=int(v)
		except ValueError as e:
			# print(e)
			pass

	return dic


def create_worksheet(name)->tuple:

	'''
		Description: Create a xlsxwriter worksheet and workbook
		Return: Tuple of workbook and worksheets
	'''

	# creating a excel workbook(document)
	workbook = xlsxwriter.Workbook(name)
	row_format = BASIC.copy()
	# creating a sheet in excel workbook
	worksheets = []
	sheet_names = ['EMPLOYEE EARNINGS','DEDUCTIONS','TAXES','EMPLOYEE PAYMENT']
	for i in range(4):
		worksheet = workbook.add_worksheet(sheet_names[i])
		
		if i==0:
			worksheet.set_column('D:E',18)
			worksheet.set_column('F:G',16)


		worksheet.set_column('C:D',20)
		
		worksheet.set_column('A:A',20)
		worksheet.set_column('B:B',27)
		
		# header rows size big
		
		worksheet.set_row(1,20)
		worksheets.append(worksheet)




	return workbook,worksheets


def main_parsing(input_file,output_path):
	
	global MARGIN
	global OUTPUT_FILENAME
	# global logger

	output_excel_filename = OUTPUT_FILENAME
	margin = MARGIN

	if output_excel_filename.endswith('.xlsx'):
		output_excel_filename=output_excel_filename[:-5]
	
	output_excel_filename+=time.strftime("_%Y_%m_%d__%H_%M_%S")+'.xlsx'
	# print('Excel File:',output_excel_filename)
	
	output_file = os.path.join(output_path,output_excel_filename)
	workbook,worksheets = create_worksheet(name=output_file)
	worksheet1,worksheet2,worksheet3,worksheet4 = worksheets
	# getting file object in read mode 
	fo = open(input_file,'r')
	# list of all the lines in file
	fli = fo.readlines()
	# filtering list for getting list of all the lines without newline chr and spaces.
	fli = list(filter(lambda i: not i.isspace(),fli))
	# found first/second header to write header only once 
	# found_l is for last line (footer)
	found_fh = found_sh = found_l = False
	row = row2 = row3 = row4 = margin+2
	col = 0
	record = 0
	validated = False
	phone_li = [None]
	error_line = None
	a = []
	# var to tract last found pay rate
	lpr = None
	# var to track last found pay rate row
	lpr_r = -1
	# regex for getting employee id
	emp_re = re.compile('^([0-9]+) ([ \w-]*)$')

	# formatting rows
	row_format = BASIC.copy()
	row_format.update({'bg_color':'#4CB5F5','font_color':'white'})
	# formatiing floats
	float_format = workbook.add_format({'num_format': '0.00'})

	for idx,i in enumerate(fli):
		if i.startswith('  EMPLOYEE NAME'):
			if not found_fh:
				a = clean_line(i)
				# set_fheader(workbook,worksheet,l)
				found_fh = True
		elif i.startswith('  ID SSN STATE/FRQ STS LOCATION'):
			if not found_sh:
				l = clean_line(i)

				# header list to be written
				hli = []
				fi = si = 0
				while fi<len(l) and si<len(l):
					if l[si]=='DESCR':
						l[si]="DESCRIPTION"
					hli.append(a[fi]+' '+l[si])
					if 'ID SSN' in hli[-1]:
						hli[-1]='EMPLOYEE ID'
					if fi in [3,4]:
						si+=1
						hli.append(a[fi]+' '+l[si])
					fi+=1
					si+=1

				del fi,si

				worksheet1.write_row(1+margin,0,hli[:7],cell_format=workbook.add_format(row_format))
				worksheet2.write_row(1+margin,0,[hli[0]]+hli[7:10],cell_format=workbook.add_format(row_format))
				worksheet3.write_row(1+margin,0,[hli[0]]+hli[10:-1],cell_format=workbook.add_format(row_format))

				ws4_heading = ['EMPLOYEE ID','EMPLOYEE NAME','PAY TYPE',hli[-1]]
				worksheet4.write_row(1+margin,0,ws4_heading,cell_format=workbook.add_format(row_format))
				found_sh = True

				del hli
		elif i.startswith('   EMPLOYEE TOTAL'):
			l = clean_line(i)


			data = list(map(div_ten_float,l[1:]))
			strt_to_end = ALPHABETS[0]+str(row+1)+':'+ALPHABETS[2]+str(row+1)
			row_format = BASIC.copy()
			row_format.update({'bg_color': '#EA6A47'})
			worksheet1.write_row(row,0,[l[0]]+['',''],workbook.add_format(row_format))
			# worksheet1.set_row(row,workbook.add_format(row_format))
			worksheet2.write_row(row2,0,[l[0]]+[''],workbook.add_format(row_format))
			worksheet3.write_row(row3,0,[l[0]]+[''],workbook.add_format(row_format))
			# worksheet1.merge_range(strt_to_end, l[0], workbook.add_format(row_format))
			row_format.update({'num_format': '0.00'})
			worksheet1.write_row(row,3,data[:4],workbook.add_format(row_format))
			data = data[4:]
			d = 0
			for i in range(1,7):
				if i%3!=1:
					if i>=4:
						worksheet3.write(row3,i-3,data[d],workbook.add_format(row_format))
						
					else:
						worksheet2.write(row2,i,data[d],workbook.add_format(row_format))
						
					d+=1
			row+=1
			row3 += 1
			row2 += 1	
		elif i.startswith(' '*65):
			pass
		elif i.startswith('-'*176):
			record+=1
		elif i.startswith(' PAYROLL REGISTER') and 'CHECK DATE' in i:
			if idx==0:
				validated = True
				# comment below two lines for not writing header
				l = clean_line(i)
				for worksheet in worksheets:
					worksheet.write_row(0,0,l,workbook.add_format(BASIC))
		elif i.startswith(' CXXX SXXXXXXX XXXXX XXX - XXXX'):
			# comment below two lines for not writing header
			l = clean_line(i)
			for worksheet in worksheets:
				worksheet.write_row(1,0,l[:-2],workbook.add_format(BASIC))
		elif 'PHONE' in i:
			l = clean_line(i)
			found_l = True
			phone_li = l
		elif i.startswith(' XXXX'):
			pass
		else:
			if idx!=0 :
				dic = get_dic(i)
				if dic['pay_rate']!='':
					lpr_r = row
					lpr = dic['pay_rate']

				# if found hourly then dividing pay rate by 1000 and writing into excel
				if 'Hourly' in dic['employee_name']:
					# print(dic['employee_name'].split()[0])
					worksheet1.write_number(lpr_r,2,lpr/10000,float_format)
					worksheet1.write(row-3,2,Decimal(dic['employee_name'].split()[0]),float_format)


				data_li = list(dic.values())
				# print(dic.values())
				worksheet1.write_row(row,1,data_li[1:7],float_format)

				# worksheet2
				if not all(''==s or s.isspace() for s in [data_li[0]]+data_li[7:10]):
					data_li[8:10] = list(map(div_ten_float,data_li[8:10]))
					# print(data_li[8:10])
					worksheet2.write_row(row2,1,data_li[7:10],float_format)
					row2+=1

				# worksheet3
				if not all(''==s or s.isspace() for s in [data_li[0]]+data_li[10:-1]):
					data_li[11:-1] = list(map(div_ten_float,data_li[11:-1]))
					worksheet3.write_row(row3,1,data_li[10:-1],float_format)
					row3+=1

				# worksheet4
				matched = re.match(emp_re,dic['employee_name'])

				if matched: 
					emp_id, emp_name = matched.groups()
					worksheet4.write_row(row4,0,list([emp_id,emp_name]))
					# worksheet1.write(row-1,0,emp_id)

					def _write_emp_id(r,wsheet,emp_id):

						rev = r-1
						# print(rev)
						for i in range(rev,rev+5):
							wsheet.write(i,0,emp_id)
			
					_write_emp_id(row,worksheet1,emp_id)
					_write_emp_id(row2-1,worksheet2,emp_id)
					_write_emp_id(row3-1,worksheet3,emp_id)
									

				if isinstance(data_li[-1],int):
					worksheet4.write(row4,3,data_li[-1])

				elif isinstance(data_li[-1],str):	
					if not (data_li[-1].isspace() or data_li[-1]==''):
						if 'DIRDEP' in data_li[-1]:
							data_li[-1]='DIRECT DEPOSIT'
						worksheet4.write(row4,2,data_li[-1])
						row4+=1

				row+=1
				

				del dic

	# writing footer
	# comment below if statement block for not writing footer into excel
	# if found_l:
	# 	worksheet1.write(row+1,0,l[0],workbook.add_format(BASIC))
	# 	worksheet1.merge_range(row+1,3,row+1,6,' '.join(phone_li[1:]),workbook.add_format(row_format))

	fo.close()

	if validated:
		workbook.close()
		# make_zip(os.path.abspath(output_file))
		msg = 'Total records processed: '+str(record)+' Output Excel: '+output_file
		logging.info(msg)
		print('Total record processed: ',record)
		print('Please open ',output_file)
	else:
		print('Error Occured. Invalid File')
		# file_name = os.path.basename(input_file)
		# error_path = OUTPUT_ERROR_DIR
		# shutil.move(input_file,os.path.join(error_path,file_name))
		# message = f'[Error] {file_name} is not a valid text file. Moved to error folder.'
		# print(message)
		# logging.error(message)
	


if __name__=='__main__':

	
	INPUT_FILE = 'payroll.txt'
	OUTPUT_DIR = os.path.join(os.getcwd(),'output')
	OUTPUT_FILENAME = 'Payroll'

	# check if the immediate parent input path is present or not.
	parent_output_dir = os.path.abspath(os.path.join(OUTPUT_DIR, os.pardir))
	# if present input path does not exist then create one.
	if not os.path.exists(parent_output_dir):
		raise Exception(f'Parent output folder {parent_output_dir} doesn\'t exists.')
	else:
		if not os.path.exists(OUTPUT_DIR):
			os.mkdir(OUTPUT_DIR)

	main_parsing(INPUT_FILE,OUTPUT_DIR)