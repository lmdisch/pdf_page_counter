# -*- coding: utf-8 -*-
####################################################################
# Import ###########################################################
####################################################################
import string
import os
import re
import Tkinter, tkFileDialog, tkSimpleDialog
from Tkinter import *
import ttk
import PyPDF2
import xlsxwriter
import time 

####################################################################
# GUI ##############################################################
####################################################################
root = Tkinter.Tk()

# GUI attributes
root.title('Tab 2 Bates Range Automation Tool')
root.geometry('500x250+50+50')
my_root_text = label(root, text='Please select a directory of Bates-stamped documents. This window will automatically close when the program finishes running.')
my_root_text.pack()
my_root_progress = ttk.Progressbar(root, orient='horizontal', length=200, mode='determinate')
my_root_progress.pack()
my_root_text.config(wraplength=500, justify=CENTER)

# Dialogue with user input
input_folder = tkFileDialog.askdirectory(parent=root, initialdir='/', title='This tool allows you to quickly determine the entire range of Bates within a document')

####################################################################
# Tool #############################################################
####################################################################

# Define variable to count pages for countPages
rxcountpages = re.compiler(r'/Type\s*/Page([^s]|$)', re.MULTILINE|re.DOTALL)

# Function to count number of pages in PDF document
def countPages(filename, folder):
	try:
		data = file(os.path.join(folder, filename.decode('utf-8')),'rb')
		page_count = PyPDF2.PdfFileReader(data)
		return page_count.numPages
	except:
		data = file(os.path.join(folder, filename.decode('utf-8')),'rb').read()
		return len(rxcountpages.findall(data))

# Function that creates vars for page count, filename, and hyperlinks to file
def Files(filename, value):
	start = filename
	if lower(start).find('.pdf'):
		page_count = countPages(filename, value)
		bate = filename.replace('.pdf','').replace('.PDF','')
		x = 0
		y = 0
		chrs = list(bate)
		while y < len(bate):
			if not(chrs[y].isdigit()):
				x = y 
			y += 1
		bateTail = bate[(x+1):]
		bateFront = bate[:(x+1)]
		strEnd = str(int(bateTail) + page_count - 1)
		while len(strEnd) < len(bateTail):
			strEnd = '0' + strEnd
		end = str(bateFront) + str(strEnd)
	else:
		bate = filename[:start.find('.')]
		end = bate 
		page_count = 1


	hyperlink = '=HYPERLINK("' + str(os.path.join(value.encode('utf-8'), filename)) + '")'
	return hyperlink, start, bate, end, page_count

# Function that outputs two lists (Dir, errorList) of Bates data
def runFiles(folder):
	Dir = []
	errorList = []
	filesDict = {}
	counter = 0

	for root, directories, files in os.walk(folder):
		for filename in files:
			filesDict[filename] = root

	my_root_progress['maximum'] = len(filesDict)

	for name, value in filesDict.items():
		name - name.encode('utf-8')
		counter += 1
		my_root_progress['value'] = counter
		my_root_progress.update()

		try:
			hyperlink, start, bate, end, page_count = Files(name, value)
			if page_count == 0:
				A_Range = []
				A_Range.append(hyperlink)
				A_Range.append(name)
				A_Range.append(page_count)
				A_Range.append(bate)
				A_Range.append(end)
				errorList.append(A_Range)
			else:
				B_Range = []
				B_Range.append(hyperlink)
				B_Range.append(name)
				B_Range.append(page_count)
				B_Range.append(bate)
				B_Range.append(end)
				Dir.append(B_Range)

		except:
			errorList.append(name)

	return Dir, errorList

# Function that outputs lists to .xlsx file
def outputDirectory(i_folder):
	Dir, errorList = runFiles(i_folder)

	# Ask for case code to ensure unique file name
	output_code = tkSimpleDialog.askstring('Case Code','Please enter case code: ')
	output_file = time.strftime('%Y.%m.%d') + '-' + str(output_code) + 'Bates.xlsm'

	# Create MS excel workbook
	workbook = xlsxwriter.Workbook(output_file)
	workbook.add_vba_project(r'...') # redacted -- hardcoded .bin file for VBA macro

	worksheet = workbook.add_worksheet('Documents')
	row = 1
	col = 0

	# Style workbook
	titles = workbook.add_format({
		'font_size': 10,
		'align': 'center_across',
		'font_name': 'Times New Roman',
		'bold': 1
		})

	text = workbook.add_format({
		'font_size': 10,
		'font_name': 'Times New Roman',
		})

	# Add button tied to VBA macro
	worksheet.insert_button('L1', {'macro': 'collapseBates',
							'caption': 'Collapse Bates',
							'width': 80,
							'height': 30})

	# Output text and formatting
	worksheet.write('A1': 'Suggested Bates start', titles)
	worksheet.write('B1': 'Suggested Bates end', titles)
	worksheet.write('C1': 'Number of pages', titles)
	worksheet.write('D1': 'Filename', titles)
	worksheet.write('E1': 'Hyperlink to File', titles)
	worksheet.write('G1': 'Bates start', titles)
	worksheet.write('I1': 'Bates end', titles)

	# Set column widths
	worksheet.set_column('A:D',20,text)
	worksheet.set_column('E:E',70,text)
	worksheet.set_column('G:G',20,text)
	worksheet.set_column('H:H',2.29,text)
	worksheet.set_column('I:I',20,text)

	# Output function lists to sheet
	for hyperlink, start, page_count, bate, end in Dir:
		worksheet.write(row, col, bate.decode('utf-8'),text)
		worksheet.write(row, col + 1, end.decode('utf-8'),text)
		worksheet.write(row, col + 2, page_count,text)
		worksheet.write(row, col + 3, start.decode('utf-8'),text)
		worksheet.write(row, col + 4, hyperlink.decode('utf-8'),text)
		row += 1

	for name in errorList:
		worksheet.write(row, col,'', text)
		worksheet.write(row, col + 1,'', text)
		worksheet.write(row, col + 2,'', text)
		worksheet.write(row, col + 3,str(name), text)
		worksheet.write(row, col + 4,'Error', text)
		row += 1

	workbook.close()

outputDirectory(input_folder)
root.withdraw()




