# coding: utf-8
import xlwt
import xlrd
from xlutils.copy import copy

#open and copy Excel sheet
rb = xlrd.open_workbook('BeneFix Small Group Plans upload template.xlsx')
wb = copy(rb)
w_sheet = wb.get_sheet(0)

#open txt file and write to Excel sheet
def fileHandling(x, y, fileName):
	with open(fileName, 'r') as file:
		line = file.readlines()
		counter = 0
		for index in line:
			line[counter].strip()
			counter =+ counter + 1
		#parsing start and end dates
		line0 = line[0 + x].split(':') 
		line0ab = line0[1]
		line0bc = line0ab.split('-')
		line0ab = line0bc[0]
		line0bc = line0bc[1]
		line0ab = line0ab.strip()
		line0bc = line0bc.strip()
		w_sheet.write(y, 0, line0ab)
		w_sheet.write(y, 1, line0bc)

		if 'PENNSYLVANIA' in line[2 + x]:
			w_sheet.write(y, 3, 'PA')
		if 'PARA01' in line[6 + x]:
			w_sheet.write(y, 4, '01')
		elif 'PARA02' in line[6 + x]:
			w_sheet.write(y, 4, '02')
		elif 'PARA03' in line[6 + x]:
			w_sheet.write(y, 4, '03')
		elif 'PARA04' in line[6 + x]:
			w_sheet.write(y, 4, '04')
		elif 'PARA05' in line[6 + x]:
			w_sheet.write(y, 4, '05')
		elif 'PARA06' in line[6 + x]:
			w_sheet.write(y, 4, '06')
		elif 'PARA07' in line[6 + x]:
			w_sheet.write(y, 4, '07')
		elif 'PARA08' in line[6 + x]:
			w_sheet.write(y, 4, '08')
		elif 'PARA09' in line[6 + x]:
			w_sheet.write(y, 4, '09')

		#handling product name
		line10 = line[10 + x].split(':')
		line10 = line10[1].strip()
		w_sheet.write(y, 2, line10)

	#automation of writing to Excel rates for each age range
		#18-20
		line[30 + x] = line[30 + x].strip()
		w_sheet.write(y, 5, line[30 + x])
		w_sheet.write(y, 6, line[30 + x])

		#21 - 34
		counter1 = 7
		for index in range(31 + x, 47 + x):
			line[index] = line[index].strip()
			w_sheet.write(y, counter1, line[index])
			counter1 += 1
		#35 - 49
		counter2 = 21
		for index2 in range(64 + x, 79 + x):
			line[index2] = line[index2].strip()
			w_sheet.write(y, counter2, line[index2])
			counter2 += 1
		#50 - 64
		counter3 = 36
		for index3 in range(98 + x, 113 + x):
			line[index3] = line[index3].strip()
			w_sheet.write(y, counter3, line[index3])
			counter3 += 1
		#65+
		w_sheet.write(y, 51, line[112 + x])

#do fileHandling function 45 times, parses and writes the whole txt file, possibly could use for loop here to 
#cut down on cutting and pasting function
def overallFileHandling(x, y, fileName):
	fileHandling(0, 1 + y, fileName)
	fileHandling(x, 2 + y, fileName)
	fileHandling(2 * x, 3 + y, fileName)
	fileHandling(3 * x, 4 + y, fileName)
	fileHandling(4 * x, 5 + y, fileName)
	fileHandling(5 * x, 6 + y, fileName)
	fileHandling(6 * x, 7 + y, fileName)
	fileHandling(7 * x, 8 + y, fileName)
	fileHandling(8 * x, 9 + y, fileName)
	fileHandling(9 * x, 10 + y, fileName)
	fileHandling(10 * x, 11 + y, fileName)
	fileHandling(11 * x, 12 + y, fileName)
	fileHandling(12 * x, 13 + y, fileName)
	fileHandling(13 * x, 14 + y, fileName)
	fileHandling(14 * x, 15 + y, fileName)
	fileHandling(15 * x, 16 + y, fileName)
	fileHandling(16 * x, 17 + y, fileName)
	fileHandling(17 * x, 18 + y, fileName)
	fileHandling(18 * x, 19 + y, fileName)
	fileHandling(19 * x, 20 + y, fileName)
	fileHandling(20 * x, 21 + y, fileName)
	fileHandling(21 * x, 22 + y, fileName)
	fileHandling(22 * x, 23 + y, fileName)
	fileHandling(23 * x, 24 + y, fileName)
	fileHandling(24 * x, 25 + y, fileName)
	fileHandling(25 * x, 26 + y, fileName)
	fileHandling(26 * x, 27 + y, fileName)
	fileHandling(27 * x, 28 + y, fileName)
	fileHandling(28 * x, 29 + y, fileName)
	fileHandling(29 * x, 30 + y, fileName)
	fileHandling(30 * x, 31 + y, fileName)
	fileHandling(31 * x, 32 + y, fileName)
	fileHandling(32 * x, 33 + y, fileName)
	fileHandling(33 * x, 34 + y, fileName)
	fileHandling(34 * x, 35 + y, fileName)
	fileHandling(35 * x, 36 + y, fileName)
	fileHandling(36 * x, 37 + y, fileName)
	fileHandling(37 * x, 38 + y, fileName)
	fileHandling(38 * x, 39 + y, fileName)
	fileHandling(39 * x, 40 + y, fileName)
	fileHandling(40 * x, 41 + y, fileName)
	fileHandling(41 * x, 42 + y, fileName)
	fileHandling(42 * x, 43 + y, fileName)
	fileHandling(43 * x, 44 + y, fileName)
	fileHandling(44 * x, 45 + y, fileName)
	
#call overallFileHandling function for each text file, writes all of them to Excel file
overallFileHandling(120, 0, 'para01.txt')
overallFileHandling(120, 45, 'para02.txt')
overallFileHandling(120, 45 * 2, 'para03.txt')
overallFileHandling(120, 45 * 3, 'para05.txt')
overallFileHandling(120, 45 * 4, 'para06.txt')
overallFileHandling(120, 45 * 5, 'para07.txt')
overallFileHandling(120, 45 * 6, 'para08.txt')
overallFileHandling(120, 45 * 7, 'para09.txt')

wb.save('BeneFix Small Group Plans upload template.xlsx')

 
