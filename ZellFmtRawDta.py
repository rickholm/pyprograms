# Reformat ComCast Raw CSV Data to Zelle CSV Data Format
# 08-31-2021 - RAH - New
# 09-04-2021 - RAH - Changed to work in pycharm
# 09-05-2021 - RAH - Added support spreadsheet xlsx data in

import os
import csv
import sys, getopt
import re
import openpyxl

progname = sys.argv[0]
argv = sys.argv[1:]

etvid = 'EffecTV_APP_EffecTV_A35-54'  # Placement EffecTv Id

# Creative Tables

crtvfr = ['GameNight', 'Homework', 'NewDog']  # Scan For
crtvto = ['GameNight_Video_', 'HomeWork_Video_', 'NewDog_Video_']  # To Value

# Target Audience Tables

tgtafr = ['ChildFreeAdults', 'NestingFamilies', 'NewlyLiberated']  # Scan for
tgtato = ['Child Free Adults', 'Nesting Families', 'Newly Liberated']  # To Name
tgtaabv = ['CF', 'NF', 'NL']  # Abbreviations

# Placement Tables

plcmfr = ['IP_15', 'IP_30', 'STB_15', 'STB_30']
plcmto = ['IPBased_Video_15', 'IPBased_Video_30', 
          'SetTop_Video_15', 'SetTop_Video_30']
plcmtm = ['15', '30', '15', '30']

# Billing information

makegood = ['_MG']
billrate = 38

# get input and output file names from command line_count

inputfile = ''
outputfile = ''
rplcfilesw = False

print(' ')
try:
	opts, args = getopt.getopt(argv, 't:h:i:o:r', ['ifile =', 'ofile ='])
except getopt.GetoptError:
	print(progname, ' -t <csv> or <xlsx> -i <inputfile> -o <outputfile> -r')
	sys.exit(2)
for opt, arg in opts:
	if opt == '-h':
		print(progname, '-t <csv or xlsx> -i <inputfile> -o <outputfile> -r')
		sys.exit()
	elif opt in ("-t"):
		filetype = arg
		wkfiletype = arg.upper()
		if (wkfiletype != 'CSV') and (wkfiletype !='XLSX'):
			print('Invalid file type, must be csv or xlsx, found: ' + filetype)
			sys.exit(1)
	elif opt in ("-i", "--ifile"):
		inputfile = arg
	elif opt in ("-o", "--ofile"):
		outputfile = arg
	elif opt == '-r':
		rplcfilesw = True
print(filetype, wkfiletype)
sys.exit()
print('Getting raw data from', inputfile, 'and reformatting data to', outputfile)

# Make sure raw data file exists and output file is new

if not os.path.exists(inputfile):
	print('Raw data file ' + inputfile + ' not found...canceled')
	sys.exit(1)
if os.path.exists(outputfile) and rplcfilesw == False:
	print('Output file ' + outputfile + ' found...canceled')
	sys.exit(1)
if os.path.exists(outputfile) and rplcfilesw == True:
	os.remove(outputfile)

# loop to get raw data and reformat

with open(inputfile) as rawdata_file:
	with open(outputfile, mode='w', newline='') as xlt_file:

		raw_reader = csv.reader(rawdata_file, delimiter=',')

		xlt_writer = csv.writer(xlt_file, delimiter=',')

		firstsw    = True
		line_count = 0
		outlinecnt = 0
		savevtdate = ''
		evtdatecnt = 0
		
		totprojrev = 0
		totactrev  = 0

		mgcnt      = 0
		mgimp      = 0
		
		errcnt     = 0
		errsw      = False
	
		for row in raw_reader:
			line_count += 1
			if firstsw == True:
				firstsw = False
				xlt_writer.writerow(['Date (Daily)', 'Targeting/Audience', 'Placement', 'Creative', 'Impressions',
				'MakeGood', 'Rate Per 1000', 'Proj Cost', 'Actual Cost'])
				outlinecnt += 1
			else:
				
				if savevtdate != row[0]:
					savevtdate = row[0]
					evtdatecnt += 1
					
				# Look for target audience
				
				i = 0	
				lusw = False
				for fr in tgtafr:
					if re.search(fr, row[2]) != None:
						tgta     = tgtato[i]			
						plcmabv  = tgtaabv[i]	
						lusw = True
						break
					i += 1
				if lusw == False:
					print('Data error in Target Audience', row[2])
					errcnt += 1
					errsw = True
					continue
					
				# Look for placement
				
				i = 0
				lusw = False
				for fr in plcmfr:
					if re.search(fr, row[1]) != None:
						plcmmeth = plcmto[i]
						plcmtime = plcmtm[i]
						lusw = True
						break
					i += 1
				if lusw == False:
					print('Data error in Placement Name', row[1])
					errcnt += 1
					errsw = True
					continue

				plcm = etvid + plcmabv + '_' + plcmmeth + '_SS'
				
				# Look for creative
				
				i = 0
				lusw = False
				for fr in crtvfr:
					if re.search(fr, row[1]) != None:
						crtv = crtvto[i] + plcmtime
						lusw = True
						break
					i += 1
				if lusw == False:
					print('Data error with Creative', row[1])
					errcnt += 1
					errsw = True
					continue

				# Compute billing information, check if this is a make good transaction
								
				i = 0
				mgtrans    = '   '
				makegoodsw = False
				for fr in makegood:
					if re.search(fr, row[2]) != None:
						mgtrans    = 'Yes'
						mgcnt     += 1
						makegoodsw = True
						break
					i += 1

				impressions = int(re.sub(',', '', row[3]))  # Remove any embedded commas from impressions
				wkimpres = impressions/1000
				projrev  = round((wkimpres * billrate), 2)
				if makegoodsw == False:
					actrev = projrev
				else:
					mgimp += impressions
					actrev = 0  # Set Make Good actual revenue to zero
					
				totprojrev += projrev
				totactrev  += actrev

				# write out formatted data
				
				xlt_writer.writerow([row[0], tgta, plcm, crtv, row[3], mgtrans,
									 billrate, projrev, actrev])
				outlinecnt += 1

print(' ')
print(f'Number of dates in file: {evtdatecnt}')
print(' ')
print(f'Processed {line_count} lines')
print(f'Lines written to reformat file: {outlinecnt}')

print(' ')
print(f'Make Good Transactions: {mgcnt}')
print(f'Make Good Impressions:  {mgimp}')

print(' ')
revdiff = totprojrev - totactrev
print(f'Projected revenue: {round(totprojrev, 2)}')
print(f'Actual revenue: {round(totactrev, 2)}')
print(f'Difference: {round(revdiff, 2)}')

if errsw == True:
	print(f'Errors found in data, count: {errcnt}')
		
rawdata_file.close()
xlt_file.close()

sys.exit()
