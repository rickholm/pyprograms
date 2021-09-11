# Reformat ComCast Raw CSV Data to Zelle CSV Data Format
# 08-31-2021 - RAH - New
# 09-04-2021 - RAH - Changed to work in pycharm
# 09-05-2021 - RAH - Added support spreadsheet xlsx data input
# 09-14-2021 - RAH - Restructure program

######################################################################
#                                                                    #
# Function to scan and translate target audience                     #
#                                                                    #
######################################################################
def scantgtaud(placement:str):

    global tgtafr

    lusw   = False
    audidx = 0

    for wkaud in tgtafr:
        if re.search(wkaud, placement) != None:
            lusw = True
            break
        audidx += 1

    return lusw, audidx

######################################################################
#                                                                    #
# Function to scan creative for placement method and return index    #
#                                                                    #
######################################################################
def scanplcmeth(creative:str):

    global plcmfr

    lusw    = False
    methidx = 0
    for wkmeth in plcmfr:
        if re.search(wkmeth, creative) != None:
            lusw = True
            break
        methidx += 1

    return lusw, methidx


######################################################################
#                                                                    #
# Function to scan creative for translating                          #
#                                                                    #
######################################################################
def scancreative(creative:str):

    global crtvfr

    lusw   = False
    crtidx = 0
    for wkcrt in crtvfr:
        if re.search(wkcrt, creative) != None:
            lusw = True
            break
        crtidx += 1

    return lusw, crtidx

######################################################################
#                                                                    #
# Function to scan placement string for "make good" indicator        #
#                                                                    #
######################################################################
def scanmakegood(placement:str):

    global makegood

    for wkmg in makegood:
        if re.search(wkmg, placement) == None:
            foundsw = False
            break
        else:
            foundsw = True
            break

    return foundsw
# ********************************************************************

import os
import csv
import sys, getopt
import re
import openpyxl

progname = sys.argv[0]

etvid = 'EffecTV_APP_EffecTV_A35-54'  # Placement EffecTv Id

# Creative Tables

crtvfr = ['GameNight', 'Homework', 'NewDog']  # Scan For
crtvto = ['GameNight_Video_', 'HomeWork_Video_', 'NewDog_Video_']  # To Value

# Target Audience Tables

tgtafr = ['ChildFreeAdults', 'NestingFamilies', 'NewlyLiberated']  # Scan for
tgtato = ['Child Free Adults', 'Nesting Families', 'Newly Liberated']  # To Name
tgtaabv = ['CF', 'NF', 'NL']  # Abbreviations

# Placement Method Tables

plcmfr = ['IP_15', 'IP_30', 'STB_15', 'STB_30']
plcmto = ['IPBased_Video_15', 'IPBased_Video_30',
          'SetTop_Video_15', 'SetTop_Video_30']
plcmtm = ['15', '30', '15', '30']

makegood = ['_MG']

# Billing information

billrate = 38

# get input and output file names from command line_count

inputfile = ''
outputfile = ''
rplcfilesw = False

# Check for parameters

if not sys.argv[1:]:
    print ('No parameters found. Must supply file type (CSV or XLSX) and input/ouput file names')
    print(progname, ' -t <csv> or <xlsx> -i  <inputfile> -o <outputfile> -r')
    sys.exit(2)

try:
    opts, args = getopt.getopt(sys.argv[1:], 't:h:i:o:r')
except getopt.GetoptError as err:
    # print help information and exit:
    print(err)  # will print something like "option -a not recognized"
    print(progname, ' -t <csv> or <xlsx> -i  <inputfile> -o <outputfile> -r')
    sys.exit(2)
for opt, arg in opts:
    if opt == '-h':
        print(progname, '-t <csv or xlsx> -i <inputfile> -o <outputfile> -r')
        sys.exit()
    elif opt in ("-t"):
        filetype = arg
        wkfiletype = arg.upper()
        if (wkfiletype != 'CSV') and (wkfiletype != 'XLSX'):
            print('Invalid file type, must be csv or xlsx, found: ' + filetype)
            sys.exit(1)
    elif opt in ("-i"):
        inputfile = arg
    elif opt in ("-o"):
        outputfile = arg
    elif opt == '-r':
        rplcfilesw = True

print('Getting raw data from', inputfile, 'and reformatting data to', outputfile)

# Make sure input file exists and output file is new

if not os.path.isfile(inputfile):
    if wkfiletype == 'CSV':
        notfoundtext = 'Raw data file '
    else:
        notfoundtext = 'XLSX data file '
    print(notfoundtext + inputfile + ' not found...canceled')
    sys.exit(1)
if os.path.exists(outputfile) and rplcfilesw == False:
    print('Output file ' + outputfile + ' found...canceled')
    sys.exit(1)
if os.path.exists(outputfile) and rplcfilesw == True:
    os.remove(outputfile)

# loop to get raw data and reformat

rawdata_file = open(inputfile)
xlt_file = open(outputfile, mode='w', newline='')

raw_reader = csv.reader(rawdata_file, delimiter=',')

xlt_writer = csv.writer(xlt_file, delimiter=',')

firstsw = True
line_count = 0
outlinecnt = 0
savevtdate = ''
evtdatecnt = 0

totprojrev = 0
totactrev = 0

mgcnt = 0
mgimp = 0

errcnt = 0
errsw = False

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

        eventdate = str(row[0])

        # Look for target audience and translate

        wkaudience = str(row[2])
        foundsw    = False
        i          = 0
        foundsw, i = scantgtaud(wkaudience)
        if foundsw == True:
            tgta    = tgtato[i]
            plcmabv = tgtaabv[i]
        else:
            print('Data error in Target Audience', wkaudience)
            errcnt += 1
            errsw = True
            continue

        # Look for placement method: IP vs STB

        wkcreative = str(row[1])
        foundsw    = False
        i          = 0  # used to point the placement method table location
        foundsw, i = scanplcmeth(wkcreative)
        if foundsw == True:
            plcmmeth = plcmto[i]
            plcmtime = plcmtm[i]
        else:
            print('Data error in creative: placement method', wkcreative)
            errcnt += 1
            errsw = True
            continue

        plcm = etvid + plcmabv + '_' + plcmmeth + '_SS'

        # Translate creative

        wkcreative = str(row[1])
        foundsw    = False
        i          = 0  # used to point the creative translate
        foundsw, i = scancreative(wkcreative)
        if foundsw == True:
            crtv = crtvto[i] + plcmtime
        else:
            print('Data error with creative translate', wkcreative)
            errcnt += 1
            errsw = True
            continue

        # Compute billing information based on impressions per 1000

        impressions = int(re.sub(',', '', row[3]))  # Remove any embedded commas from impressions
        wkimpres = impressions / 1000
        projrev = round((wkimpres * billrate), 2)

        wkplacement = str(row[2])
        makegoodsw = scanmakegood(wkplacement)
        if makegoodsw == True:
            mgtrans = 'Yes'
            mgimp += impressions
            actrev = 0  # Set Make Good actual revenue to zero
            mgcnt += 1
        else:
            mgtrans = ' '
            actrev = projrev

        totprojrev += projrev
        totactrev += actrev

        # write out formatted data

        xlt_writer.writerow([eventdate, tgta, plcm, crtv, impressions, mgtrans,
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
