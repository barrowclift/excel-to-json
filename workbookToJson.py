#!/usr/local/bin/python3

"""
workbookToJson.py: Takes any number of Excel workbooks (either .xlsx or .xls)
and saves each of their respective worksheets as JSON files to the working
directory
"""

__author__ = "Marc Barrowclift"
__email__ = "marc@barrowclift.me"
__date__ = "September 3, 2015 1:33:32 PM"


import xlrd
import sys
import os
import string
import re
import simplejson as json
from collections import OrderedDict
from xlrd.sheet import ctype_text

TAB_WIDTH = 4*' '

workbooks = []
splitFiles = False
splitLimiter = 1000
prettify = False
currentTerminalSize = (0,0)


# -----------------------------------------------------------------------
# METHOD DEFINITIONS
# -----------------------------------------------------------------------

# PRINTING ============================

"""
Thanks to http://stackoverflow.com/users/836407/chown for the robust and
elegant solution.

SOLUTION: http://stackoverflow.com/a/566752 
"""
def getTerminalSize():
	import os
	env = os.environ
	def ioctl_GWINSZ(fd):
		try:
			import fcntl, termios, struct, os
			cr = struct.unpack('hh', fcntl.ioctl(fd, termios.TIOCGWINSZ,
		'1234'))
		except:
			return
		return cr
	cr = ioctl_GWINSZ(0) or ioctl_GWINSZ(1) or ioctl_GWINSZ(2)
	if not cr:
		try:
			fd = os.open(os.ctermid(), os.O_RDONLY)
			cr = ioctl_GWINSZ(fd)
			os.close(fd)
		except:
			pass
	if not cr:
		cr = (env.get('LINES', 25), env.get('COLUMNS', 80))

	return int(cr[1]), int(cr[0])

"""
Thanks to http://stackoverflow.com/users/4934338/jettico for the regex
solution to strip out all ANSII escape color codes for determining the "true"
length of any given string to print to console.

SOLUTION: http://stackoverflow.com/a/30500866
"""
def trimAnsii(a):
	ESC = r'\x1b'
	CSI = ESC + r'\['
	OSC = ESC + r'\]'
	CMD = '[@-~]'
	ST = ESC + r'\\'
	BEL = r'\x07'
	pattern = '(' + CSI + '.*?' + CMD + '|' + OSC + '.*?' + '(' + ST + '|' + BEL + ')' + ')'
	return re.sub(pattern, '', a)

"""
For printy-printing anything that involves indented blocks of text. It takes
into account the current width of the user's terminal and prints out the
perfectly fit and formatted text to match it's current size.

This method assumes that the global tuple variable currentTerminalSize is SET
and UPDATED for the current Terminal dimensions. This is to save us from
running all the code to calculate it's dimensions for every single call to
smartPrint (but at the expense of requiring the calling method to
appropriately manage it first thing themselves).
"""
def smartPrint(indent, message):
	maxWidth = currentTerminalSize[0]

	indentLength = len(indent)
	if indentLength > 0:
		sys.stdout.write(indent)
	availableWidth = indentLength

	messageWords = message.split(' ')
	for word in messageWords:
		# When determining the length of our current word, characters that are
		# escaped (our colors) will actually be counted in the length method
		# even if they are not actually printed to the screen, therefore we
		# need to filter it down to just the printed characters.
		printableCharacters = trimAnsii(word)
		if availableWidth == indentLength:
			wordLength = len(printableCharacters)
		else:
			wordLength = len(printableCharacters) + 1

		if availableWidth + wordLength > maxWidth:
			sys.stdout.write('\n%s%s' % (indent, word))
			availableWidth = indentLength + wordLength
		else:
			if availableWidth == indentLength:
				sys.stdout.write('%s' % word)
			else:
				sys.stdout.write(' %s' % word)

			availableWidth = availableWidth + wordLength
	
	if availableWidth != indentLength:
		print("")

def printUsage():
	currentTerminalSize = getTerminalSize()

	smartPrint("", "%sUSAGE%s: xmlxToJson.py [%sOPTIONS%s] ... [%sFILE.xlsx%s|%sFILE.xls%s] ..." % (colors.OKBLUE, colors.END, colors.HEADER, colors.END, colors.HEADER, colors.END, colors.HEADER, colors.END))
	smartPrint(TAB_WIDTH, "%s-p%s, %s--pretty%s" % (colors.OKBLUE, colors.END, colors.OKBLUE, colors.END))
	smartPrint(TAB_WIDTH*2, "Pretty-print the built JSON files (off by default to help save space, the generated JSON files are big)")
	print()
	smartPrint(TAB_WIDTH, "%s-s%s %sN%s, %s--split%s=%sN%s" % (colors.OKBLUE, colors.END, colors.HEADER, colors.END, colors.OKBLUE, colors.END, colors.HEADER, colors.END))
	smartPrint(TAB_WIDTH*2, "Split up the generated JSON into multiple files. For example, an N of 1000 would start a new JSON file when the 1000th root child element was written.")
	print()
	smartPrint(TAB_WIDTH*2, "The generated files will have the naming convention FILENAME_1.json, FILENAME_2.json, etc.")
	print()
	smartPrint(TAB_WIDTH, "%s-h%s, %s--help%s" % (colors.OKBLUE, colors.END, colors.OKBLUE, colors.END))
	smartPrint(TAB_WIDTH*2, "Print the usage documentation")

class colors:
	HEADER = '\033[95m'
	OKBLUE = '\033[1;94m'
	OKGREEN = '\033[92m'
	WARNING = '\033[1;93m'
	FAIL = '\033[91m'
	END = '\033[0m'
	BOLD = '\033[1m'
	UNDERLINE = '\033[4m'

def printError(message):
	print("%sERROR: %s%s" % (colors.FAIL, message, colors.END))

def printWarning(message):
	print("%sWARNING: %s%s" % (colors.WARNING, message, colors.END))

def printSuccess(message):
	print("%s%s%s" % (colors.OKGREEN, message, colors.END))

# PROCESS METHODS ============================

def parseArguments(argv):
	global prettify, splitFiles, splitLimiter, workbooks

	arguments = iter(argv)
	next(arguments)
	awaitingSplitFilesValue = False
	for argument in arguments:
		if awaitingSplitFilesValue:
			splitLimiter = int(argument)
			awaitingSplitFilesValue = False
		elif argument == "-p" or argument == "--pretty":
			prettify = True
		elif argument == "-s" or argument == "--split":
			splitFiles = True
			awaitingSplitFilesValue = True
		elif argument.startswith("--split="):
			tokens = argument.split('=')
			if len(tokens) != 2:
				printError("Could not parse split number argument '%s'\n" % argument)
				printUsage()
				sys.exit(1)
			else:
				splitFiles = True
				splitLimiter = int(tokens[1])
		elif argument == '-h' or argument == "--help":
			printUsage()
			sys.exit(0)
		elif os.path.isfile(argument): # Assume any other arguments given are workbooks to process
			workbooks.append(argument)
		else:
			printError("Unknown argument '%s' encountered\n" % argument)
			printUsage()
			sys.exit(1)

	if not workbooks:
		print("No Excel workbooks were given to process, please try again while supplying your .xls or .xlsx file(s) to parse")
		return False
	else:
		return True

def getColumnNamesListFromSheet(sheet):
	columnNamesRow = sheet.row(0)
	columnNames = []
	for index, columnName in enumerate(columnNamesRow):
		cellDataType = ctype_text.get(columnName.ctype, 'unknown type')
		if cellDataType != "text":
			print("ERROR: Only 'text' data types for column names is currently supported");
			sys.exit(1)
		
		columnNames.append(columnName.value)
	return columnNames

def getOrderedDictListOfAllRows(sheet):
	columnNames = getColumnNamesListFromSheet(sheet)

	rows = []
	for rowIndex in range(1, sheet.nrows):
		row = OrderedDict()
		rowValues = sheet.row_values(rowIndex)
		for columnIndex, columnName in enumerate(columnNames):
			row[columnName] = rowValues[columnIndex]
		rows.append(row)
	return rows

def makeJsonFromList(rows):
	if prettify:
		return json.dumps(rows, indent=TAB_WIDTH, ensure_ascii=False, encoding="utf-8")
	else:
		return json.dumps(rows, ensure_ascii=False, encoding="utf-8")

def writeRowsAsJsonToFile(rows, filePath):
	json = makeJsonFromList(rows)
	writeJsonToFile(json, filePath)

def writeJsonToFile(json, filePath):
	with open(filePath, 'w') as sheetFile:
		sheetFile.write(json)


# -----------------------------------------------------------------------
# MAIN PROGRAM
# -----------------------------------------------------------------------

"""
Argument validation first
"""
currentTerminalSize = getTerminalSize()
if len(sys.argv) <= 1:
	print("No Excel workbooks were given to process, please try again while supplying your .xls or .xlsx file(s) to parse")
	sys.exit(0)
else:
	# Skipping the program's name (always the first argument)
	success = parseArguments(sys.argv)
	if not success:
		printUsage()
		sys.exit(0)

"""
Making JSON files
"""
# Parsing all workbooks and their respective sheets into json files 
for workbook in workbooks:
	book = xlrd.open_workbook(workbook)
	if (book.nsheets > 1):
		print("Parsing workbook '%s' with %d worksheets" % (workbook, book.nsheets))
	else:
		print("Parsing workbook '%s' with %d worksheet" % (workbook, book.nsheets))

	for sheetIndex in range(book.nsheets):
		sheet = book.sheet_by_index(sheetIndex)
		smartPrint(TAB_WIDTH, "Processing worksheet '%s'..." % sheet.name)

		rows = getOrderedDictListOfAllRows(sheet)
		if splitFiles:
			currentSplitFileNumber = 1
			currentRow = 1
			bufferedRows = []
			for row in rows:
				bufferedRows.append(row)
				currentRow += 1
				if currentRow == splitLimiter:
					fileName = sheet.name+"-"+str(currentSplitFileNumber)+".json"
					writeRowsAsJsonToFile(bufferedRows, fileName)

					# Resetting for the next JSON file part
					bufferedRows = []
					currentSplitFileNumber += 1
					currentRow = 1
		else:
			fileName = sheet.name+".json"
			writeRowsAsJsonToFile(rows, fileName)

printSuccess("Completed all workbook to JSON conversions")