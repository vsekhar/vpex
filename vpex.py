import time
import sys

import win32com.client as w32

def init_excel(visible=False):
	xl = w32.gencache.EnsureDispatch('Excel.Application')
	xl.Visible = visible
	return xl

def end_excel(xl):
	xl.Application.Quit()

def autofit_all(ws):
	'Autofit all columns and rows in a worksheet'

	# can also do specific columns/rows using ws.Columns(1).AutoFit() etc.
	ws.Columns.AutoFit()
	ws.Rows.AutoFit()

def load_workbook(xl, filename):
	try:
		wb = xl.Workbooks.Open(filename)
	except:
		print("Failed to open%s" % filename)
		raise
	return wb

def get_data(ws):
	'''Grab all data into python (for processing without COM calls)
	
	Access by row tuples as xldata[0] == (c1, c2, c3, ...)'''
	
	xldata = sheet.UsedRange.Value
	return xldata


if __name__ == "__main__":
	if len(sys.argv) != 2:
		print('USAGE: python vpex.py {excel_workbook}')
		sys.exit(1)
	filename = sys.argv[1]
	xl = init_excel(visible=True)
	wb = load_workbook(xl, filename)
	print("Workbook '%s' loaded successfully" % filename)
	input()
	
	# do something