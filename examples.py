# python examples

import win32com.client as w32

def example(xl):
	# Create workbook
	wb = xl.Workbooks.Add()
	
	# Access / create a sheet
	# sheet = wb.Sheets('Sheet1')
	sheet = wb.ActiveSheet
	
	# Write to cells
	# NB: cell coordinates are NOT 0-based
	sheet.Cells(1,1).Value = 'Hacking excel with python'
	for i in range(2,8):
		sheet.Cells(i,1).Value = 'Line %i' % i
	for i in range(9,10):
		sheet.Cells(i,1).Value = i
	sheet.Cells(11,1).Formula = '=sum(a9:a10)'
	
	# Named ranges
	wb.Names.Add(Name = 'myname', RefersToR1C1 = '=Sheet1!R1C1:R2C1')
	for cell in sheet.Range('myname'):
		print(cell.Value)
