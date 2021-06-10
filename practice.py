from openpyxl import load_workbook
from openpyxl.styles import Font, Fill, colors, PatternFill, Protection

def print_all_worksheet_names(wbname):
	wb = load_workbook(wbname)

	print (wb.sheetnames)

	for wsheet in wb.sheetnames:
		print (wsheet)
	print (" ")

def print_row_column1(wbname,src_wname,dst_wname,filter_str):
	wb = load_workbook(wbname)

    #opening source worksheet in workbook
	swsheet = wb.get_sheet_by_name(src_wname)

    #Create destination worksheet, delete if exists

	try:
		dwsheet = wb.get_sheet_by_name(dst_wname)
		print ("worksheet '%s' found"%(dst_wname))
		print ("removing worksheet:'%s'"%(dst_wname))
		wb.remove_sheet(dwsheet)
	except:
		print ("worksheet '%s' not found"%(dst_wname))
	finally:
		print ("creating new worksheet:'%s'"%(dst_wname))
		dwsheet = wb.create_sheet(dst_wname,0) 

    #max rows and cols
	srcount = swsheet.max_row
	sccount = swsheet.max_column
	print ("%s: max row:col (%d:%d)"%(src_wname,srcount,sccount))

	drcount =  dwsheet.max_row
	dccount =  dwsheet.max_column
	print("%s: max row:col (%d:%d)" % (dst_wname, drcount, dccount))

    #copy titles
	row = swsheet[1]
	frow = [cell.value for cell in row]
	dwsheet.append(frow)
	

    #dump dsheet rows	
	for row in dwsheet.iter_rows():
		for cell in row:
			print (cell.value, " ",end=' ')
		print (" ")

    #copy all rows that matched with given string
	for row in swsheet.iter_rows():
		if (row[2].value != None and row[2].value == filter_str):
			frow = [cell.value for cell in row]
			print ("appending {0}".format(frow))
			dwsheet.append(frow)

	print("Saving :%s" % (wbname))
	wb.save(wbname)

def main():
	filename = "employee-details.xlsx"
	src_wname = "02-stories"
	dst_wname = "generated"
	filter_str = "balachandras"
	print_all_worksheet_names(filename)
	print_row_column1(filename,src_wname,dst_wname,filter_str)

if (__name__ == '__main__'):
	main()