from os.path import join
import win32com.client


class ExportCharts (object):
	"""
	This will help in exporting charts created in excel file in png
	format. just pass the file name/path and enjoy !!!

	"""

	def __init__ (self,file,path):
		self.excel_file = file
		self.path = path
		self.workbook = join (self.path,self.excel_file)
		self.xlsApp = win32com.client.Dispatch("Excel.Application")
		self.xlsApp.Interactive = False
		self.xlsApp.Visible = False
		self.xlsWB = self.xlsApp.Workbooks.Open(self.workbook)


	def export_charts (self,sheet=None,path=None):
		path = path if path else self.path
		worksheets=sheet if sheet else [sheet.Name for sheet in self.xlsWB.Sheets]
		print (worksheets)
		for sheet in worksheets:
			xlsSheet = self.xlsWB.Sheets(sheet)
			chartObj = xlsSheet.ChartObjects()
			for index in range(1,chartObj.Count+1):
				mychart = chartObj(index)
				mychart.Copy
				#mychart.Chart.Export("E:\\MyCode\\Test_Projects\\chart" + str(index)+ "_"+str(sheet) + ".png")
				mychart.Chart.Export(join(path,mychart.Name+'_'+str(index)+".png"))

	def __del__(self):
		self.xlsWB.Save()
		self.xlsWB.Close()
		self.xlsWB = None
		self.xlsApp = None


class ExcelMacro(object):

	def __init__ (self,file,path):
		self.excel_file = file
		self.path = path
		self.workbook = join (self.path,self.excel_file)
		self.xlsApp = win32com.client.Dispatch("Excel.Application")
		self.xlsApp.Interactive = False
		self.xlsApp.Visible = False
		self.xlsWB = self.xlsApp.Workbooks.Open(self.workbook)
	
	def execute_excel_macro(self,macro):
		try:
			self.xlsApp.Run(macro)
			self.xlsWB.Save()
			print("Macro ran successfully!")
		except Exception as e:
			print("Error found while running the excel macro!")
			print (e)
			
	def __del__(self):
		self.xlsWB.Save()
		self.xlsWB.Close()
		self.xlsWB = None
		self.xlsApp = None


class ExcelWrite (object):

	def __init__ (self,file,path):
		self.excel_file = file
		self.path = path
		self.workbook = join (self.path,self.excel_file)
		self.xlsApp = win32com.client.Dispatch("Excel.Application")
		self.xlsApp.Interactive = False
		self.xlsApp.Visible = False
		self.xlsWB = self.xlsApp.Workbooks.Open(self.workbook)

	def write_data (self,sheet,data_list):
		xlsSheet = self.xlsWB.Sheets(sheet)
		row= 1
		for line in data_list:
			xlsSheet.Range(xlsSheet.Cells(row,1), xlsSheet.Cells(row, len(line))).Value = tuple(line)
			row += 1

	def __del__(self):
		self.xlsWB.Save()
		self.xlsWB.Close()
		self.xlsWB = None
		self.xlsApp = None



class ExcelReader (object):

	def __init__ (self,file,path):
		self.excel_file = file
		self.path = path
		self.workbook = join (self.path,self.excel_file) 
		self.xlsApp = win32com.client.Dispatch("Excel.Application")
		self.xlsApp.Interactive = False
		self.xlsApp.Visible = False
		self.xlsWB = self.xlsApp.Workbooks.Open(self.workbook)

	def get_all_data (self,sheet):
		xlsSheet = self.xlsWB.Sheets(sheet)
		return [list(row) for row in xlsSheet.UsedRange.Value]

	def Get_values_by_range(self,sheet,cell_range):
		xlsSheet = self.xlsWB.Sheets(sheet)
		return [list(v) for v in xlsSheet.Range(cell_range).Value]

	def __del__(self):
		self.xlsWB.Save()
		self.xlsWB.Close()
		self.xlsWB = None
		self.xlsApp = None




if __name__== '__main__':
	#obj=ExportCharts('file.xlsm','E:\MyCode\Test_Projects')
	print (ExportCharts.__doc__)
	#obj.exportCharts(path='E:\MyCode')
	#obj=CreateCharts('file.xlsx','E:\MyCode\Test_Projects')
	#obj.createChart ('Test','testing1', 'Sheet1')
	#obj=ExcelMacro('file.xlsm','E:\MyCode\Test_Projects')
	#obj.execute_excel_macro ('test_mac')
	#obj=ExcelReader('file.xlsx','E:\MyCode\Test_Projects')
	#print (obj.get_all_data ('Sheet1'))
	#obj1=ExcelWrite('file.xlsm','E:\MyCode\Test_Projects')
	#print (obj.Get_values_by_range ('Sheet1','A1:A6'))
	#obj1.write_data('Sheet1',data_list)




