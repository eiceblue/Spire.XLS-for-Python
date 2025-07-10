from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ShowSubTotals.xlsx"
outputFile = "ShowSubTotals.xlsx"

#Create a workbook
workbook = Workbook()
#Load an Excel file including pivot table
workbook.LoadFromFile(inputFile)
#Get the sheet in which the pivot table is located
sheet = workbook.Worksheets["Pivot Table"]
pt = sheet.PivotTables[0]
#Show Subtotals
pt.ShowSubtotals = True
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
