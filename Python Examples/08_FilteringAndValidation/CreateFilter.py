from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/CreateFilter.xlsx"
outputFile = "CreateFilter.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Create filter
sheet.AutoFilters.Range = sheet.Range["A1:J1"]
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

