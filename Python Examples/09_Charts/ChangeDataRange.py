from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SampeB_4.xlsx"
outputFile = "ChangeDataRange.xlsx"

#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Get chart
chart = sheet.Charts[0]
#Change data range
chart.DataRange = sheet.Range["A1:C4"]
#Save and launch result file 
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
