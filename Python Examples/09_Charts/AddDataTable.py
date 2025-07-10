from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/AddDataTable.xlsx"
outputFile = "AddDataTable.xlsx"

#Create a Workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the first chart
chart = sheet.Charts[0]
chart.HasDataTable = True
#Save and Launch
workbook.SaveToFile(outputFile, FileFormat.Version2010)
workbook.Dispose()

