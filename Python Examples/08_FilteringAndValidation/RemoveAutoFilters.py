from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/RemoveAutoFilters.xlsx"
outputFile = "RemoveAutoFilters.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Remove the auto filters.
sheet.AutoFilters.Clear()
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


