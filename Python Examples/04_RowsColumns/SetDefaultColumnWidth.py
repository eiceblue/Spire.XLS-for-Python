from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/CommonTemplate.xlsx"
outputFile = "SetDefaultColumnWidth.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set default column width
sheet.DefaultColumnWidth = 25
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()



