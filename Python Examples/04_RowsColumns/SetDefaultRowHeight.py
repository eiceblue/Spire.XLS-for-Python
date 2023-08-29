from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/CommonTemplate.xlsx"
outputFile = "SetDefaultRowHeight.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set default row height
sheet.DefaultRowHeight = 30
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

