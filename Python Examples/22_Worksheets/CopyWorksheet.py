from spire.xls import *
from spire.xls.common import *


inputFile1 = "./Demos/Data/ReadImages.xlsx"
inputFile2 = "./Demos/Data/sample.xlsx"
outputFile = "CopyWorksheet.xlsx"

#Create a workbook and load a file
sourceWorkbook = Workbook()
sourceWorkbook.LoadFromFile(inputFile1)
#Get the first worksheet
srcWorksheet = sourceWorkbook.Worksheets[0]
#Create a workbook
targetWorkbook = Workbook()
#Load the target Excel document from disk
targetWorkbook.LoadFromFile(inputFile2)
#Add a new worksheet
targetWorksheet = targetWorkbook.Worksheets.Add("added")
#Copy the first worksheet of source Excel document to the new added worksheet of target Excel document
targetWorksheet.CopyFrom(srcWorksheet)
#Save the document
targetWorkbook.SaveToFile(outputFile, ExcelVersion.Version2013)
targetWorkbook.Dispose()

