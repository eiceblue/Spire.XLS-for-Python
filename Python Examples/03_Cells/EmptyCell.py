from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_1.xlsx"
outputFile = "EmptyCell.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Set the value as null to remove the original content from the Excel Cell.
sheet.Range["C6"].Value = ""
#Clear the contents to remove the original content from the Excel Cell.
sheet.Range["B6"].ClearContents()
#Remove the contents with format from the Excel cell.
sheet.Range["D6"].ClearAll()
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


