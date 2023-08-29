from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_1.xlsx"
outputFile = "ShrinkTextToFitInACell.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#The cell range to shrink text.
cell = sheet.Range["B13:C13"]
#Enable ShrinkToFit.
style = cell.Style
style.ShrinkToFit = True
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


