from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_1.xlsx"
outputFile = "DetectMergedCells.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the merged cell ranges in the first worksheet and put them into a CellRange array.
range = sheet.MergedCells
#Traverse through the array and unmerge the merged cells.
for cell in range:
    cell.UnMerge()
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


